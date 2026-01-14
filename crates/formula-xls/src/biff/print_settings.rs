use formula_model::{ManualPageBreaks, Orientation, PageSetup, Scaling};

use super::records;

/// Cap warnings collected by best-effort worksheet print-settings scans.
///
/// These scans run on untrusted input (legacy BIFF streams) and should not allow a crafted file to
/// allocate an unbounded number of warning strings.
const MAX_WARNINGS_PER_SHEET_PRINT_SETTINGS: usize = 50;
const PRINT_SETTINGS_WARNINGS_SUPPRESSED_MESSAGE: &str =
    "additional print settings warnings suppressed";

/// Hard cap on the number of BIFF records scanned while searching for print settings.
///
/// The `.xls` importer performs multiple best-effort passes over each worksheet substream (print
/// settings, hyperlinks, view state, etc.). Without a cap, a crafted workbook with millions of cell
/// records can force excessive work even when a particular feature is absent.
#[cfg(not(test))]
const MAX_RECORDS_SCANNED_PER_SHEET_PRINT_SETTINGS_SCAN: usize = 500_000;
// Keep unit tests fast by using a smaller cap.
#[cfg(test)]
const MAX_RECORDS_SCANNED_PER_SHEET_PRINT_SETTINGS_SCAN: usize = 1_000;

fn push_warning_bounded(warnings: &mut Vec<String>, warning: impl Into<String>) {
    if warnings.len() < MAX_WARNINGS_PER_SHEET_PRINT_SETTINGS {
        warnings.push(warning.into());
        return;
    }
    if warnings.len() == MAX_WARNINGS_PER_SHEET_PRINT_SETTINGS {
        warnings.push(PRINT_SETTINGS_WARNINGS_SUPPRESSED_MESSAGE.to_string());
    }
}

/// Push a warning but ensure it is present even if the warning buffer is already full.
///
/// This is used for "critical" warnings (e.g. hardening caps) where we want to surface the
/// condition even if earlier best-effort parsing already exhausted the warning budget.
fn push_warning_bounded_force(warnings: &mut Vec<String>, warning: impl Into<String>) {
    let warning = warning.into();

    if warnings.len() < MAX_WARNINGS_PER_SHEET_PRINT_SETTINGS {
        warnings.push(warning);
        return;
    }

    // Keep the warning buffer size bounded. Prefer preserving the terminal suppression marker (when
    // present) and replace the last "real" warning.
    let replace_idx = if warnings.len() == MAX_WARNINGS_PER_SHEET_PRINT_SETTINGS + 1
        && warnings
            .last()
            .is_some_and(|w| w == PRINT_SETTINGS_WARNINGS_SUPPRESSED_MESSAGE)
    {
        MAX_WARNINGS_PER_SHEET_PRINT_SETTINGS.saturating_sub(1)
    } else {
        warnings.len().saturating_sub(1)
    };

    if let Some(slot) = warnings.get_mut(replace_idx) {
        *slot = warning;
    } else {
        // Should be unreachable, but fall back to the bounded helper for safety.
        push_warning_bounded(warnings, warning);
    }
}
// Excel constrains page margin values (in inches) to the range 0 <= x < 49.
// Treat 49 as inclusive for simplicity.
const MARGIN_MIN_INCHES: f64 = 0.0;
const MARGIN_MAX_INCHES: f64 = 49.0;

// Worksheet print/page setup related record ids.
// See [MS-XLS] sections:
// - SETUP: 2.4.257
// - LEFTMARGIN/RIGHTMARGIN/TOPMARGIN/BOTTOMMARGIN: 2.4.132/2.4.214/2.4.326/2.4.38
// - HORIZONTALPAGEBREAKS/VERTICALPAGEBREAKS: 2.4.122/2.4.355
// - WSBOOL: 2.4.376
const RECORD_SETUP: u16 = 0x00A1;
const RECORD_LEFTMARGIN: u16 = 0x0026;
const RECORD_RIGHTMARGIN: u16 = 0x0027;
const RECORD_TOPMARGIN: u16 = 0x0028;
const RECORD_BOTTOMMARGIN: u16 = 0x0029;
// WSBOOL [MS-XLS 2.4.376] stores worksheet boolean properties; we only care about fFitToPage
// (bit 8 / mask 0x0100).
const RECORD_WSBOOL: u16 = 0x0081;
// WSBOOL options ([MS-XLS] 2.4.376).
//
// In BIFF8, `WSBOOL.fFitToPage` controls whether SETUP.iFitWidth/iFitHeight apply. When unset,
// Excel uses SETUP.iScale percent scaling instead.
const WSBOOL_OPTION_FIT_TO_PAGE: u16 = 0x0100;
// SETUP grbit flags.
//
// In BIFF8, SETUP.grbit bit 1 is `fPortrait`:
// - 0 => landscape
// - 1 => portrait
const SETUP_GRBIT_PORTRAIT: u16 = 0x0002;
// If set, printer-related fields in SETUP (including iPaperSize/iScale/iRes/iVRes/iCopies/fNoOrient/fPortrait)
// are undefined and must be ignored. See [MS-XLS] 2.4.257 (SETUP), `fNoPls`.
const SETUP_GRBIT_F_NOPLS: u16 = 0x0004;
// If set, `fPortrait` must be ignored and orientation defaults to portrait.
// See [MS-XLS] 2.4.257 (SETUP), `fNoOrient`.
const SETUP_GRBIT_F_NOORIENT: u16 = 0x0040;
const SETUP_MAX_FIT_DIMENSION: u16 = 32767;

#[derive(Debug, Clone)]
pub(crate) struct BiffSheetPrintSettings {
    pub(crate) page_setup: Option<PageSetup>,
    pub(crate) manual_page_breaks: ManualPageBreaks,
    pub(crate) warnings: Vec<String>,
}

impl Default for BiffSheetPrintSettings {
    fn default() -> Self {
        Self {
            page_setup: None,
            manual_page_breaks: ManualPageBreaks::default(),
            warnings: Vec::new(),
        }
    }
}

/// Best-effort parse of worksheet print settings (page setup + margins + manual page breaks).
///
/// This scan is resilient to malformed records: payload-level parse failures are surfaced as
/// warnings and otherwise ignored.
pub(crate) fn parse_biff_sheet_print_settings(
    workbook_stream: &[u8],
    start: usize,
) -> Result<BiffSheetPrintSettings, String> {
    parse_biff_sheet_print_settings_with_record_cap(
        workbook_stream,
        start,
        MAX_RECORDS_SCANNED_PER_SHEET_PRINT_SETTINGS_SCAN,
    )
}

fn parse_biff_sheet_print_settings_with_record_cap(
    workbook_stream: &[u8],
    start: usize,
    record_cap: usize,
) -> Result<BiffSheetPrintSettings, String> {
    let mut out = BiffSheetPrintSettings::default();

    let mut page_setup = PageSetup::default();
    let mut saw_any_record = false;
    // Scaling in BIFF8 uses two orthogonal signals:
    // - SETUP stores both iScale and iFitWidth/iFitHeight
    // - WSBOOL.fFitToPage indicates whether iFit* are active; otherwise use iScale.
    //
    // Keep these raw values until the end so WSBOOL order doesn't matter. Some `.xls` producers
    // omit SETUP even when fit-to-page is enabled; in that case we preserve the intent as
    // `FitTo { width: 0, height: 0 }`.
    let mut wsbool_fit_to_page: Option<bool> = None;
    let mut setup_scale: Option<u16> = None;
    let mut setup_fit_width: Option<u16> = None;
    let mut setup_fit_height: Option<u16> = None;
    // Track whether the last SETUP record had `fNoPls=1` so we can avoid inferring FitTo mode from
    // `iFitWidth`/`iFitHeight` when WSBOOL is missing. Per [MS-XLS], `fNoPls` makes various
    // printer-related fields undefined, which can otherwise result in false positives.
    let mut setup_no_pls: bool = false;

    let mut iter = records::BiffRecordIter::from_offset(workbook_stream, start)?;
    let mut scanned = 0usize;

    while let Some(next) = iter.next() {
        let record = match next {
            Ok(r) => r,
            Err(err) => {
                push_warning_bounded(&mut out.warnings, format!("malformed BIFF record: {err}"));
                break;
            }
        };

        if record.offset != start && records::is_bof_record(record.record_id) {
            break;
        }

        scanned = scanned.saturating_add(1);
        if scanned > record_cap {
            push_warning_bounded_force(
                &mut out.warnings,
                format!(
                    "too many BIFF records while scanning sheet print settings (cap={record_cap}); stopping early"
                ),
            );
            break;
        }

        let data = record.data;
        match record.record_id {
            RECORD_SETUP => {
                saw_any_record = true;
                let (scale, fit_width, fit_height, no_pls) =
                    parse_setup_record(&mut page_setup, data, record.offset, &mut out.warnings);
                setup_no_pls = no_pls;
                // Best-effort: preserve the most recent values when later SETUP records are
                // truncated/malformed (so corrupt files don't clobber earlier state).
                //
                // `fNoPls=1` indicates the printer-related fields (including iScale) are undefined.
                // Treat this as a "clear" so later scaling logic falls back to defaults.
                if no_pls {
                    setup_scale = None;
                } else if let Some(value) = scale {
                    setup_scale = Some(value);
                }
                if let Some(value) = fit_width {
                    setup_fit_width = Some(value);
                }
                if let Some(value) = fit_height {
                    setup_fit_height = Some(value);
                }
            }
            RECORD_LEFTMARGIN => parse_margin_record(
                &mut page_setup.margins.left,
                "LEFTMARGIN",
                data,
                record.offset,
                &mut out.warnings,
            ),
            RECORD_RIGHTMARGIN => parse_margin_record(
                &mut page_setup.margins.right,
                "RIGHTMARGIN",
                data,
                record.offset,
                &mut out.warnings,
            ),
            RECORD_TOPMARGIN => parse_margin_record(
                &mut page_setup.margins.top,
                "TOPMARGIN",
                data,
                record.offset,
                &mut out.warnings,
            ),
            RECORD_BOTTOMMARGIN => parse_margin_record(
                &mut page_setup.margins.bottom,
                "BOTTOMMARGIN",
                data,
                record.offset,
                &mut out.warnings,
            ),
            RECORD_WSBOOL => {
                // WSBOOL [MS-XLS 2.4.376]
                // fFitToPage: bit8 (mask 0x0100).
                if data.len() < 2 {
                    push_warning_bounded(
                        &mut out.warnings,
                        format!(
                            "truncated WSBOOL record at offset {} (expected >=2 bytes, got {})",
                            record.offset,
                            data.len()
                        ),
                    );
                    continue;
                }
                let grbit = u16::from_le_bytes([data[0], data[1]]);
                let fit_to_page = (grbit & WSBOOL_OPTION_FIT_TO_PAGE) != 0;
                wsbool_fit_to_page = Some(fit_to_page);
                if fit_to_page {
                    // Treat fit-to-page as a meaningful print setting even when SETUP is missing.
                    saw_any_record = true;
                }
            }
            records::RECORD_EOF => break,
            _ => {}
        }

        if matches!(
            record.record_id,
            RECORD_LEFTMARGIN | RECORD_RIGHTMARGIN | RECORD_TOPMARGIN | RECORD_BOTTOMMARGIN
        ) {
            saw_any_record = true;
        }
    }

    let fit_to_page = wsbool_fit_to_page.unwrap_or_else(|| {
        // Heuristic: infer FitTo mode when WSBOOL is missing but the SETUP record contains
        // non-default fit dimensions. Disable this inference when `fNoPls=1` so we don't treat
        // undefined printer state as a FitTo signal.
        if setup_no_pls {
            false
        } else {
            setup_fit_width.unwrap_or(0) != 0 || setup_fit_height.unwrap_or(0) != 0
        }
    });

    if saw_any_record {
        page_setup.scaling = if fit_to_page {
            // Some `.xls` writers omit the SETUP record even when fit-to-page is enabled.
            // Preserve the scaling intent: FitTo {0,0} means "as many pages as needed".
            Scaling::FitTo {
                width: setup_fit_width.unwrap_or(0),
                height: setup_fit_height.unwrap_or(0),
            }
        } else {
            let scale = setup_scale.unwrap_or(100);
            Scaling::Percent(if scale == 0 { 100 } else { scale })
        };
        out.page_setup = Some(page_setup);
    }

    // Manual page breaks are stored in dedicated worksheet records. Delegate to the existing
    // page-break parser so we share the same semantics as the rest of the importer.
    match super::sheet::parse_biff_sheet_manual_page_breaks(workbook_stream, start) {
        Ok(mut breaks) => {
            out.manual_page_breaks = breaks.manual_page_breaks;
            for warning in breaks.warnings.drain(..) {
                push_warning_bounded_force(&mut out.warnings, warning);
            }
        }
        Err(err) => push_warning_bounded_force(
            &mut out.warnings,
            format!("failed to parse manual page breaks: {err}"),
        ),
    }

    Ok(out)
}

fn parse_u16_at(data: &[u8], offset: usize) -> Option<u16> {
    let bytes = data.get(offset..offset + 2)?;
    Some(u16::from_le_bytes([bytes[0], bytes[1]]))
}

fn parse_f64_at(data: &[u8], offset: usize) -> Option<f64> {
    let bytes = data.get(offset..offset + 8)?;
    Some(f64::from_le_bytes([
        bytes[0], bytes[1], bytes[2], bytes[3], bytes[4], bytes[5], bytes[6], bytes[7],
    ]))
}

fn parse_setup_record(
    page_setup: &mut PageSetup,
    data: &[u8],
    offset: usize,
    warnings: &mut Vec<String>,
) -> (Option<u16>, Option<u16>, Option<u16>, bool) {
    // BIFF8 SETUP record is 34 bytes:
    // [iPaperSize:u16][iScale:u16][iPageStart:u16][iFitWidth:u16][iFitHeight:u16]
    // [grbit:u16][iRes:u16][iVRes:u16][numHdr:f64][numFtr:f64][iCopies:u16]
    //
    // We parse fields opportunistically to stay best-effort on truncated records.
    if data.len() < 34 {
        push_warning_bounded(
            warnings,
            format!(
                "truncated SETUP record at offset {offset} (len={}, expected 34)",
                data.len()
            ),
        );
    }

    let paper_size = parse_u16_at(data, 0);
    let scale = parse_u16_at(data, 2);
    let mut fit_width = parse_u16_at(data, 6);
    let mut fit_height = parse_u16_at(data, 8);
    let grbit = parse_u16_at(data, 10);
    let header_margin = parse_f64_at(data, 16);
    let footer_margin = parse_f64_at(data, 24);

    let (f_no_pls, f_no_orient, f_portrait) = match grbit {
        Some(grbit) => (
            (grbit & SETUP_GRBIT_F_NOPLS) != 0,
            (grbit & SETUP_GRBIT_F_NOORIENT) != 0,
            (grbit & SETUP_GRBIT_PORTRAIT) != 0,
        ),
        None => (false, false, false),
    };

    // Excel clamps fit dimensions to 32767 (the maximum value representable in the UI).
    if let Some(w) = fit_width {
        if w > SETUP_MAX_FIT_DIMENSION {
            push_warning_bounded(
                warnings,
                format!(
                    "invalid SETUP.iFitWidth value {w} at offset {offset}: must be <= {SETUP_MAX_FIT_DIMENSION}; clamped to {SETUP_MAX_FIT_DIMENSION}"
                ),
            );
            fit_width = Some(SETUP_MAX_FIT_DIMENSION);
        }
    }
    if let Some(h) = fit_height {
        if h > SETUP_MAX_FIT_DIMENSION {
            push_warning_bounded(
                warnings,
                format!(
                    "invalid SETUP.iFitHeight value {h} at offset {offset}: must be <= {SETUP_MAX_FIT_DIMENSION}; clamped to {SETUP_MAX_FIT_DIMENSION}"
                ),
            );
            fit_height = Some(SETUP_MAX_FIT_DIMENSION);
        }
    }

    // Printer fields (including iScale) are undefined when fNoPls is set.
    let scale_out = if f_no_pls { None } else { scale };

    // Per [MS-XLS], when `fNoPls` is set the printer-related fields are undefined and must be
    // ignored. The spec does not list `iFitWidth`/`iFitHeight` as undefined, and Excel will honor
    // them when WSBOOL.fFitToPage is set, so we always preserve the fit dimensions.
    if f_no_pls {
        // Reset any previously imported printer settings so "last wins" behaves as if the printer
        // fields were absent.
        page_setup.paper_size = PageSetup::default().paper_size;
        page_setup.orientation = Orientation::Portrait;
    } else {
        if let Some(code) = paper_size {
            // BIFF8 uses `iPaperSize==0` and values >=256 for printer-specific/custom paper sizes.
            // These values do not map cleanly onto OpenXML `ST_PaperSize` numeric codes and are not
            // representable in the model. Ignore them and keep the default paper size.
            if code == 0 || code >= 256 {
                push_warning_bounded(
                    warnings,
                    format!(
                        "ignoring custom/invalid paper size code {code} in SETUP record at offset {offset}"
                    ),
                );
            } else {
                page_setup.paper_size.code = code;
            }
        }

        if grbit.is_some() {
            page_setup.orientation = if f_no_orient || f_portrait {
                Orientation::Portrait
            } else {
                Orientation::Landscape
            };
        }
    }
    if let Some(v) = header_margin {
        if is_valid_margin(v) {
            page_setup.margins.header = v;
        } else {
            push_warning_bounded(
                warnings,
                format!("invalid SETUP header margin (numHdr) value {v:?} at offset {offset}"),
            );
        }
    }
    if let Some(v) = footer_margin {
        if is_valid_margin(v) {
            page_setup.margins.footer = v;
        } else {
            push_warning_bounded(
                warnings,
                format!("invalid SETUP footer margin (numFtr) value {v:?} at offset {offset}"),
            );
        }
    }

    (scale_out, fit_width, fit_height, f_no_pls)
}

fn parse_margin_record(
    out: &mut f64,
    name: &'static str,
    data: &[u8],
    offset: usize,
    warnings: &mut Vec<String>,
) {
    if data.len() < 8 {
        push_warning_bounded(
            warnings,
            format!(
                "truncated {name} record at offset {offset} (len={}, expected 8)",
                data.len()
            ),
        );
        return;
    }
    let value = parse_f64_at(data, 0).expect("len check");
    if !is_valid_margin(value) {
        push_warning_bounded(
            warnings,
            format!("invalid {name} value {value:?} at offset {offset}"),
        );
        return;
    }
    *out = value;
}

fn is_valid_margin(value: f64) -> bool {
    value.is_finite() && value >= MARGIN_MIN_INCHES && value <= MARGIN_MAX_INCHES
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record(id: u16, data: &[u8]) -> Vec<u8> {
        let mut out = Vec::with_capacity(4 + data.len());
        out.extend_from_slice(&id.to_le_bytes());
        out.extend_from_slice(&(data.len() as u16).to_le_bytes());
        out.extend_from_slice(data);
        out
    }

    fn setup_payload(
        i_paper_size: u16,
        i_scale: u16,
        i_fit_width: u16,
        i_fit_height: u16,
        grbit: u16,
        num_hdr: f64,
        num_ftr: f64,
    ) -> Vec<u8> {
        // BIFF8 SETUP record payload.
        let mut out = Vec::new();
        out.extend_from_slice(&i_paper_size.to_le_bytes());
        out.extend_from_slice(&i_scale.to_le_bytes());
        out.extend_from_slice(&0u16.to_le_bytes()); // iPageStart
        out.extend_from_slice(&i_fit_width.to_le_bytes());
        out.extend_from_slice(&i_fit_height.to_le_bytes());
        out.extend_from_slice(&grbit.to_le_bytes());
        out.extend_from_slice(&0u16.to_le_bytes()); // iRes
        out.extend_from_slice(&0u16.to_le_bytes()); // iVRes
        out.extend_from_slice(&num_hdr.to_le_bytes());
        out.extend_from_slice(&num_ftr.to_le_bytes());
        out.extend_from_slice(&1u16.to_le_bytes()); // iCopies
        out
    }

    #[test]
    fn print_settings_parser_is_reexported_from_print_settings_module() {
        let reexport = super::super::parse_biff_sheet_print_settings as usize;
        let direct = parse_biff_sheet_print_settings as usize;
        assert_eq!(
            reexport, direct,
            "biff::parse_biff_sheet_print_settings should re-export biff::print_settings::parse_biff_sheet_print_settings"
        );
    }

    #[test]
    fn print_settings_warnings_are_bounded_and_preserve_page_break_cap_warning() {
        // Fill the print-settings warning buffer with many truncated WSBOOL records.
        let mut stream = Vec::new();
        for _ in 0..(MAX_WARNINGS_PER_SHEET_PRINT_SETTINGS + 25) {
            stream.extend_from_slice(&record(RECORD_WSBOOL, &[]));
        }

        // Add a malformed HorizontalPageBreaks record that triggers a `cbrk` cap warning.
        let mut page_breaks = Vec::new();
        page_breaks.extend_from_slice(&u16::MAX.to_le_bytes()); // cbrk
        page_breaks.extend_from_slice(&2u16.to_le_bytes()); // row
        page_breaks.extend_from_slice(&0u16.to_le_bytes()); // colStart
        page_breaks.extend_from_slice(&0u16.to_le_bytes()); // colEnd
        stream.extend_from_slice(&record(0x001B, &page_breaks));
        stream.extend_from_slice(&record(records::RECORD_EOF, &[]));

        let parsed = parse_biff_sheet_print_settings(&stream, 0).expect("parse");

        assert!(
            parsed.warnings.len() <= MAX_WARNINGS_PER_SHEET_PRINT_SETTINGS + 1,
            "warnings should be bounded, got len={} warnings={:?}",
            parsed.warnings.len(),
            parsed.warnings
        );
        assert!(
            parsed
                .warnings
                .iter()
                .any(|w| w.contains("HorizontalPageBreaks") && w.contains("cbrk=")),
            "expected page-break cap warning to be preserved, got {:?}",
            parsed.warnings
        );
        assert_eq!(
            parsed.warnings.last().map(String::as_str),
            Some(PRINT_SETTINGS_WARNINGS_SUPPRESSED_MESSAGE),
            "suppression marker should be preserved, got {:?}",
            parsed.warnings
        );
        assert!(
            parsed.manual_page_breaks.row_breaks_after.contains(&1),
            "expected page break after row 1, got {:?}",
            parsed.manual_page_breaks.row_breaks_after
        );
    }

    #[test]
    fn parses_page_setup_margins_and_fit_to_page_scaling() {
        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(
                RECORD_SETUP,
                &setup_payload(
                    9,      // A4
                    77,     // iScale (ignored when fit-to-page)
                    2,      // iFitWidth
                    3,      // iFitHeight
                    0x0000, // landscape (SETUP.grbit.fPortrait=0)
                    0.5,    // header inches
                    0.6,    // footer inches
                ),
            ),
            record(RECORD_LEFTMARGIN, &1.0f64.to_le_bytes()),
            record(RECORD_LEFTMARGIN, &2.0f64.to_le_bytes()), // last wins
            record(RECORD_RIGHTMARGIN, &1.2f64.to_le_bytes()),
            record(RECORD_TOPMARGIN, &1.3f64.to_le_bytes()),
            record(RECORD_BOTTOMMARGIN, &1.4f64.to_le_bytes()),
            record(RECORD_WSBOOL, &0x0100u16.to_le_bytes()), // fFitToPage=1
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff_sheet_print_settings(&stream, 0).expect("parse");
        let setup = parsed
            .page_setup
            .as_ref()
            .expect("expected page setup from SETUP record");
        assert_eq!(setup.paper_size.code, 9);
        assert_eq!(setup.orientation, Orientation::Landscape);
        assert_eq!(setup.scaling, Scaling::FitTo { width: 2, height: 3 });
        assert_eq!(setup.margins.left, 2.0);
        assert_eq!(setup.margins.right, 1.2);
        assert_eq!(setup.margins.top, 1.3);
        assert_eq!(setup.margins.bottom, 1.4);
        assert_eq!(setup.margins.header, 0.5);
        assert_eq!(setup.margins.footer, 0.6);
        assert!(
            parsed.warnings.is_empty(),
            "expected no warnings, got {:?}",
            parsed.warnings
        );
    }

    #[test]
    fn parses_percent_scaling_when_fit_to_page_disabled() {
        let grbit = 0x0000u16; // landscape (SETUP.grbit.fPortrait=0)
        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(
                RECORD_SETUP,
                &setup_payload(1, 80, 1, 1, grbit, 0.3, 0.3), // iScale=80%
            ),
            record(RECORD_WSBOOL, &0u16.to_le_bytes()), // fFitToPage=0
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff_sheet_print_settings(&stream, 0).expect("parse");
        let setup = parsed
            .page_setup
            .as_ref()
            .expect("expected page setup from SETUP record");
        assert_eq!(setup.scaling, Scaling::Percent(80));
    }

    #[test]
    fn setup_f_nopls_ignores_printer_fields() {
        // fNoPls=1 => iPaperSize/iScale/fLandscape are undefined and must be ignored.
        let grbit = 0x0004u16; // fNoPls
        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(
                RECORD_SETUP,
                &setup_payload(9, 80, 1, 1, grbit, 0.4, 0.5), // values ignored except header/footer
            ),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff_sheet_print_settings(&stream, 0).expect("parse");
        let setup = parsed
            .page_setup
            .as_ref()
            .expect("expected page setup from SETUP record");
        assert_eq!(setup.paper_size, PageSetup::default().paper_size);
        assert_eq!(setup.orientation, Orientation::Portrait);
        assert_eq!(setup.scaling, Scaling::Percent(100));
        assert_eq!(setup.margins.header, 0.4);
        assert_eq!(setup.margins.footer, 0.5);
    }

    #[test]
    fn warns_on_truncated_margin_records_and_continues() {
        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_TOPMARGIN, &[0xAA, 0xBB]), // truncated
            record(RECORD_LEFTMARGIN, &1.0f64.to_le_bytes()),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff_sheet_print_settings(&stream, 0).expect("parse");
        let setup = parsed
            .page_setup
            .as_ref()
            .expect("expected page setup from margin record");
        assert_eq!(setup.margins.left, 1.0);
        assert!(
            parsed
                .warnings
                .iter()
                .any(|w| w.contains("truncated TOPMARGIN record")),
            "expected truncated-TOPMARGIN warning, got {:?}",
            parsed.warnings
        );
    }

    #[test]
    fn print_settings_scan_stops_after_record_cap() {
        let record_cap = 10usize;

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            // Exceed the record-scan cap with junk records.
            (0..(record_cap + 10))
                .flat_map(|_| record(0x1234, &[]))
                .collect::<Vec<u8>>(),
            // This SETUP record should be ignored because we stop scanning early.
            record(
                RECORD_SETUP,
                &setup_payload(1, 80, 1, 1, 0x0000, 0.3, 0.3),
            ),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff_sheet_print_settings_with_record_cap(&stream, 0, record_cap)
            .expect("parse");
        assert!(parsed.page_setup.is_none());
        assert!(
            parsed
                .warnings
                .iter()
                .any(|w| w.contains("too many BIFF records") && w.contains("print settings")),
            "expected record-cap warning, got {:?}",
            parsed.warnings
        );
    }

    #[test]
    fn print_settings_record_cap_warning_is_emitted_even_when_other_warnings_are_suppressed() {
        let record_cap = MAX_WARNINGS_PER_SHEET_PRINT_SETTINGS + 20;

        let mut stream: Vec<u8> = Vec::new();
        stream.extend_from_slice(&record(records::RECORD_BOF_BIFF8, &[0u8; 16]));

        // Fill the print-settings warning buffer with many truncated WSBOOL records.
        for _ in 0..(record_cap + 10) {
            stream.extend_from_slice(&record(RECORD_WSBOOL, &[]));
        }
        stream.extend_from_slice(&record(records::RECORD_EOF, &[]));

        let parsed = parse_biff_sheet_print_settings_with_record_cap(&stream, 0, record_cap)
            .expect("parse");
        assert_eq!(
            parsed.warnings.len(),
            MAX_WARNINGS_PER_SHEET_PRINT_SETTINGS + 1,
            "warnings should remain capped; warnings={:?}",
            parsed.warnings
        );
        assert!(
            parsed
                .warnings
                .iter()
                .any(|w| w.contains("too many BIFF records") && w.contains("print settings")),
            "expected forced record-cap warning, got {:?}",
            parsed.warnings
        );
        assert_eq!(
            parsed.warnings.last().map(String::as_str),
            Some(PRINT_SETTINGS_WARNINGS_SUPPRESSED_MESSAGE),
            "suppression marker should remain last; warnings={:?}",
            parsed.warnings
        );
    }
}
