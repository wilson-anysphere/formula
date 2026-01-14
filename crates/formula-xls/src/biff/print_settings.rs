use formula_model::{ManualPageBreaks, Orientation, PageSetup, Scaling};

use super::records;

/// Cap warnings collected by best-effort worksheet print-settings scans.
///
/// These scans run on untrusted input (legacy BIFF streams) and should not allow a crafted file to
/// allocate an unbounded number of warning strings.
const MAX_WARNINGS_PER_SHEET_PRINT_SETTINGS: usize = 50;
const PRINT_SETTINGS_WARNINGS_SUPPRESSED_MESSAGE: &str =
    "additional print settings warnings suppressed";

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
// When set, printer settings are undefined (ignore paper size, scale, orientation, etc).
const SETUP_GRBIT_F_NOPLS: u16 = 0x0004;
// When set, the `fPortrait` bit must be ignored and orientation defaults to portrait.
const SETUP_GRBIT_F_NOORIENT: u16 = 0x0040;

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
    let mut out = BiffSheetPrintSettings::default();

    let mut page_setup = PageSetup::default();
    let mut saw_any_record = false;
    // WSBOOL.fFitToPage controls whether SETUP's iFitWidth/iFitHeight apply.
    // Keep the raw SETUP scaling fields around and compute scaling at the end so record order
    // doesn't matter and "last wins" semantics are respected.
    let mut wsbool_fit_to_page: Option<bool> = None;
    let mut setup_scale: Option<u16> = None;
    let mut setup_fit_width: Option<u16> = None;
    let mut setup_fit_height: Option<u16> = None;

    let mut iter = records::BiffRecordIter::from_offset(workbook_stream, start)?;

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

        let data = record.data;
        match record.record_id {
            RECORD_SETUP => {
                saw_any_record = true;
                let (scale, fit_width, fit_height) =
                    parse_setup_record(&mut page_setup, data, record.offset, &mut out.warnings);
                setup_scale = scale;
                setup_fit_width = fit_width;
                setup_fit_height = fit_height;
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
                // fFitToPage: bit8 (mask 0x0100)
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
                wsbool_fit_to_page = Some((grbit & WSBOOL_OPTION_FIT_TO_PAGE) != 0);
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
        setup_fit_width.unwrap_or(0) != 0 || setup_fit_height.unwrap_or(0) != 0
    });

    // WSBOOL.fFitToPage can imply non-default scaling even when SETUP is missing. Treat it as a
    // signal that print settings exist so downstream import paths can preserve FitTo mode.
    if saw_any_record || fit_to_page {
        if fit_to_page {
            page_setup.scaling = Scaling::FitTo {
                width: setup_fit_width.unwrap_or(0),
                height: setup_fit_height.unwrap_or(0),
            };
        } else {
            let scale = setup_scale.unwrap_or(100);
            page_setup.scaling = Scaling::Percent(if scale == 0 { 100 } else { scale });
        }
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
) -> (Option<u16>, Option<u16>, Option<u16>) {
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

    let mut scale_out = scale;
    let (f_no_pls, f_no_orient, f_portrait) = match grbit {
        Some(grbit) => (
            (grbit & SETUP_GRBIT_F_NOPLS) != 0,
            (grbit & SETUP_GRBIT_F_NOORIENT) != 0,
            (grbit & SETUP_GRBIT_PORTRAIT) != 0,
        ),
        None => (false, false, false),
    };

    if f_no_pls {
        // Per spec: iPaperSize/iScale/fPortrait/iRes/iVRes/iCopies are undefined; ignore them.
        // Also reset any previously seen printer-related values so "last wins" behaves as if
        // printer settings were not present.
        page_setup.paper_size = PageSetup::default().paper_size;
        page_setup.orientation = Orientation::Portrait;
        scale_out = None;
        // Best-effort: treat fit-to-page fields as undefined as well so we do not infer FitTo mode
        // from non-zero iFitWidth/iFitHeight when WSBOOL is absent.
        fit_width = None;
        fit_height = None;
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
        if v.is_finite() {
            page_setup.margins.header = v;
        } else {
            push_warning_bounded(
                warnings,
                format!("invalid SETUP header margin value {v:?} at offset {offset}"),
            );
        }
    }
    if let Some(v) = footer_margin {
        if v.is_finite() {
            page_setup.margins.footer = v;
        } else {
            push_warning_bounded(
                warnings,
                format!("invalid SETUP footer margin value {v:?} at offset {offset}"),
            );
        }
    }

    (scale_out, fit_width, fit_height)
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
    if !value.is_finite() {
        push_warning_bounded(
            warnings,
            format!("invalid {name} value {value:?} at offset {offset}"),
        );
        return;
    }
    *out = value;
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
}
