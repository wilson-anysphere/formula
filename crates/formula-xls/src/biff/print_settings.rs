use formula_model::{ManualPageBreaks, Orientation, PageSetup, Scaling};

use super::records;

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
const RECORD_WSBOOL: u16 = 0x0081;

// SETUP grbit flags.
//
// In BIFF8, SETUP.grbit bit 1 is `fPortrait`:
// - 0 => landscape
// - 1 => portrait
const SETUP_GRBIT_PORTRAIT: u16 = 0x0002;

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
    let mut setup_scale: Option<u16> = None;
    let mut setup_fit_width: Option<u16> = None;
    let mut setup_fit_height: Option<u16> = None;
    let mut wsbool_fit_to_page: Option<bool> = None;

    let mut iter = records::BiffRecordIter::from_offset(workbook_stream, start)?;

    while let Some(next) = iter.next() {
        let record = match next {
            Ok(r) => r,
            Err(err) => {
                out.warnings.push(format!("malformed BIFF record: {err}"));
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
                    out.warnings.push(format!(
                        "truncated WSBOOL record at offset {} (len={}, expected >=2)",
                        record.offset,
                        data.len()
                    ));
                    continue;
                }
                let grbit = u16::from_le_bytes([data[0], data[1]]);
                wsbool_fit_to_page = Some((grbit & 0x0100) != 0);
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
    if saw_any_record || fit_to_page {
        if fit_to_page {
            if let (Some(width), Some(height)) = (setup_fit_width, setup_fit_height) {
                page_setup.scaling = Scaling::FitTo { width, height };
            } else {
                // Some `.xls` writers omit SETUP even when fit-to-page is enabled; preserve
                // the mode as best-effort.
                page_setup.scaling = Scaling::FitTo {
                    width: 0,
                    height: 0,
                };
            }
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
            out.warnings.extend(breaks.warnings.drain(..));
        }
        Err(err) => out
            .warnings
            .push(format!("failed to parse manual page breaks: {err}")),
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
        warnings.push(format!(
            "truncated SETUP record at offset {offset} (len={}, expected 34)",
            data.len()
        ));
    }

    let paper_size = parse_u16_at(data, 0);
    let scale = parse_u16_at(data, 2);
    let fit_width = parse_u16_at(data, 6);
    let fit_height = parse_u16_at(data, 8);
    let grbit = parse_u16_at(data, 10);
    let header_margin = parse_f64_at(data, 16);
    let footer_margin = parse_f64_at(data, 24);

    if let Some(code) = paper_size {
        // BIFF8 uses `iPaperSize==0` and values >=256 for printer-specific/custom paper sizes.
        // These values do not map cleanly onto OpenXML `ST_PaperSize` numeric codes and are not
        // representable in the model. Ignore them and keep the default paper size.
        if code == 0 || code >= 256 {
            warnings.push(format!(
                "ignoring custom/invalid paper size code {code} in SETUP record at offset {offset}"
            ));
        } else {
            page_setup.paper_size.code = code;
        }
    }

    if let Some(grbit) = grbit {
        page_setup.orientation = if (grbit & SETUP_GRBIT_PORTRAIT) != 0 {
            Orientation::Portrait
        } else {
            Orientation::Landscape
        };
    }

    if let Some(v) = header_margin {
        if v.is_finite() {
            page_setup.margins.header = v;
        } else {
            warnings.push(format!(
                "invalid SETUP header margin value {v:?} at offset {offset}"
            ));
        }
    }
    if let Some(v) = footer_margin {
        if v.is_finite() {
            page_setup.margins.footer = v;
        } else {
            warnings.push(format!(
                "invalid SETUP footer margin value {v:?} at offset {offset}"
            ));
        }
    }

    (scale, fit_width, fit_height)
}

fn parse_margin_record(
    out: &mut f64,
    name: &'static str,
    data: &[u8],
    offset: usize,
    warnings: &mut Vec<String>,
) {
    if data.len() < 8 {
        warnings.push(format!(
            "truncated {name} record at offset {offset} (len={}, expected 8)",
            data.len()
        ));
        return;
    }
    let value = parse_f64_at(data, 0).expect("len check");
    if !value.is_finite() {
        warnings.push(format!("invalid {name} value {value:?} at offset {offset}"));
        return;
    }
    *out = value;
}
