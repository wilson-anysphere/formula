//! Best-effort parsing of BIFF8 shared formulas (`SHRFMLA` records) from worksheet substreams.
//!
//! In BIFF8, shared-formula groups are represented by:
//! - A `SHRFMLA` record that stores the shared rgce token stream and the cell range it applies to.
//! - `FORMULA` records in cells within the range that often contain only `PtgExp` pointing at the
//!   shared formula base cell.
//!
//! Some writers appear to omit the base cell's full formula token stream (or even omit the base
//! cell's `FORMULA` record entirely) and rely solely on `SHRFMLA` + `PtgExp`. Calamine may fail to
//! surface such formulas, so the `.xls` importer uses this parser as a fallback to recover formula
//! text for affected cells.

use super::records;

/// SHRFMLA [MS-XLS 2.4.255]
const RECORD_SHRFMLA: u16 = 0x04BC;

/// Cap warnings collected by best-effort shared-formula scanning so a crafted `.xls` cannot
/// allocate an unbounded number of warning strings.
const MAX_WARNINGS_PER_SHEET: usize = 50;
const WARNINGS_SUPPRESSED_MESSAGE: &str = "additional warnings suppressed";

fn push_warning_bounded(warnings: &mut Vec<String>, warning: impl Into<String>) {
    if warnings.len() < MAX_WARNINGS_PER_SHEET {
        warnings.push(warning.into());
        return;
    }
    // Add a single terminal warning so callers have a hint that the import was noisy.
    if warnings.len() == MAX_WARNINGS_PER_SHEET {
        warnings.push(WARNINGS_SUPPRESSED_MESSAGE.to_string());
    }
}

#[derive(Debug, Clone)]
pub(crate) struct SharedFormulaDef {
    pub(crate) row_first: u16,
    pub(crate) row_last: u16,
    pub(crate) col_first: u16,
    pub(crate) col_last: u16,
    /// Raw BIFF8 rgce token stream stored in the SHRFMLA record.
    pub(crate) rgce: Vec<u8>,
}

#[derive(Debug, Default)]
pub(crate) struct SheetSharedFormulas {
    pub(crate) shared_formulas: Vec<SharedFormulaDef>,
    pub(crate) warnings: Vec<String>,
}

pub(crate) fn parse_biff_sheet_shared_formulas(
    workbook_stream: &[u8],
    start: usize,
) -> Result<SheetSharedFormulas, String> {
    let mut out = SheetSharedFormulas::default();

    // SHRFMLA records can be large (shared formula rgce) and may be split across CONTINUE records.
    // Use the logical iterator so we reassemble the record payload before parsing.
    let allows_continuation = |record_id: u16| record_id == RECORD_SHRFMLA;
    let iter = records::LogicalBiffRecordIter::from_offset(workbook_stream, start, allows_continuation)?;

    for record in iter {
        let record = match record {
            Ok(record) => record,
            Err(err) => {
                push_warning_bounded(&mut out.warnings, format!("malformed BIFF record: {err}"));
                break;
            }
        };

        if record.offset != start && records::is_bof_record(record.record_id) {
            break;
        }

        match record.record_id {
            RECORD_SHRFMLA => match parse_shrfmla_record(&record) {
                Some(def) => out.shared_formulas.push(def),
                None => push_warning_bounded(
                    &mut out.warnings,
                    format!("failed to parse SHRFMLA record at offset {}", record.offset),
                ),
            },
            records::RECORD_EOF => break,
            _ => {}
        }
    }

    Ok(out)
}

fn parse_shrfmla_record(record: &records::LogicalBiffRecord<'_>) -> Option<SharedFormulaDef> {
    // SHRFMLA [MS-XLS 2.4.255]
    //
    // Producers in the wild appear to vary slightly in their record layout. Try a small set of
    // plausible header shapes:
    // - RefU (6 bytes): [rwFirst: u16][rwLast: u16][colFirst: u8][colLast: u8]
    // - Ref8 (8 bytes): [rwFirst: u16][rwLast: u16][colFirst: u16][colLast: u16]
    //
    // After the range header, accept an optional small prefix (0/2/4 bytes) followed by:
    //   [cce: u16][rgce: cce bytes]
    //
    // Any trailing bytes are ignored.

    #[derive(Clone, Copy)]
    struct RangeHeader {
        row_first: u16,
        row_last: u16,
        col_first: u16,
        col_last: u16,
    }

    fn parse_refu_range(data: &[u8]) -> Option<RangeHeader> {
        let chunk = data.get(0..6)?;
        let row_first = u16::from_le_bytes([chunk[0], chunk[1]]);
        let row_last = u16::from_le_bytes([chunk[2], chunk[3]]);
        let col_first = chunk[4] as u16;
        let col_last = chunk[5] as u16;
        Some(RangeHeader {
            row_first,
            row_last,
            col_first,
            col_last,
        })
    }

    fn parse_ref8_range(data: &[u8]) -> Option<RangeHeader> {
        let chunk = data.get(0..8)?;
        let row_first = u16::from_le_bytes([chunk[0], chunk[1]]);
        let row_last = u16::from_le_bytes([chunk[2], chunk[3]]);
        let col_first = u16::from_le_bytes([chunk[4], chunk[5]]) & 0x3FFF;
        let col_last = u16::from_le_bytes([chunk[6], chunk[7]]) & 0x3FFF;
        Some(RangeHeader {
            row_first,
            row_last,
            col_first,
            col_last,
        })
    }

    let parsed = super::worksheet_formulas::parse_biff8_shrfmla_record(record).ok()?;
    if parsed.rgce.is_empty() {
        return None;
    }

    // Try to identify the correct SHRFMLA header layout by matching the `cce` value stored in the
    // record body against the length of the parsed rgce stream. This avoids misidentifying Ref8
    // headers (which use u16 column fields) as RefU when the shared range starts in column A.
    let data = record.data.as_ref();
    let expected_cce = parsed.rgce.len();
    let header = {
        let mut out: Option<RangeHeader> = None;
        let set_if_matches =
            |header: Option<RangeHeader>, cce_offset: usize, data: &[u8], out: &mut Option<RangeHeader>| {
                let Some(header) = header else {
                    return;
                };
                if header.row_first > header.row_last || header.col_first > header.col_last {
                    return;
                }
                let cce_bytes = match data.get(cce_offset..cce_offset + 2) {
                    Some(v) => v,
                    None => return,
                };
                let cce = u16::from_le_bytes([cce_bytes[0], cce_bytes[1]]) as usize;
                if cce == expected_cce {
                    *out = Some(header);
                }
            };

        // Match the parsing order used by `worksheet_formulas::parse_biff8_shrfmla_record`.
        // Layout A: RefU (6) + cUse (2) + cce (2).
        set_if_matches(parse_refu_range(data), 8, data, &mut out);
        // Layout B: Ref8 (8) + cUse (2) + cce (2).
        if out.is_none() {
            set_if_matches(parse_ref8_range(data), 10, data, &mut out);
        }
        // Layout C: RefU (6) + cce (2) (cUse omitted).
        if out.is_none() {
            set_if_matches(parse_refu_range(data), 6, data, &mut out);
        }
        // Layout D: Ref8 (8) + cce (2) (cUse omitted).
        if out.is_none() {
            set_if_matches(parse_ref8_range(data), 8, data, &mut out);
        }

        out
    }
    // Fallback to the old heuristic if we fail to match a header layout.
    .or_else(|| {
        [parse_refu_range(data), parse_ref8_range(data)]
            .into_iter()
            .flatten()
            .find(|header| header.row_first <= header.row_last && header.col_first <= header.col_last)
    })?;

    Some(SharedFormulaDef {
        row_first: header.row_first,
        row_last: header.row_last,
        col_first: header.col_first,
        col_last: header.col_last,
        rgce: parsed.rgce,
    })
}
