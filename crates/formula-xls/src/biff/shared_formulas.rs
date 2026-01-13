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
                out.warnings.push(format!("malformed BIFF record: {err}"));
                break;
            }
        };

        if record.offset != start && records::is_bof_record(record.record_id) {
            break;
        }

        match record.record_id {
            RECORD_SHRFMLA => match parse_shrfmla_record(record.data.as_ref()) {
                Some(def) => out.shared_formulas.push(def),
                None => out.warnings.push(format!(
                    "failed to parse SHRFMLA record at offset {}",
                    record.offset
                )),
            },
            records::RECORD_EOF => break,
            _ => {}
        }
    }

    Ok(out)
}

fn parse_shrfmla_record(data: &[u8]) -> Option<SharedFormulaDef> {
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
        end: usize,
    }

    fn parse_refu(data: &[u8]) -> Option<RangeHeader> {
        if data.len() < 6 {
            return None;
        }
        let row_first = u16::from_le_bytes([data[0], data[1]]);
        let row_last = u16::from_le_bytes([data[2], data[3]]);
        let col_first = data[4] as u16;
        let col_last = data[5] as u16;
        Some(RangeHeader {
            row_first,
            row_last,
            col_first,
            col_last,
            end: 6,
        })
    }

    fn parse_ref8(data: &[u8]) -> Option<RangeHeader> {
        if data.len() < 8 {
            return None;
        }
        let row_first = u16::from_le_bytes([data[0], data[1]]);
        let row_last = u16::from_le_bytes([data[2], data[3]]);
        let col_first = u16::from_le_bytes([data[4], data[5]]) & 0x3FFF;
        let col_last = u16::from_le_bytes([data[6], data[7]]) & 0x3FFF;
        Some(RangeHeader {
            row_first,
            row_last,
            col_first,
            col_last,
            end: 8,
        })
    }

    let headers = [parse_refu(data), parse_ref8(data)];

    for header in headers.into_iter().flatten() {
        if header.row_first > header.row_last || header.col_first > header.col_last {
            continue;
        }

        // In BIFF8, SHRFMLA usually includes a 2-byte `cUse` field before `cce`. Prefer that layout
        // first so we don't misinterpret `cUse=0` as `cce=0` and return an empty rgce.
        for prefix in [2usize, 0, 4] {
            let cce_offset = match header.end.checked_add(prefix) {
                Some(v) => v,
                None => continue,
            };
            if data.len() < cce_offset + 2 {
                continue;
            }
            let cce = u16::from_le_bytes([data[cce_offset], data[cce_offset + 1]]) as usize;
            let rgce_start = cce_offset + 2;
            let rgce_end = match rgce_start.checked_add(cce) {
                Some(v) => v,
                None => continue,
            };
            if rgce_end > data.len() {
                continue;
            }
            let rgce = data.get(rgce_start..rgce_end)?.to_vec();
            return Some(SharedFormulaDef {
                row_first: header.row_first,
                row_last: header.row_last,
                col_first: header.col_first,
                col_last: header.col_last,
                rgce,
            });
        }
    }

    None
}
