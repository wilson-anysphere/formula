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

use super::{records, worksheet_formulas};

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
    /// Trailing data blocks (`rgcb`) referenced by certain ptgs (notably `PtgArray`).
    pub(crate) rgcb: Vec<u8>,
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
    // The `SHRFMLA.rgce` token stream can be split across `CONTINUE` records. When a `PtgStr`
    // (ShortXLUnicodeString) payload crosses a continuation boundary, Excel inserts an extra 1-byte
    // "continued segment option flags" prefix at the start of the continued fragment.
    //
    // Naively concatenating record fragments therefore corrupts the rgce byte stream. Use the
    // fragment-aware parser in `worksheet_formulas` so those continuation flag bytes are skipped.
    //
    // After the range header, accept an optional small prefix (0/2/4 bytes) followed by:
    //   [cce: u16][rgce: cce bytes]
    //
    // Any trailing bytes after `rgce` are treated as `rgcb` (data blocks referenced by ptgs like
    // `PtgArray`).

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

    let parsed = worksheet_formulas::parse_biff8_shrfmla_record(record).ok()?;
    if parsed.rgce.is_empty() {
        return None;
    }

    let data = record.data.as_ref();
    let expected_cce = parsed.rgce.len();

    // Try to identify the correct SHRFMLA header layout by matching the `cce` value stored in the
    // record body against the length of the parsed rgce stream. This avoids misidentifying Ref8
    // headers (which use u16 column fields) as RefU when the shared range starts in column A.
    //
    // Note: When `cUse` is omitted, `Ref8 + cce` (Layout D) and `RefU + cUse + cce` (Layout A) both
    // place `cce` at offset 8. To avoid truncating Ref8 ranges to column A, we collect all matching
    // candidates and apply a small heuristic: if a `cUse` field is present and non-zero, it is
    // expected to match the range area (number of cells) in well-formed sheets.
    let header = {
        #[derive(Clone, Copy)]
        struct Candidate {
            header: RangeHeader,
            uses_ref8: bool,
            cuse: Option<u16>,
        }

        fn valid_range(h: RangeHeader) -> bool {
            h.row_first <= h.row_last && h.col_first <= h.col_last
        }

        fn range_area(h: RangeHeader) -> u64 {
            let rows = (h.row_last.saturating_sub(h.row_first) as u64).saturating_add(1);
            let cols = (h.col_last.saturating_sub(h.col_first) as u64).saturating_add(1);
            rows.saturating_mul(cols)
        }

        let expected_cce_u16 = u16::try_from(expected_cce).ok()?;

        let mut candidates: Vec<Candidate> = Vec::new();
        let push_candidate =
            |header: Option<RangeHeader>,
             uses_ref8: bool,
             cuse: Option<u16>,
             cce_offset: usize,
             candidates: &mut Vec<Candidate>| {
                let Some(header) = header else {
                    return;
                };
                if !valid_range(header) {
                    return;
                }
                let cce_bytes = match data.get(cce_offset..cce_offset + 2) {
                    Some(v) => v,
                    None => return,
                };
                let cce = u16::from_le_bytes([cce_bytes[0], cce_bytes[1]]);
                if cce == expected_cce_u16 {
                    candidates.push(Candidate {
                        header,
                        uses_ref8,
                        cuse,
                    });
                }
            };

        // Layout A: RefU (6) + cUse (2) + cce (2).
        let cuse_a = data
            .get(6..8)
            .map(|v| u16::from_le_bytes([v[0], v[1]]));
        push_candidate(parse_refu_range(data), false, cuse_a, 8, &mut candidates);

        // Layout B: Ref8 (8) + cUse (2) + cce (2).
        let cuse_b = data
            .get(8..10)
            .map(|v| u16::from_le_bytes([v[0], v[1]]));
        push_candidate(parse_ref8_range(data), true, cuse_b, 10, &mut candidates);

        // Layout C: RefU (6) + cce (2) (cUse omitted).
        push_candidate(parse_refu_range(data), false, None, 6, &mut candidates);

        // Layout D: Ref8 (8) + cce (2) (cUse omitted).
        push_candidate(parse_ref8_range(data), true, None, 8, &mut candidates);

        if candidates.is_empty() {
            None
        } else {
            candidates
                .into_iter()
                .min_by_key(|c| {
                    let area = range_area(c.header);
                    let cuse_rank: u8 = match c.cuse {
                        Some(cuse) if cuse != 0 && (cuse as u64) == area => 0,
                        Some(0) => 1,
                        None => 1,
                        Some(_) => 2,
                    };
                    (
                        cuse_rank,
                        // Prefer Ref8 headers when other signals are equal: they are more general
                        // and avoid dropping columns when a producer uses Ref8 unnecessarily.
                        if c.uses_ref8 { 0u8 } else { 1u8 },
                        area,
                        c.header.row_first,
                        c.header.row_last,
                        c.header.col_first,
                        c.header.col_last,
                    )
                })
                .map(|c| c.header)
        }
    }
    // Fallback: accept the first parseable header when we cannot match `cce`.
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
        rgcb: parsed.rgcb,
    })
}
