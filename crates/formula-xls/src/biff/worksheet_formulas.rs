//! BIFF8 worksheet formula record parsing helpers.
//!
//! BIFF8 worksheet formulas are stored as `rgce` token streams inside `FORMULA`, `SHRFMLA`, and
//! `ARRAY` records. These records can be split across `CONTINUE` boundaries.
//!
//! When a `PtgStr` (ShortXLUnicodeString) payload is continued into a `CONTINUE` record, Excel
//! inserts an extra 1-byte "continued segment" option flags prefix at the fragment boundary.
//! Naively concatenating record payload bytes therefore corrupts the rgce stream (token alignment),
//! typically producing string literals containing an embedded NUL and leaving trailing bytes.
//!
//! This module implements a fragment-aware `rgce` reader that tokenizes the stream and skips those
//! continuation flag bytes so downstream formula decoding sees the canonical rgce bytes.

#![allow(dead_code)]

use std::collections::{HashMap, HashSet};

use formula_model::{CellRef, EXCEL_MAX_COLS, EXCEL_MAX_ROWS};

use super::{records, rgce};

// Worksheet record ids (BIFF8).
// See [MS-XLS]:
// - FORMULA: 2.4.127 (0x0006)
// - ARRAY: 2.4.19 (0x0221)
// - SHRFMLA: 2.4.276 (0x04BC)
// - TABLE: 2.4.328 (0x0236)
pub(crate) const RECORD_FORMULA: u16 = 0x0006;
pub(crate) const RECORD_ARRAY: u16 = 0x0221;
pub(crate) const RECORD_SHRFMLA: u16 = 0x04BC;
pub(crate) const RECORD_TABLE: u16 = 0x0236;

/// BIFF8 `FORMULA.grbit` bitfield.
///
/// We only decode the subset needed to disambiguate `PtgExp`/`PtgTbl` resolution.
///
/// [MS-XLS] 2.4.127 (FORMULA) specifies these relevant bits:
/// - `0x0008` (`fShrFmla`): the formula is part of a shared formula group (expects `SHRFMLA` + `PtgExp`)
/// - `0x0010` (`fArray`): the formula is part of an array formula (expects `ARRAY` + `PtgExp`)
/// - `0x0020` (`fTbl`): the formula is part of a data table (expects `TABLE` + `PtgTbl`)
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct FormulaGrbit(pub(crate) u16);

impl FormulaGrbit {
    pub(crate) const F_SHR_FMLA: u16 = 0x0008;
    pub(crate) const F_ARRAY: u16 = 0x0010;
    pub(crate) const F_TBL: u16 = 0x0020;

    pub(crate) fn is_shared(self) -> bool {
        (self.0 & Self::F_SHR_FMLA) != 0
    }

    pub(crate) fn is_array(self) -> bool {
        (self.0 & Self::F_ARRAY) != 0
    }

    pub(crate) fn is_table(self) -> bool {
        (self.0 & Self::F_TBL) != 0
    }

    /// Returns a single, unambiguous membership hint if exactly one of `fShrFmla`, `fArray`, or
    /// `fTbl` is set.
    pub(crate) fn membership_hint(self) -> Option<FormulaMembershipHint> {
        let mut out: Option<FormulaMembershipHint> = None;
        let mut count = 0usize;
        if self.is_shared() {
            out = Some(FormulaMembershipHint::Shared);
            count += 1;
        }
        if self.is_array() {
            out = Some(FormulaMembershipHint::Array);
            count += 1;
        }
        if self.is_table() {
            out = Some(FormulaMembershipHint::Table);
            count += 1;
        }

        if count == 1 {
            out
        } else {
            None
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub(crate) enum FormulaMembershipHint {
    Shared,
    Array,
    Table,
}

// BIFF8 string option flags used by ShortXLUnicodeString.
// See [MS-XLS] 2.5.293.
const STR_FLAG_HIGH_BYTE: u8 = 0x01;
const STR_FLAG_EXT: u8 = 0x04;
const STR_FLAG_RICH_TEXT: u8 = 0x08;

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct ParsedFormulaRecord {
    pub(crate) row: u16,
    pub(crate) col: u16,
    pub(crate) xf: u16,
    pub(crate) grbit: FormulaGrbit,
    pub(crate) rgce: Vec<u8>,
    /// Trailing data blocks (`rgcb`) referenced by certain ptgs (notably `PtgArray`).
    pub(crate) rgcb: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct ParsedSharedFormulaRecord {
    pub(crate) rgce: Vec<u8>,
    pub(crate) rgcb: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct ParsedArrayRecord {
    pub(crate) rgce: Vec<u8>,
    pub(crate) rgcb: Vec<u8>,
}

pub(crate) fn parse_biff8_formula_record(
    record: &records::LogicalBiffRecord<'_>,
) -> Result<ParsedFormulaRecord, String> {
    let fragments: Vec<&[u8]> = record.fragments().collect();
    let mut cursor = FragmentCursor::new(&fragments, 0, 0);

    // FORMULA [MS-XLS 2.4.127]
    let row = cursor.read_u16_le()?;
    let col = cursor.read_u16_le()?;
    let xf = cursor.read_u16_le()?;

    // Skip cached result (8), read flags/grbit (2), skip calc chain (4).
    cursor.skip_bytes(8)?;
    let grbit = FormulaGrbit(cursor.read_u16_le()?);
    cursor.skip_bytes(4)?;

    let cce = cursor.read_u16_le()? as usize;
    let rgce = cursor.read_biff8_rgce(cce)?;
    let rgcb = cursor.read_remaining_bytes()?;

    Ok(ParsedFormulaRecord {
        row,
        col,
        xf,
        grbit,
        rgce,
        rgcb,
    })
}

pub(crate) fn parse_biff8_shrfmla_record(
    record: &records::LogicalBiffRecord<'_>,
) -> Result<ParsedSharedFormulaRecord, String> {
    let fragments: Vec<&[u8]> = record.fragments().collect();
    let cursor = FragmentCursor::new(&fragments, 0, 0);

    // SHRFMLA layouts vary slightly between producers (RefU vs Ref8 for the shared range). Try a
    // small set of plausible BIFF8 layouts.
    //
    // Most writers follow [MS-XLS] and include the `cUse` field, but some emit a shorter layout
    // (RefU/Ref8 + cce + rgce) without it. We treat both as best-effort compatible.
    //
    // Layout A: RefU (6) + cUse (2) + cce (2).
    let mut c = cursor.clone();
    if let Ok((rgce, rgcb)) = parse_shrfmla_with_refu(&mut c) {
        if !rgce.is_empty() {
            return Ok(ParsedSharedFormulaRecord { rgce, rgcb });
        }
    }
    // Layout B: Ref8 (8) + cUse (2) + cce (2).
    let mut c = cursor.clone();
    if let Ok((rgce, rgcb)) = parse_shrfmla_with_ref8(&mut c) {
        if !rgce.is_empty() {
            return Ok(ParsedSharedFormulaRecord { rgce, rgcb });
        }
    }
    // Layout C: RefU (6) + cce (2) (cUse omitted).
    let mut c = cursor.clone();
    if let Ok((rgce, rgcb)) = parse_shrfmla_with_refu_no_cuse(&mut c) {
        if !rgce.is_empty() {
            return Ok(ParsedSharedFormulaRecord { rgce, rgcb });
        }
    }
    // Layout D: Ref8 (8) + cce (2) (cUse omitted).
    let mut c = cursor;
    if let Ok((rgce, rgcb)) = parse_shrfmla_with_ref8_no_cuse(&mut c) {
        if !rgce.is_empty() {
            return Ok(ParsedSharedFormulaRecord { rgce, rgcb });
        }
    }

    Err("unrecognized SHRFMLA record layout".to_string())
}

/// Parse a BIFF8 `SHRFMLA` record into its `rgce` token stream plus any trailing `rgcb` bytes.
///
/// `rgcb` stores payloads for certain tokens (notably `PtgArray`) that are not embedded directly
/// in the `rgce` token stream.
///
/// This parser is fragment-aware: if the `rgce` stream contains a continued `PtgStr`
/// (ShortXLUnicodeString) whose character payload crosses a `CONTINUE` boundary, it skips the
/// extra 1-byte continued-segment option flags prefix inserted at the boundary so the returned
/// `rgce` bytes match the canonical stream.
pub(crate) fn parse_biff8_shrfmla_record_with_rgcb(
    record: &records::LogicalBiffRecord<'_>,
) -> Result<(Vec<u8>, Vec<u8>), String> {
    let parsed = parse_biff8_shrfmla_record(record)?;
    Ok((parsed.rgce, parsed.rgcb))
}

pub(crate) fn parse_biff8_array_record(
    record: &records::LogicalBiffRecord<'_>,
) -> Result<ParsedArrayRecord, String> {
    let fragments: Vec<&[u8]> = record.fragments().collect();
    let cursor = FragmentCursor::new(&fragments, 0, 0);

    // ARRAY layouts vary slightly:
    // - RefU (6) vs Ref8 (8) for the array range header.
    // - Some producers include 2 or 4 bytes of reserved/flags before `cce`.
    for reserved_len in [2usize, 4] {
        let mut c = cursor.clone();
        if let Ok((rgce, rgcb)) = parse_array_with_refu(&mut c, reserved_len) {
            return Ok(ParsedArrayRecord { rgce, rgcb });
        }
    }
    for reserved_len in [2usize, 4] {
        let mut c = cursor.clone();
        if let Ok((rgce, rgcb)) = parse_array_with_ref8(&mut c, reserved_len) {
            return Ok(ParsedArrayRecord { rgce, rgcb });
        }
    }

    Err("unrecognized ARRAY record layout".to_string())
}

/// Parsed BIFF8 worksheet formula cell payload.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Biff8FormulaCell {
    pub(crate) cell: CellRef,
    /// Raw `FORMULA.grbit` flags.
    pub(crate) grbit: FormulaGrbit,
    /// Raw formula token stream (`rgce`).
    pub(crate) rgce: Vec<u8>,
    /// Trailing data blocks (`rgcb`) referenced by certain ptgs (notably `PtgArray`).
    pub(crate) rgcb: Vec<u8>,
}

/// Minimal parsed representation of a BIFF8 SHRFMLA record.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Biff8ShrFmlaRecord {
    pub(crate) range: (CellRef, CellRef),
    pub(crate) rgce: Vec<u8>,
    pub(crate) rgcb: Vec<u8>,
}

/// Minimal parsed representation of a BIFF8 ARRAY record.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Biff8ArrayRecord {
    pub(crate) range: (CellRef, CellRef),
    pub(crate) rgce: Vec<u8>,
    pub(crate) rgcb: Vec<u8>,
}

/// Minimal parsed representation of a BIFF8 TABLE record.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Biff8TableRecord {
    pub(crate) range: (CellRef, CellRef),
    /// Raw TABLE record payload (best-effort; preserved for diagnostics).
    pub(crate) data: Vec<u8>,
}

#[derive(Debug, Default)]
pub(crate) struct ParsedBiff8WorksheetFormulas {
    pub(crate) formula_cells: HashMap<CellRef, Biff8FormulaCell>,
    /// Shared formula definitions keyed by the anchor/base cell (top-left of `range`).
    pub(crate) shrfmla: HashMap<CellRef, Biff8ShrFmlaRecord>,
    /// Array formula definitions keyed by the anchor/base cell (top-left of `range`).
    pub(crate) array: HashMap<CellRef, Biff8ArrayRecord>,
    /// Data table definitions keyed by the anchor/base cell (top-left of `range`).
    pub(crate) table: HashMap<CellRef, Biff8TableRecord>,
    /// Non-fatal issues encountered during parse/resolution.
    pub(crate) warnings: Vec<crate::ImportWarning>,
}

/// Cap warnings collected during best-effort worksheet formula scans so a crafted `.xls` cannot
/// allocate an unbounded number of warnings.
const MAX_WARNINGS_PER_SHEET: usize = 50;
const WARNINGS_SUPPRESSED_MESSAGE: &str = "additional warnings suppressed";

fn warn(warnings: &mut Vec<crate::ImportWarning>, msg: impl Into<String>) {
    if warnings.len() < MAX_WARNINGS_PER_SHEET {
        warnings.push(crate::ImportWarning {
            message: msg.into(),
        });
        return;
    }
    // Add a single terminal warning so callers have a hint that the import was noisy.
    if warnings.len() == MAX_WARNINGS_PER_SHEET {
        warnings.push(crate::ImportWarning {
            message: WARNINGS_SUPPRESSED_MESSAGE.to_string(),
        });
    }
}

fn warn_string(warnings: &mut Vec<String>, msg: impl Into<String>) {
    if warnings.len() < MAX_WARNINGS_PER_SHEET {
        warnings.push(msg.into());
        return;
    }
    // Add a single terminal warning so callers have a hint that the import was noisy.
    if warnings.len() == MAX_WARNINGS_PER_SHEET {
        warnings.push(WARNINGS_SUPPRESSED_MESSAGE.to_string());
    }
}

fn parse_cell_ref_u16(row: u16, col: u16) -> Option<CellRef> {
    let row = row as u32;
    let col = col as u32;
    if row >= EXCEL_MAX_ROWS || col >= EXCEL_MAX_COLS {
        return None;
    }
    Some(CellRef::new(row, col))
}

// `Ref8` columns can carry flags in their high bits; mask down to the 14-bit payload.
const REF8_COL_MASK: u16 = 0x3FFF;

fn parse_ref8(data: &[u8]) -> Option<(u16, u16, u16, u16)> {
    let chunk = data.get(0..8)?;
    let rw_first = u16::from_le_bytes([chunk[0], chunk[1]]);
    let rw_last = u16::from_le_bytes([chunk[2], chunk[3]]);
    let col_first_raw = u16::from_le_bytes([chunk[4], chunk[5]]);
    let col_last_raw = u16::from_le_bytes([chunk[6], chunk[7]]);
    let col_first = col_first_raw & REF8_COL_MASK;
    let col_last = col_last_raw & REF8_COL_MASK;
    if rw_first > rw_last || col_first > col_last {
        return None;
    }
    Some((rw_first, rw_last, col_first, col_last))
}

fn parse_refu(data: &[u8]) -> Option<(u16, u16, u16, u16)> {
    let chunk = data.get(0..6)?;
    let rw_first = u16::from_le_bytes([chunk[0], chunk[1]]);
    let rw_last = u16::from_le_bytes([chunk[2], chunk[3]]);
    let col_first = chunk[4] as u16;
    let col_last = chunk[5] as u16;
    if rw_first > rw_last || col_first > col_last {
        return None;
    }
    Some((rw_first, rw_last, col_first, col_last))
}

fn parse_ref_any(data: &[u8]) -> Option<(u16, u16, u16, u16)> {
    // Prefer Ref8 when it decodes to "classic" `.xls` column bounds (<=255). Some producers store
    // Ref8 even when RefU would suffice.
    if let Some(r) = parse_ref8(data) {
        if r.2 <= 0x00FF && r.3 <= 0x00FF {
            return Some(r);
        }
    }

    parse_refu(data).or_else(|| parse_ref8(data))
}

fn parse_ref_any_best_effort(data: &[u8]) -> Option<(CellRef, CellRef)> {
    let (rw_first, rw_last, col_first, col_last) = parse_ref_any(data)?;
    let start = parse_cell_ref_u16(rw_first, col_first)?;
    let end = parse_cell_ref_u16(rw_last, col_last)?;
    Some((start, end))
}

fn parse_shrfmla_range_best_effort(data: &[u8], expected_cce: usize) -> Option<(CellRef, CellRef)> {
    // SHRFMLA stores the shared formula range using either:
    // - RefU: [rwFirst:u16][rwLast:u16][colFirst:u8][colLast:u8]
    // - Ref8: [rwFirst:u16][rwLast:u16][colFirst:u16][colLast:u16]
    //
    // Many producers include the `cUse` field after the range header, but some omit it. This makes
    // the header ambiguous in certain byte patterns. In particular:
    //   RefU(A..A) + cUse + cce
    // can be misread as:
    //   Ref8(A..?) + cce
    // when `cUse` is small, because `cce` then appears at the same offset.
    //
    // To disambiguate, match the stored `cce` against the parsed rgce length (`expected_cce`) and
    // prefer layouts whose `cUse` matches the range area.

    #[derive(Clone, Copy)]
    struct RangeHeader {
        rw_first: u16,
        rw_last: u16,
        col_first: u16,
        col_last: u16,
    }

    fn valid_range(h: RangeHeader) -> bool {
        h.rw_first <= h.rw_last && h.col_first <= h.col_last
    }

    fn range_area(h: RangeHeader) -> u64 {
        let rows = (h.rw_last.saturating_sub(h.rw_first) as u64).saturating_add(1);
        let cols = (h.col_last.saturating_sub(h.col_first) as u64).saturating_add(1);
        rows.saturating_mul(cols)
    }

    fn parse_refu_range(data: &[u8]) -> Option<RangeHeader> {
        let chunk = data.get(0..6)?;
        let rw_first = u16::from_le_bytes([chunk[0], chunk[1]]);
        let rw_last = u16::from_le_bytes([chunk[2], chunk[3]]);
        let col_first = chunk[4] as u16;
        let col_last = chunk[5] as u16;
        Some(RangeHeader {
            rw_first,
            rw_last,
            col_first,
            col_last,
        })
    }

    fn parse_ref8_range(data: &[u8]) -> Option<RangeHeader> {
        let chunk = data.get(0..8)?;
        let rw_first = u16::from_le_bytes([chunk[0], chunk[1]]);
        let rw_last = u16::from_le_bytes([chunk[2], chunk[3]]);
        let col_first = u16::from_le_bytes([chunk[4], chunk[5]]) & REF8_COL_MASK;
        let col_last = u16::from_le_bytes([chunk[6], chunk[7]]) & REF8_COL_MASK;
        Some(RangeHeader {
            rw_first,
            rw_last,
            col_first,
            col_last,
        })
    }

    #[derive(Clone, Copy)]
    struct Candidate {
        header: RangeHeader,
        uses_ref8: bool,
        cuse: Option<u16>,
    }

    let expected_cce_u16 = u16::try_from(expected_cce).ok()?;
    let mut candidates: Vec<Candidate> = Vec::new();
    let mut push_candidate = |header: Option<RangeHeader>,
                              uses_ref8: bool,
                              cuse: Option<u16>,
                              cce_offset: usize| {
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

    // Match the parsing order used by `parse_biff8_shrfmla_record`.
    // Layout A: RefU (6) + cUse (2) + cce (2).
    let cuse_a = data
        .get(6..8)
        .map(|v| u16::from_le_bytes([v[0], v[1]]));
    push_candidate(parse_refu_range(data), false, cuse_a, 8);
    // Layout B: Ref8 (8) + cUse (2) + cce (2).
    let cuse_b = data
        .get(8..10)
        .map(|v| u16::from_le_bytes([v[0], v[1]]));
    push_candidate(parse_ref8_range(data), true, cuse_b, 10);
    // Layout C: RefU (6) + cce (2) (cUse omitted).
    push_candidate(parse_refu_range(data), false, None, 6);
    // Layout D: Ref8 (8) + cce (2) (cUse omitted).
    push_candidate(parse_ref8_range(data), true, None, 8);

    let selected = if candidates.is_empty() {
        None
    } else {
        candidates
            .into_iter()
            .min_by_key(|c| {
                let area = range_area(c.header);
                let cuse_rank: u8 = match c.cuse {
                    Some(cuse) if cuse != 0 && (cuse as u64) == area => 0,
                    Some(0) | None => 1,
                    Some(_) => 2,
                };
                (
                    cuse_rank,
                    if c.uses_ref8 { 0u8 } else { 1u8 },
                    area,
                    c.header.rw_first,
                    c.header.rw_last,
                    c.header.col_first,
                    c.header.col_last,
                )
            })
            .map(|c| c.header)
    };

    let header = match selected {
        Some(h) => h,
        None => {
            // Fallback: parse without disambiguation (matches prior behaviour).
            let (start, end) = parse_ref_any_best_effort(data)?;
            return Some((start, end));
        }
    };

    let start = parse_cell_ref_u16(header.rw_first, header.col_first)?;
    let end = parse_cell_ref_u16(header.rw_last, header.col_last)?;
    Some((start, end))
}

/// Best-effort parsing of a BIFF8 worksheet substream's formula-related records.
///
/// This collects:
/// - Cell `FORMULA` records (including `grbit`)
/// - Shared formula `SHRFMLA` definitions
/// - Array formula `ARRAY` definitions
/// - Data table `TABLE` definitions
pub(crate) fn parse_biff8_worksheet_formulas(
    workbook_stream: &[u8],
    start: usize,
) -> Result<ParsedBiff8WorksheetFormulas, String> {
    let mut out = ParsedBiff8WorksheetFormulas::default();

    let allows_continuation = |id: u16| {
        id == RECORD_FORMULA || id == RECORD_SHRFMLA || id == RECORD_ARRAY || id == RECORD_TABLE
    };
    let mut iter =
        records::LogicalBiffRecordIter::from_offset(workbook_stream, start, allows_continuation)?;

    while let Some(next) = iter.next() {
        let record = match next {
            Ok(r) => r,
            Err(err) => {
                warn(
                    &mut out.warnings,
                    format!("malformed BIFF record in worksheet stream: {err}"),
                );
                break;
            }
        };

        if record.offset != start && records::is_bof_record(record.record_id) {
            break;
        }
        if record.record_id == records::RECORD_EOF {
            break;
        }

        match record.record_id {
            RECORD_FORMULA => match parse_biff8_formula_record(&record) {
                Ok(parsed_formula) => {
                    let Some(cell) = parse_cell_ref_u16(parsed_formula.row, parsed_formula.col)
                    else {
                        continue;
                    };
                    out.formula_cells.insert(
                        cell,
                        Biff8FormulaCell {
                            cell,
                            grbit: parsed_formula.grbit,
                            rgce: parsed_formula.rgce,
                            rgcb: parsed_formula.rgcb,
                        },
                    );
                }
                Err(err) => warn(
                    &mut out.warnings,
                    format!(
                        "failed to parse FORMULA record at offset {}: {err}",
                        record.offset
                    ),
                ),
            },
            RECORD_SHRFMLA => {
                match parse_biff8_shrfmla_record(&record) {
                    Ok(parsed) => {
                        let Some(range) = parse_shrfmla_range_best_effort(
                            record.data.as_ref(),
                            parsed.rgce.len(),
                        ) else {
                            warn(
                                &mut out.warnings,
                                format!(
                                    "failed to parse SHRFMLA range at offset {} (len={})",
                                    record.offset,
                                    record.data.len()
                                ),
                            );
                            continue;
                        };
                        let anchor = range.0;
                        out.shrfmla.insert(
                            anchor,
                            Biff8ShrFmlaRecord {
                                range,
                                rgce: parsed.rgce,
                                rgcb: parsed.rgcb,
                            },
                        );
                    }
                    Err(err) => warn(
                        &mut out.warnings,
                        format!(
                            "failed to parse SHRFMLA record at offset {}: {err}",
                            record.offset
                        ),
                    ),
                };
            }
            RECORD_ARRAY => {
                let Some(range) = parse_ref_any_best_effort(record.data.as_ref()) else {
                    warn(
                        &mut out.warnings,
                        format!(
                            "failed to parse ARRAY range at offset {} (len={})",
                            record.offset,
                            record.data.len()
                        ),
                    );
                    continue;
                };
                let anchor = range.0;
                match parse_biff8_array_record(&record) {
                    Ok(parsed) => {
                        out.array.insert(
                            anchor,
                            Biff8ArrayRecord {
                                range,
                                rgce: parsed.rgce,
                                rgcb: parsed.rgcb,
                            },
                        );
                    }
                    Err(err) => warn(
                        &mut out.warnings,
                        format!(
                            "failed to parse ARRAY record at offset {}: {err}",
                            record.offset
                        ),
                    ),
                }
            }
            RECORD_TABLE => {
                let Some(range) = parse_ref_any_best_effort(record.data.as_ref()) else {
                    warn(
                        &mut out.warnings,
                        format!(
                            "failed to parse TABLE range at offset {} (len={})",
                            record.offset,
                            record.data.len()
                        ),
                    );
                    continue;
                };
                let anchor = range.0;
                out.table.insert(
                    anchor,
                    Biff8TableRecord {
                        range,
                        data: record.data.as_ref().to_vec(),
                    },
                );
            }
            _ => {}
        }
    }

    Ok(out)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum LeadingPtg {
    /// `PtgExp` (0x01): points at a shared/array formula definition.
    Exp {
        base: CellRef,
    },
    /// `PtgTbl` (0x02): points at a data table definition.
    Tbl {
        base: CellRef,
    },
    Other,
    Empty,
}

fn decode_leading_ptg(rgce: &[u8]) -> LeadingPtg {
    if rgce.is_empty() {
        return LeadingPtg::Empty;
    }

    match rgce[0] {
        0x01 | 0x02 => {
            if rgce.len() < 5 {
                return LeadingPtg::Other;
            }
            let row = u16::from_le_bytes([rgce[1], rgce[2]]);
            let col = u16::from_le_bytes([rgce[3], rgce[4]]);
            let Some(base) = parse_cell_ref_u16(row, col) else {
                return LeadingPtg::Other;
            };
            if rgce[0] == 0x01 {
                LeadingPtg::Exp { base }
            } else {
                LeadingPtg::Tbl { base }
            }
        }
        _ => LeadingPtg::Other,
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum PtgReferenceResolution {
    /// No PtgExp/PtgTbl indirection; use the cell's own `rgce`.
    None,
    Shared {
        base: CellRef,
    },
    Array {
        base: CellRef,
    },
    Table {
        base: CellRef,
    },
    /// `PtgExp`/`PtgTbl` present but no suitable backing record could be found.
    Unresolved,
}

fn resolve_anchor_by_range_containment<T>(
    records: &HashMap<CellRef, T>,
    ptgexp_base: CellRef,
    cell: CellRef,
    range_of: impl Fn(&T) -> (CellRef, CellRef),
) -> Option<CellRef> {
    // Fast path: most producers use the range anchor as the PtgExp/PtgTbl base cell, so the key
    // lookup succeeds.
    if records.contains_key(&ptgexp_base) {
        return Some(ptgexp_base);
    }

    // Best-effort fallback: some `.xls` producers point PtgExp/PtgTbl at a *non-anchor* cell inside
    // the backing SHRFMLA/ARRAY/TABLE range. In that case, scan for a definition range that
    // contains both the current cell and the referenced base cell, and return its anchor.
    let mut matches: Vec<CellRef> = records
        .iter()
        .filter_map(|(anchor, record)| {
            let range = range_of(record);
            if range_contains_cell(range, cell) && range_contains_cell(range, ptgexp_base) {
                Some(*anchor)
            } else {
                None
            }
        })
        .collect();

    if matches.is_empty() {
        return None;
    }

    // Deterministic selection: choose the top-most, left-most anchor.
    matches.sort_by_key(|cell| (cell.row, cell.col));
    matches.first().copied()
}

/// Resolve a BIFF8 `FORMULA.rgce` that begins with `PtgExp` or `PtgTbl` into a backing record type.
///
/// This is best-effort and uses [`FormulaGrbit`] flags as a hint to disambiguate whether `PtgExp`
/// should resolve via `SHRFMLA` (shared formula) vs `ARRAY` (array formula).
pub(crate) fn resolve_ptgexp_or_ptgtbl_best_effort(
    parsed: &ParsedBiff8WorksheetFormulas,
    cell: &Biff8FormulaCell,
    warnings: &mut Vec<crate::ImportWarning>,
) -> PtgReferenceResolution {
    let leading = decode_leading_ptg(&cell.rgce);
    let base = match leading {
        LeadingPtg::Exp { base } | LeadingPtg::Tbl { base } => Some(base),
        _ => None,
    };

    let hint = cell.grbit.membership_hint();

    // Warn on inconsistent flag/token combinations (but keep decoding best-effort).
    match (hint, leading) {
        (Some(FormulaMembershipHint::Table), LeadingPtg::Tbl { .. }) => {}
        (Some(FormulaMembershipHint::Table), _) => warn(
            warnings,
            format!(
                "formula {} has grbit.fTbl set but rgce does not start with PtgTbl",
                cell.cell.to_a1()
            ),
        ),
        (
            Some(FormulaMembershipHint::Shared) | Some(FormulaMembershipHint::Array),
            LeadingPtg::Exp { .. },
        ) => {}
        (Some(FormulaMembershipHint::Shared), _) => warn(
            warnings,
            format!(
                "formula {} has grbit.fShrFmla set but rgce does not start with PtgExp",
                cell.cell.to_a1()
            ),
        ),
        (Some(FormulaMembershipHint::Array), _) => warn(
            warnings,
            format!(
                "formula {} has grbit.fArray set but rgce does not start with PtgExp",
                cell.cell.to_a1()
            ),
        ),
        _ => {}
    }

    // Candidate resolution order:
    // 1) Flag-indicated membership (if unambiguous)
    // 2) Token-based heuristics (`PtgExp` -> prefer SHRFMLA then ARRAY; `PtgTbl` -> TABLE)
    let mut order: Vec<FormulaMembershipHint> = Vec::new();
    if let Some(h) = hint {
        order.push(h);
    }
    match leading {
        LeadingPtg::Exp { .. } => {
            order.push(FormulaMembershipHint::Shared);
            order.push(FormulaMembershipHint::Array);
        }
        LeadingPtg::Tbl { .. } => {
            order.push(FormulaMembershipHint::Table);
        }
        LeadingPtg::Empty | LeadingPtg::Other => {}
    }
    // Deduplicate while preserving first occurrence (small vector).
    let mut seen = HashSet::new();
    order.retain(|k| seen.insert(*k));

    let Some(base) = base else {
        return match leading {
            LeadingPtg::Empty | LeadingPtg::Other => PtgReferenceResolution::None,
            _ => PtgReferenceResolution::Unresolved,
        };
    };

    for kind in order {
        match kind {
            FormulaMembershipHint::Shared => {
                if let Some(anchor) =
                    resolve_anchor_by_range_containment(&parsed.shrfmla, base, cell.cell, |r| r.range)
                {
                    return PtgReferenceResolution::Shared { base: anchor };
                }
                if hint == Some(FormulaMembershipHint::Shared) {
                    warn(
                        warnings,
                        format!(
                            "formula {} indicates shared formula membership (grbit.fShrFmla), but no SHRFMLA record was found for base {}",
                            cell.cell.to_a1(),
                            base.to_a1()
                        ),
                    );
                }
            }
            FormulaMembershipHint::Array => {
                if let Some(anchor) =
                    resolve_anchor_by_range_containment(&parsed.array, base, cell.cell, |r| r.range)
                {
                    return PtgReferenceResolution::Array { base: anchor };
                }
                if hint == Some(FormulaMembershipHint::Array) {
                    warn(
                        warnings,
                        format!(
                            "formula {} indicates array formula membership (grbit.fArray), but no ARRAY record was found for base {}",
                            cell.cell.to_a1(),
                            base.to_a1()
                        ),
                    );
                }
            }
            FormulaMembershipHint::Table => {
                if let Some(anchor) =
                    resolve_anchor_by_range_containment(&parsed.table, base, cell.cell, |r| r.range)
                {
                    return PtgReferenceResolution::Table { base: anchor };
                }
                if hint == Some(FormulaMembershipHint::Table) {
                    warn(
                        warnings,
                        format!(
                            "formula {} indicates table formula membership (grbit.fTbl), but no TABLE record was found for base {}",
                            cell.cell.to_a1(),
                            base.to_a1()
                        ),
                    );
                }
            }
        }
    }

    // No candidates matched.
    match leading {
        LeadingPtg::Empty | LeadingPtg::Other => PtgReferenceResolution::None,
        _ => PtgReferenceResolution::Unresolved,
    }
}

fn parse_shrfmla_with_refu(cursor: &mut FragmentCursor<'_>) -> Result<(Vec<u8>, Vec<u8>), String> {
    // ref (rwFirst:u16, rwLast:u16, colFirst:u8, colLast:u8)
    cursor.skip_bytes(2 + 2 + 1 + 1)?;
    // cUse
    cursor.skip_bytes(2)?;
    let cce = cursor.read_u16_le()? as usize;
    let rgce = cursor.read_biff8_rgce(cce)?;
    let rgcb = cursor.read_remaining_bytes()?;
    Ok((rgce, rgcb))
}

fn parse_shrfmla_with_refu_no_cuse(
    cursor: &mut FragmentCursor<'_>,
) -> Result<(Vec<u8>, Vec<u8>), String> {
    // ref (rwFirst:u16, rwLast:u16, colFirst:u8, colLast:u8)
    cursor.skip_bytes(2 + 2 + 1 + 1)?;
    // cce
    let cce = cursor.read_u16_le()? as usize;
    let rgce = cursor.read_biff8_rgce(cce)?;
    let rgcb = cursor.read_remaining_bytes()?;
    Ok((rgce, rgcb))
}

fn parse_shrfmla_with_ref8(cursor: &mut FragmentCursor<'_>) -> Result<(Vec<u8>, Vec<u8>), String> {
    // ref (rwFirst:u16, rwLast:u16, colFirst:u16, colLast:u16)
    cursor.skip_bytes(8)?;
    // cUse
    cursor.skip_bytes(2)?;
    let cce = cursor.read_u16_le()? as usize;
    let rgce = cursor.read_biff8_rgce(cce)?;
    let rgcb = cursor.read_remaining_bytes()?;
    Ok((rgce, rgcb))
}

fn parse_shrfmla_with_ref8_no_cuse(
    cursor: &mut FragmentCursor<'_>,
) -> Result<(Vec<u8>, Vec<u8>), String> {
    // ref (rwFirst:u16, rwLast:u16, colFirst:u16, colLast:u16)
    cursor.skip_bytes(8)?;
    let cce = cursor.read_u16_le()? as usize;
    let rgce = cursor.read_biff8_rgce(cce)?;
    let rgcb = cursor.read_remaining_bytes()?;
    Ok((rgce, rgcb))
}

fn parse_array_with_refu(
    cursor: &mut FragmentCursor<'_>,
    reserved_len: usize,
) -> Result<(Vec<u8>, Vec<u8>), String> {
    // ref (rwFirst:u16, rwLast:u16, colFirst:u8, colLast:u8)
    cursor.skip_bytes(2 + 2 + 1 + 1)?;
    cursor.skip_bytes(reserved_len)?;
    let cce = cursor.read_u16_le()? as usize;
    let rgce = cursor.read_biff8_rgce(cce)?;
    let rgcb = cursor.read_remaining_bytes()?;
    Ok((rgce, rgcb))
}

fn parse_array_with_ref8(
    cursor: &mut FragmentCursor<'_>,
    reserved_len: usize,
) -> Result<(Vec<u8>, Vec<u8>), String> {
    // ref (rwFirst:u16, rwLast:u16, colFirst:u16, colLast:u16)
    cursor.skip_bytes(8)?;
    cursor.skip_bytes(reserved_len)?;
    let cce = cursor.read_u16_le()? as usize;
    let rgce = cursor.read_biff8_rgce(cce)?;
    let rgcb = cursor.read_remaining_bytes()?;
    Ok((rgce, rgcb))
}

#[derive(Debug, Clone)]
struct FragmentCursor<'a> {
    fragments: &'a [&'a [u8]],
    frag_idx: usize,
    offset: usize,
}

impl<'a> FragmentCursor<'a> {
    fn new(fragments: &'a [&'a [u8]], frag_idx: usize, offset: usize) -> Self {
        Self {
            fragments,
            frag_idx,
            offset,
        }
    }

    fn remaining_in_fragment(&self) -> usize {
        self.fragments
            .get(self.frag_idx)
            .map(|f| f.len().saturating_sub(self.offset))
            .unwrap_or(0)
    }

    fn advance_fragment(&mut self) -> Result<(), String> {
        self.frag_idx = self
            .frag_idx
            .checked_add(1)
            .ok_or_else(|| "fragment index overflow".to_string())?;
        self.offset = 0;
        if self.frag_idx >= self.fragments.len() {
            return Err("unexpected end of record".to_string());
        }
        Ok(())
    }

    fn read_u8(&mut self) -> Result<u8, String> {
        loop {
            let frag = self
                .fragments
                .get(self.frag_idx)
                .ok_or_else(|| "unexpected end of record".to_string())?;
            if self.offset < frag.len() {
                let b = frag[self.offset];
                self.offset += 1;
                return Ok(b);
            }
            self.advance_fragment()?;
        }
    }

    fn read_u16_le(&mut self) -> Result<u16, String> {
        let lo = self.read_u8()?;
        let hi = self.read_u8()?;
        Ok(u16::from_le_bytes([lo, hi]))
    }

    fn read_u32_le(&mut self) -> Result<u32, String> {
        let b0 = self.read_u8()?;
        let b1 = self.read_u8()?;
        let b2 = self.read_u8()?;
        let b3 = self.read_u8()?;
        Ok(u32::from_le_bytes([b0, b1, b2, b3]))
    }

    fn read_exact_from_current(&mut self, n: usize) -> Result<&'a [u8], String> {
        let frag = self
            .fragments
            .get(self.frag_idx)
            .ok_or_else(|| "unexpected end of record".to_string())?;
        let end = self
            .offset
            .checked_add(n)
            .ok_or_else(|| "offset overflow".to_string())?;
        if end > frag.len() {
            return Err("unexpected end of record".to_string());
        }
        let out = &frag[self.offset..end];
        self.offset = end;
        Ok(out)
    }

    fn read_bytes(&mut self, mut n: usize) -> Result<Vec<u8>, String> {
        let mut out = Vec::with_capacity(n);
        while n > 0 {
            let available = self.remaining_in_fragment();
            if available == 0 {
                self.advance_fragment()?;
                continue;
            }
            let take = n.min(available);
            let bytes = self.read_exact_from_current(take)?;
            out.extend_from_slice(bytes);
            n -= take;
        }
        Ok(out)
    }

    fn remaining_total_bytes(&self) -> usize {
        let mut total = 0usize;
        for (idx, frag) in self.fragments.iter().enumerate().skip(self.frag_idx) {
            if idx == self.frag_idx {
                total = total.saturating_add(frag.len().saturating_sub(self.offset));
            } else {
                total = total.saturating_add(frag.len());
            }
        }
        total
    }

    fn read_remaining_bytes(&mut self) -> Result<Vec<u8>, String> {
        let remaining = self.remaining_total_bytes();
        self.read_bytes(remaining)
    }

    fn skip_bytes(&mut self, mut n: usize) -> Result<(), String> {
        while n > 0 {
            let available = self.remaining_in_fragment();
            if available == 0 {
                self.advance_fragment()?;
                continue;
            }
            let take = n.min(available);
            self.offset += take;
            n -= take;
        }
        Ok(())
    }

    fn advance_fragment_in_biff8_string(&mut self, is_unicode: &mut bool) -> Result<(), String> {
        self.advance_fragment()?;
        // When a BIFF8 string spans a CONTINUE boundary, Excel inserts a 1-byte option flags prefix
        // at the start of the continued fragment. The only relevant bit for formula string tokens is
        // `fHighByte` (unicode vs compressed).
        let cont_flags = self.read_u8()?;
        *is_unicode = (cont_flags & STR_FLAG_HIGH_BYTE) != 0;
        Ok(())
    }

    fn read_biff8_string_bytes(
        &mut self,
        mut n: usize,
        is_unicode: &mut bool,
    ) -> Result<Vec<u8>, String> {
        // Read `n` canonical bytes from a BIFF8 continued string payload, skipping the 1-byte
        // continuation flags prefix that appears at the start of each continued fragment.
        let mut out = Vec::with_capacity(n);
        while n > 0 {
            if self.remaining_in_fragment() == 0 {
                self.advance_fragment_in_biff8_string(is_unicode)?;
                continue;
            }
            let available = self.remaining_in_fragment();
            let take = n.min(available);
            let bytes = self.read_exact_from_current(take)?;
            out.extend_from_slice(bytes);
            n -= take;
        }
        Ok(out)
    }

    fn read_biff8_rgce(&mut self, cce: usize) -> Result<Vec<u8>, String> {
        // Best-effort: parse BIFF8 ptg tokens so we can skip the continuation flags byte injected
        // at fragment boundaries when a `PtgStr` (ShortXLUnicodeString) payload is split across
        // `CONTINUE` records.
        //
        // If we encounter an unsupported token, fall back to raw byte copying for the remainder of
        // the `rgce` stream (without special continuation handling).
        let mut out = Vec::with_capacity(cce);

        while out.len() < cce {
            let ptg = self.read_u8()?;
            out.push(ptg);

            match ptg {
                // PtgExp / PtgTbl: shared/array formula tokens.
                0x01 | 0x02 => {
                    // Canonical BIFF8 payload is 4 bytes (`[rw:u16][col:u16]`).
                    //
                    // Best-effort: some producers emit wider coordinates (e.g. `[rw:u32][col:u16]`
                    // or `[rw:u32][col:u32]`) even in `.xls` files. Those appear as *non-canonical*
                    // token stream lengths where the entire `rgce` is just `PtgExp`/`PtgTbl`:
                    //   cce = 1 + payload_len
                    //
                    // Preserve the normal fixed-width behavior so subsequent tokens (e.g. `PtgStr`)
                    // stay aligned and we can still skip continuation flags.
                    let remaining = cce.saturating_sub(out.len());
                    let payload_len = match (out.len(), remaining) {
                        // Non-standard payload widths: treat the whole stream as a single token.
                        (1, 6) | (1, 8) => remaining,
                        // Canonical BIFF8.
                        _ => 4.min(remaining),
                    };
                    let bytes = self.read_bytes(payload_len)?;
                    out.extend_from_slice(&bytes);
                }
                // Binary operators.
                0x03..=0x11
                // Unary +/- and postfix/paren/missarg.
                | 0x12
                | 0x13
                | 0x14
                | 0x15
                | 0x16 => {}
                // Spill range postfix (`#`).
                0x2F => {}
                // PtgStr (ShortXLUnicodeString) [MS-XLS 2.5.293]
                0x17 => {
                    let cch = self.read_u8()? as usize;
                    let flags = self.read_u8()?;
                    out.push(cch as u8);
                    out.push(flags);

                    let mut is_unicode = (flags & STR_FLAG_HIGH_BYTE) != 0;

                    let richtext_runs = if (flags & STR_FLAG_RICH_TEXT) != 0 {
                        let bytes = self.read_biff8_string_bytes(2, &mut is_unicode)?;
                        out.extend_from_slice(&bytes);
                        u16::from_le_bytes([bytes[0], bytes[1]]) as usize
                    } else {
                        0
                    };

                    let ext_size = if (flags & STR_FLAG_EXT) != 0 {
                        let bytes = self.read_biff8_string_bytes(4, &mut is_unicode)?;
                        out.extend_from_slice(&bytes);
                        u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]) as usize
                    } else {
                        0
                    };

                    let mut remaining_chars = cch;

                    while remaining_chars > 0 {
                        if self.remaining_in_fragment() == 0 {
                            self.advance_fragment_in_biff8_string(&mut is_unicode)?;
                            continue;
                        }

                        let bytes_per_char = if is_unicode { 2 } else { 1 };
                        let available_bytes = self.remaining_in_fragment();
                        let available_chars = available_bytes / bytes_per_char;
                        if available_chars == 0 {
                            return Err("string continuation split mid-character".to_string());
                        }

                        let take_chars = remaining_chars.min(available_chars);
                        let take_bytes = take_chars * bytes_per_char;
                        if out.len() + take_bytes > cce {
                            return Err(
                                "PtgStr character payload exceeds declared rgce length"
                                    .to_string(),
                            );
                        }
                        let bytes = self.read_exact_from_current(take_bytes)?;
                        out.extend_from_slice(bytes);
                        remaining_chars -= take_chars;
                    }

                    let richtext_bytes = richtext_runs
                        .checked_mul(4)
                        .ok_or_else(|| "rich text run count overflow".to_string())?;
                    let extra_len = richtext_bytes
                        .checked_add(ext_size)
                        .ok_or_else(|| "PtgStr extra payload length overflow".to_string())?;
                    if extra_len > 0 {
                        let remaining = cce.saturating_sub(out.len());
                        if extra_len > remaining {
                            return Err(
                                "PtgStr extra payload exceeds declared rgce length".to_string(),
                            );
                        }
                        let extra = self.read_biff8_string_bytes(extra_len, &mut is_unicode)?;
                        out.extend_from_slice(&extra);
                    }
                }
                // PtgExtend* token 0x18 (and class variants).
                0x18 | 0x38 | 0x58 | 0x78 => {
                    let etpg = self.read_u8()?;
                    out.push(etpg);
                    if etpg == 0x19 {
                        let bytes = self.read_bytes(12)?;
                        out.extend_from_slice(&bytes);
                    } else {
                        let bytes = self.read_bytes(4)?;
                        out.extend_from_slice(&bytes);
                    }
                }
                // PtgAttr (evaluation hints / jump tables).
                0x19 => {
                    let grbit = self.read_u8()?;
                    let w_attr = self.read_u16_le()?;
                    out.push(grbit);
                    out.extend_from_slice(&w_attr.to_le_bytes());

                    // tAttrChoose includes a jump table of `u16` offsets (wAttr entries).
                    const T_ATTR_CHOOSE: u8 = 0x04;
                    if (grbit & T_ATTR_CHOOSE) != 0 {
                        let entries = w_attr as usize;
                        let bytes = entries
                            .checked_mul(2)
                            .ok_or_else(|| "tAttrChoose jump table length overflow".to_string())?;
                        let table = self.read_bytes(bytes)?;
                        out.extend_from_slice(&table);
                    }
                }
                // PtgErr / PtgBool (1 byte)
                0x1C | 0x1D => {
                    out.push(self.read_u8()?);
                }
                // PtgInt (2 bytes)
                0x1E => {
                    let bytes = self.read_bytes(2)?;
                    out.extend_from_slice(&bytes);
                }
                // PtgNum (8 bytes)
                0x1F => {
                    let bytes = self.read_bytes(8)?;
                    out.extend_from_slice(&bytes);
                }
                // PtgArray (7 bytes) [MS-XLS 2.5.198.8]
                0x20 | 0x40 | 0x60 => {
                    let bytes = self.read_bytes(7)?;
                    out.extend_from_slice(&bytes);
                }
                // PtgFunc (2 bytes)
                0x21 | 0x41 | 0x61 => {
                    let bytes = self.read_bytes(2)?;
                    out.extend_from_slice(&bytes);
                }
                // PtgFuncVar (3 bytes)
                0x22 | 0x42 | 0x62 => {
                    let bytes = self.read_bytes(3)?;
                    out.extend_from_slice(&bytes);
                }
                // PtgName (defined name reference) (6 bytes).
                0x23 | 0x43 | 0x63 => {
                    let bytes = self.read_bytes(6)?;
                    out.extend_from_slice(&bytes);
                }
                // PtgRef (4 bytes)
                0x24 | 0x44 | 0x64 => {
                    let bytes = self.read_bytes(4)?;
                    out.extend_from_slice(&bytes);
                }
                // PtgArea (8 bytes)
                0x25 | 0x45 | 0x65 => {
                    let bytes = self.read_bytes(8)?;
                    out.extend_from_slice(&bytes);
                }
                // PtgRefErr (4 bytes)
                0x2A | 0x4A | 0x6A => {
                    let bytes = self.read_bytes(4)?;
                    out.extend_from_slice(&bytes);
                }
                // PtgAreaErr (8 bytes)
                0x2B | 0x4B | 0x6B => {
                    let bytes = self.read_bytes(8)?;
                    out.extend_from_slice(&bytes);
                }
                // PtgRefN (4 bytes)
                0x2C | 0x4C | 0x6C => {
                    let bytes = self.read_bytes(4)?;
                    out.extend_from_slice(&bytes);
                }
                // PtgAreaN (8 bytes)
                0x2D | 0x4D | 0x6D => {
                    let bytes = self.read_bytes(8)?;
                    out.extend_from_slice(&bytes);
                }
                // PtgNameX (external name) [MS-XLS 2.5.198.41]
                0x39 | 0x59 | 0x79 => {
                    let bytes = self.read_bytes(6)?;
                    out.extend_from_slice(&bytes);
                }
                // 3D references: PtgRef3d / PtgArea3d.
                0x3A | 0x5A | 0x7A => {
                    let bytes = self.read_bytes(6)?;
                    out.extend_from_slice(&bytes);
                }
                0x3B | 0x5B | 0x7B => {
                    let bytes = self.read_bytes(10)?;
                    out.extend_from_slice(&bytes);
                }
                // 3D error references: PtgRefErr3d / PtgAreaErr3d.
                0x3C | 0x5C | 0x7C => {
                    let bytes = self.read_bytes(6)?;
                    out.extend_from_slice(&bytes);
                }
                0x3D | 0x5D | 0x7D => {
                    let bytes = self.read_bytes(10)?;
                    out.extend_from_slice(&bytes);
                }
                // 3D relative references: PtgRefN3d / PtgAreaN3d.
                0x3E | 0x5E | 0x7E => {
                    let bytes = self.read_bytes(6)?;
                    out.extend_from_slice(&bytes);
                }
                0x3F | 0x5F | 0x7F => {
                    let bytes = self.read_bytes(10)?;
                    out.extend_from_slice(&bytes);
                }
                // PtgMem* tokens: [ptg][cce: u16][rgce: cce bytes]
                0x26 | 0x46 | 0x66 | 0x27 | 0x47 | 0x67 | 0x28 | 0x48 | 0x68 | 0x29 | 0x49
                | 0x69 | 0x2E | 0x4E | 0x6E => {
                    let inner_cce = self.read_u16_le()? as usize;
                    out.extend_from_slice(&(inner_cce as u16).to_le_bytes());
                    let inner = self.read_biff8_rgce(inner_cce)?;
                    out.extend_from_slice(&inner);
                }
                _ => {
                    // Unsupported token: copy the remaining bytes as-is to satisfy the `cce`
                    // contract and avoid dropping the formula entirely.
                    let remaining = cce.saturating_sub(out.len());
                    if remaining > 0 {
                        let bytes = self.read_bytes(remaining)?;
                        out.extend_from_slice(&bytes);
                    }
                }
            }
        }

        if out.len() != cce {
            return Err(format!(
                "rgce length mismatch (expected {cce} bytes, got {})",
                out.len()
            ));
        }

        Ok(out)
    }
}

// -------------------------------------------------------------------------------------------------
// Non-standard `PtgExp` / `PtgTbl` payload width recovery
// -------------------------------------------------------------------------------------------------
//
// BIFF8 specifies a 4-byte payload for `PtgExp` / `PtgTbl` tokens:
//   [rw: u16][col: u16]
//
// In the wild, some producers embed wider row/col values (e.g. BIFF12/XLSB-like u32 fields) even
// inside `.xls` files. Libraries that assume the 4-byte payload can panic or fail to resolve the
// token back to a SHRFMLA/ARRAY/TABLE definition.
//
// This section implements a narrow best-effort parser that:
// - Accepts multiple payload widths (u32/u32, u32/u16, u16/u16)
// - Chooses the first candidate that matches a SHRFMLA/ARRAY/TABLE definition and contains the
//   current cell in that definition’s range
// - Only applies to non-canonical token stream lengths (`cce != 5`) to avoid changing behavior for
//   normal BIFF8 encodings.

const BIFF8_MAX_ROW0: u32 = u16::MAX as u32;
const BIFF8_MAX_COL0: u32 = 0x00FF;

#[derive(Debug, Default)]
pub(crate) struct ParsedWorksheetExpFormulas {
    /// Resolved formulas keyed by cell reference (0-based).
    pub(crate) formulas: HashMap<CellRef, String>,
    /// Non-fatal issues encountered while resolving/decoding.
    pub(crate) warnings: Vec<String>,
}

/// Return plausible `(row, col)` base-cell coordinate candidates for a BIFF8 `PtgExp`/`PtgTbl`
/// payload.
///
/// Candidates are returned in preference order:
/// - row u32 + col u32 (8 bytes)
/// - row u32 + col u16 (6 bytes)
/// - row u16 + col u16 (4 bytes)
///
/// Candidates are filtered to BIFF8/Excel 2003 bounds (row <= 65535, col <= 255).
pub(crate) fn ptgexp_candidates(payload: &[u8]) -> Vec<(u32, u32)> {
    const MAX_ROW: u32 = BIFF8_MAX_ROW0;
    const MAX_COL: u32 = BIFF8_MAX_COL0;

    let mut candidates: Vec<(u32, u32, usize)> = Vec::new();

    if payload.len() >= 8 {
        let row = u32::from_le_bytes([payload[0], payload[1], payload[2], payload[3]]);
        let col = u32::from_le_bytes([payload[4], payload[5], payload[6], payload[7]]);
        if row <= MAX_ROW && col <= MAX_COL {
            candidates.push((row, col, 8));
        }
    }

    if payload.len() >= 6 {
        let row = u32::from_le_bytes([payload[0], payload[1], payload[2], payload[3]]);
        let col = u16::from_le_bytes([payload[4], payload[5]]) as u32;
        if row <= MAX_ROW && col <= MAX_COL {
            candidates.push((row, col, 6));
        }
    }

    if payload.len() >= 4 {
        let row = u16::from_le_bytes([payload[0], payload[1]]) as u32;
        let col = u16::from_le_bytes([payload[2], payload[3]]) as u32;
        if row <= MAX_ROW && col <= MAX_COL {
            candidates.push((row, col, 4));
        }
    }

    // Prefer wider payload interpretations deterministically.
    candidates.sort_by_key(|(_, _, n)| *n);
    candidates.reverse();

    // Deduplicate coordinates while preserving width preference (avoid spurious ambiguity warnings
    // when the upper bytes are all zero).
    let mut out: Vec<(u32, u32)> = Vec::new();
    for (r, c, _) in candidates {
        if !out.iter().any(|(rr, cc)| *rr == r && *cc == c) {
            out.push((r, c));
        }
    }
    out
}

#[derive(Debug, Clone, Copy)]
enum ExpKind {
    Exp,
    Tbl,
}

#[derive(Debug, Clone)]
struct PendingExp {
    kind: ExpKind,
    cell: CellRef,
    payload: Vec<u8>,
}

fn range_contains_cell(range: (CellRef, CellRef), cell: CellRef) -> bool {
    cell.row >= range.0.row
        && cell.row <= range.1.row
        && cell.col >= range.0.col
        && cell.col <= range.1.col
}

fn parse_formula_record_for_wide_ptgexp(
    record: &records::LogicalBiffRecord<'_>,
) -> Result<Option<PendingExp>, String> {
    // FORMULA record header is 22 bytes before rgce.
    let data = record.data.as_ref();
    if data.len() < 22 {
        return Ok(None);
    }
    let row = u16::from_le_bytes([data[0], data[1]]) as u32;
    let col = u16::from_le_bytes([data[2], data[3]]) as u32;
    if row >= EXCEL_MAX_ROWS || col >= EXCEL_MAX_COLS {
        return Ok(None);
    }
    let cell = CellRef::new(row, col);

    let cce = u16::from_le_bytes([data[20], data[21]]) as usize;
    let rgce_start = 22usize;
    let rgce_end = rgce_start
        .checked_add(cce)
        .ok_or_else(|| "FORMULA cce overflow".to_string())?;
    if data.len() < rgce_end {
        return Err(format!(
            "truncated FORMULA rgce payload at offset {} (cce={cce}, have {} bytes)",
            record.offset,
            data.len().saturating_sub(rgce_start)
        ));
    }
    let rgce = &data[rgce_start..rgce_end];
    let Some((&ptg, payload)) = rgce.split_first() else {
        return Ok(None);
    };

    let kind = match ptg {
        0x01 => ExpKind::Exp,
        0x02 => ExpKind::Tbl,
        _ => return Ok(None),
    };

    // Canonical BIFF8 payload is 4 bytes => cce=5 including ptg.
    if rgce.len() == 5 {
        return Ok(None);
    }

    Ok(Some(PendingExp {
        kind,
        cell,
        payload: payload.to_vec(),
    }))
}

/// Decode formulas for BIFF8 cells whose `FORMULA.rgce` begins with `PtgExp` / `PtgTbl` and uses a
/// non-canonical payload width (i.e. `cce != 5`).
///
/// This is intended as a narrow robustness fallback when upstream decoders cannot resolve
/// wide-payload encodings back to the corresponding SHRFMLA/ARRAY/TABLE record.
pub(crate) fn parse_biff8_worksheet_ptgexp_formulas(
    workbook_stream: &[u8],
    start: usize,
    ctx: &rgce::RgceDecodeContext<'_>,
) -> Result<ParsedWorksheetExpFormulas, String> {
    let mut out = ParsedWorksheetExpFormulas::default();

    let allows_continuation = |id: u16| {
        id == RECORD_FORMULA || id == RECORD_SHRFMLA || id == RECORD_ARRAY || id == RECORD_TABLE
    };
    let mut iter =
        records::LogicalBiffRecordIter::from_offset(workbook_stream, start, allows_continuation)?;

    let mut shrfmla: HashMap<CellRef, Biff8ShrFmlaRecord> = HashMap::new();
    let mut shrfmla_analysis_by_base: HashMap<
        CellRef,
        Option<rgce::Biff8SharedFormulaRgceAnalysis>,
    > = HashMap::new();
    let mut array: HashMap<CellRef, Biff8ArrayRecord> = HashMap::new();
    let mut table: HashMap<CellRef, Biff8TableRecord> = HashMap::new();
    let mut pending: Vec<PendingExp> = Vec::new();

    while let Some(next) = iter.next() {
        let record = match next {
            Ok(r) => r,
            Err(err) => {
                warn_string(
                    &mut out.warnings,
                    format!("malformed BIFF record in worksheet stream: {err}"),
                );
                break;
            }
        };

        if record.offset != start && records::is_bof_record(record.record_id) {
            break;
        }
        if record.record_id == records::RECORD_EOF {
            break;
        }

        match record.record_id {
            RECORD_SHRFMLA => {
                match parse_biff8_shrfmla_record_with_rgcb(&record) {
                    Ok((rgce, rgcb)) => {
                        let Some(range) =
                            parse_shrfmla_range_best_effort(record.data.as_ref(), rgce.len())
                        else {
                            warn_string(
                                &mut out.warnings,
                                format!(
                                    "failed to parse SHRFMLA range at offset {} (len={})",
                                    record.offset,
                                    record.data.len()
                                ),
                            );
                            continue;
                        };
                        let anchor = range.0;
                        shrfmla.insert(anchor, Biff8ShrFmlaRecord { range, rgce, rgcb });
                    }
                    Err(err) => warn_string(
                        &mut out.warnings,
                        format!(
                            "failed to parse SHRFMLA record at offset {}: {err}",
                            record.offset
                        ),
                    ),
                };
            }
            RECORD_ARRAY => {
                let Some(range) = parse_ref_any_best_effort(record.data.as_ref()) else {
                    warn_string(
                        &mut out.warnings,
                        format!(
                            "failed to parse ARRAY range at offset {} (len={})",
                            record.offset,
                            record.data.len()
                        ),
                    );
                    continue;
                };
                let anchor = range.0;
                match parse_biff8_array_record(&record) {
                    Ok(parsed) => {
                        array.insert(
                            anchor,
                            Biff8ArrayRecord {
                                range,
                                rgce: parsed.rgce,
                                rgcb: parsed.rgcb,
                            },
                        );
                    }
                    Err(err) => warn_string(
                        &mut out.warnings,
                        format!(
                            "failed to parse ARRAY record at offset {}: {err}",
                            record.offset
                        ),
                    ),
                }
            }
            RECORD_TABLE => {
                let Some(range) = parse_ref_any_best_effort(record.data.as_ref()) else {
                    warn_string(
                        &mut out.warnings,
                        format!(
                            "failed to parse TABLE range at offset {} (len={})",
                            record.offset,
                            record.data.len()
                        ),
                    );
                    continue;
                };
                let anchor = range.0;
                table.insert(
                    anchor,
                    Biff8TableRecord {
                        range,
                        data: record.data.as_ref().to_vec(),
                    },
                );
            }
            RECORD_FORMULA => match parse_formula_record_for_wide_ptgexp(&record) {
                Ok(Some(p)) => pending.push(p),
                Ok(None) => {}
                Err(err) => warn_string(&mut out.warnings, err),
            },
            _ => {}
        }
    }

    if pending.is_empty() {
        return Ok(out);
    }

    for exp in pending {
        let candidates = ptgexp_candidates(&exp.payload);
        if candidates.is_empty() {
            warn_string(
                &mut out.warnings,
                format!(
                    "cell {}: {:?} token has no in-bounds coordinate candidates (payload_len={})",
                    exp.cell.to_a1(),
                    exp.kind,
                    exp.payload.len()
                ),
            );
            continue;
        }

        let mut matches: Vec<(u32, u32)> = Vec::new();
        for &(base_row, base_col) in &candidates {
            let base_cell = CellRef::new(base_row, base_col);

            let has_match = match exp.kind {
                ExpKind::Exp => {
                    shrfmla
                        .get(&base_cell)
                        .is_some_and(|d| range_contains_cell(d.range, exp.cell))
                        || array
                            .get(&base_cell)
                            .is_some_and(|d| range_contains_cell(d.range, exp.cell))
                        || table
                            .get(&base_cell)
                            .is_some_and(|d| range_contains_cell(d.range, exp.cell))
                }
                ExpKind::Tbl => {
                    table
                        .get(&base_cell)
                        .is_some_and(|d| range_contains_cell(d.range, exp.cell))
                        || shrfmla
                            .get(&base_cell)
                            .is_some_and(|d| range_contains_cell(d.range, exp.cell))
                        || array
                            .get(&base_cell)
                            .is_some_and(|d| range_contains_cell(d.range, exp.cell))
                }
            };

            if has_match {
                matches.push((base_row, base_col));
            }
        }

        let Some(&(base_row, base_col)) = matches.first() else {
            warn_string(
                &mut out.warnings,
                format!(
                    "cell {}: {:?} token base cell could not be resolved: candidates={candidates:?}",
                    exp.cell.to_a1(),
                    exp.kind
                ),
            );
            continue;
        };

        if matches.len() > 1 {
            warn_string(
                &mut out.warnings,
                format!(
                    "cell {}: {:?} token has multiple base-cell candidates that match definitions: {matches:?}; choosing ({base_row},{base_col})",
                    exp.cell.to_a1(),
                    exp.kind
                ),
            );
        }

        let base_cell = CellRef::new(base_row, base_col);
        let decoded = match exp.kind {
            ExpKind::Exp => {
                if let Some(def) = shrfmla
                    .get(&base_cell)
                    .filter(|d| range_contains_cell(d.range, exp.cell))
                {
                    let base_coord = rgce::CellCoord::new(base_row, base_col);
                    let target_coord = rgce::CellCoord::new(exp.cell.row, exp.cell.col);

                    let analysis = shrfmla_analysis_by_base
                        .entry(base_cell)
                        .or_insert_with(|| rgce::analyze_biff8_shared_formula_rgce(&def.rgce).ok());

                    let delta_is_zero = exp.cell == base_cell;
                    let needs_materialization = analysis
                        .as_ref()
                        .is_some_and(|analysis| !delta_is_zero && analysis.has_abs_refs_with_relative_flags)
                        // Best-effort fallback: if analysis failed, attempt materialization for
                        // follower cells and fall back on failure.
                        || (analysis.is_none() && !delta_is_zero);

                    let rgce_to_decode: std::borrow::Cow<'_, [u8]> = if needs_materialization {
                        match rgce::materialize_biff8_shared_formula_rgce(
                            &def.rgce,
                            base_coord,
                            target_coord,
                        ) {
                            Ok(v) => std::borrow::Cow::Owned(v),
                            Err(err) => {
                                out.warnings.push(format!(
                                    "cell {}: failed to materialize shared formula base {}: {err}",
                                    exp.cell.to_a1(),
                                    base_cell.to_a1()
                                ));
                                std::borrow::Cow::Borrowed(&def.rgce)
                            }
                        }
                    } else {
                        std::borrow::Cow::Borrowed(&def.rgce)
                    };

                    if def.rgcb.is_empty() {
                        rgce::decode_biff8_rgce_with_base(&rgce_to_decode, ctx, Some(target_coord))
                    } else {
                        rgce::decode_biff8_rgce_with_base_and_rgcb(
                            &rgce_to_decode,
                            &def.rgcb,
                            ctx,
                            Some(target_coord),
                        )
                    }
                } else if let Some(def) = array
                    .get(&base_cell)
                    .filter(|d| range_contains_cell(d.range, exp.cell))
                {
                    if def.rgcb.is_empty() {
                        rgce::decode_biff8_rgce_with_base(
                            &def.rgce,
                            ctx,
                            Some(rgce::CellCoord::new(def.range.0.row, def.range.0.col)),
                        )
                    } else {
                        rgce::decode_biff8_rgce_with_base_and_rgcb(
                            &def.rgce,
                            &def.rgcb,
                            ctx,
                            Some(rgce::CellCoord::new(def.range.0.row, def.range.0.col)),
                        )
                    }
                } else if table
                    .get(&base_cell)
                    .is_some_and(|d| range_contains_cell(d.range, exp.cell))
                {
                    warn_string(
                        &mut out.warnings,
                        format!(
                            "cell {}: TABLE formula decoding is not supported; rendering #UNKNOWN!",
                            exp.cell.to_a1()
                        ),
                    );
                    out.formulas.insert(exp.cell, "#UNKNOWN!".to_string());
                    continue;
                } else {
                    warn_string(
                        &mut out.warnings,
                        format!(
                            "cell {}: {:?} token resolution inconsistency: base=({base_row},{base_col}) candidates={candidates:?}",
                            exp.cell.to_a1(),
                            exp.kind
                        ),
                    );
                    continue;
                }
            }
            ExpKind::Tbl => {
                if table
                    .get(&base_cell)
                    .is_some_and(|d| range_contains_cell(d.range, exp.cell))
                {
                    warn_string(
                        &mut out.warnings,
                        format!(
                            "cell {}: TABLE formula decoding is not supported; rendering #UNKNOWN!",
                            exp.cell.to_a1()
                        ),
                    );
                    out.formulas.insert(exp.cell, "#UNKNOWN!".to_string());
                    continue;
                } else if let Some(def) = shrfmla
                    .get(&base_cell)
                    .filter(|d| range_contains_cell(d.range, exp.cell))
                {
                    if def.rgcb.is_empty() {
                        rgce::decode_biff8_rgce_with_base(
                            &def.rgce,
                            ctx,
                            Some(rgce::CellCoord::new(exp.cell.row, exp.cell.col)),
                        )
                    } else {
                        rgce::decode_biff8_rgce_with_base_and_rgcb(
                            &def.rgce,
                            &def.rgcb,
                            ctx,
                            Some(rgce::CellCoord::new(exp.cell.row, exp.cell.col)),
                        )
                    }
                } else if let Some(def) = array
                    .get(&base_cell)
                    .filter(|d| range_contains_cell(d.range, exp.cell))
                {
                    if def.rgcb.is_empty() {
                        rgce::decode_biff8_rgce_with_base(
                            &def.rgce,
                            ctx,
                            Some(rgce::CellCoord::new(def.range.0.row, def.range.0.col)),
                        )
                    } else {
                        rgce::decode_biff8_rgce_with_base_and_rgcb(
                            &def.rgce,
                            &def.rgcb,
                            ctx,
                            Some(rgce::CellCoord::new(def.range.0.row, def.range.0.col)),
                        )
                    }
                } else {
                    warn_string(
                        &mut out.warnings,
                        format!(
                            "cell {}: {:?} token resolution inconsistency: base=({base_row},{base_col}) candidates={candidates:?}",
                            exp.cell.to_a1(),
                            exp.kind
                        ),
                    );
                    continue;
                }
            }
        };

        for w in decoded.warnings {
            warn_string(&mut out.warnings, format!("cell {}: {w}", exp.cell.to_a1()));
        }
        out.formulas.insert(exp.cell, decoded.text);
    }

    Ok(out)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record(id: u16, payload: &[u8]) -> Vec<u8> {
        let mut out = Vec::with_capacity(4 + payload.len());
        out.extend_from_slice(&id.to_le_bytes());
        out.extend_from_slice(&(payload.len() as u16).to_le_bytes());
        out.extend_from_slice(payload);
        out
    }

    #[test]
    fn parses_shrfmla_record_without_cuse() {
        // Some producers omit the `cUse` field in the SHRFMLA header, yielding:
        //   RefU (6 bytes) + cce (u16) + rgce (cce bytes)
        let mut payload = Vec::new();
        // RefU: rwFirst=0, rwLast=0, colFirst=0, colLast=0.
        payload.extend_from_slice(&0u16.to_le_bytes());
        payload.extend_from_slice(&0u16.to_le_bytes());
        payload.push(0);
        payload.push(0);
        // cce=1, rgce=[0x03] (arbitrary 1-byte token with no payload).
        payload.extend_from_slice(&1u16.to_le_bytes());
        payload.push(0x03);

        let stream = record(RECORD_SHRFMLA, &payload);
        let allows_continuation = |id: u16| id == RECORD_SHRFMLA;
        let mut iter = records::LogicalBiffRecordIter::new(&stream, allows_continuation);
        let record = iter.next().expect("record").expect("logical record");
        assert_eq!(record.record_id, RECORD_SHRFMLA);

        let parsed = parse_biff8_shrfmla_record(&record).expect("parse shrfmla");
        assert_eq!(parsed.rgce, vec![0x03]);
    }

    #[test]
    fn formula_grbit_membership_hint_is_unambiguous_only_when_single_bit_set() {
        assert_eq!(FormulaGrbit(0).membership_hint(), None);
        assert_eq!(
            FormulaGrbit(FormulaGrbit::F_SHR_FMLA).membership_hint(),
            Some(FormulaMembershipHint::Shared)
        );
        assert_eq!(
            FormulaGrbit(FormulaGrbit::F_ARRAY).membership_hint(),
            Some(FormulaMembershipHint::Array)
        );
        assert_eq!(
            FormulaGrbit(FormulaGrbit::F_TBL).membership_hint(),
            Some(FormulaMembershipHint::Table)
        );

        // Ambiguous combinations should yield no hint.
        assert_eq!(
            FormulaGrbit(FormulaGrbit::F_SHR_FMLA | FormulaGrbit::F_ARRAY).membership_hint(),
            None
        );
        assert_eq!(
            FormulaGrbit(FormulaGrbit::F_SHR_FMLA | FormulaGrbit::F_TBL).membership_hint(),
            None
        );
        assert_eq!(
            FormulaGrbit(FormulaGrbit::F_ARRAY | FormulaGrbit::F_TBL).membership_hint(),
            None
        );
        assert_eq!(
            FormulaGrbit(FormulaGrbit::F_SHR_FMLA | FormulaGrbit::F_ARRAY | FormulaGrbit::F_TBL)
                .membership_hint(),
            None
        );
    }

    #[test]
    fn parses_formula_rgce_with_continued_ptgstr_token() {
        // Build a FORMULA record whose rgce contains a PtgStr token split across a CONTINUE
        // boundary. Excel inserts a 1-byte "continued segment" option flags prefix at the start of
        // the continued fragment; ensure we skip it so the recovered rgce matches the canonical
        // stream.
        let literal = "ABCDE";

        let rgce_expected: Vec<u8> = [
            vec![0x17, literal.len() as u8, 0u8], // PtgStr + cch + flags (compressed)
            literal.as_bytes().to_vec(),
        ]
        .concat();

        // Split after the first two characters ("AB"). The continued fragment begins with the
        // continued-segment option flags byte (fHighByte), then the remaining bytes.
        let first_rgce = &rgce_expected[..(3 + 2)]; // ptg + cch + flags + "AB"
        let remaining_chars = &literal.as_bytes()[2..]; // "CDE"
        let mut continue_payload = Vec::new();
        continue_payload.push(0); // continued segment option flags (compressed)
        continue_payload.extend_from_slice(remaining_chars);

        // Minimal BIFF8 FORMULA record header (matches `xls_fixture_builder::formula_cell`):
        // [row][col][xf][cached_result:f64][grbit][chn][cce][rgce]
        let row = 1u16;
        let col = 2u16;
        let xf = 3u16;
        let cached_result = 0f64;
        let cce = rgce_expected.len() as u16;

        let mut formula_payload_part1 = Vec::new();
        formula_payload_part1.extend_from_slice(&row.to_le_bytes());
        formula_payload_part1.extend_from_slice(&col.to_le_bytes());
        formula_payload_part1.extend_from_slice(&xf.to_le_bytes());
        formula_payload_part1.extend_from_slice(&cached_result.to_le_bytes());
        formula_payload_part1.extend_from_slice(&0u16.to_le_bytes()); // grbit
        formula_payload_part1.extend_from_slice(&0u32.to_le_bytes()); // chn
        formula_payload_part1.extend_from_slice(&cce.to_le_bytes());
        formula_payload_part1.extend_from_slice(first_rgce);

        let stream = [
            record(RECORD_FORMULA, &formula_payload_part1),
            record(records::RECORD_CONTINUE, &continue_payload),
        ]
        .concat();

        let allows_continuation = |id: u16| id == RECORD_FORMULA;
        let mut iter = records::LogicalBiffRecordIter::new(&stream, allows_continuation);
        let record = iter.next().expect("record").expect("logical record");
        assert_eq!(record.record_id, RECORD_FORMULA);
        assert!(record.is_continued());

        let parsed = parse_biff8_formula_record(&record).expect("parse formula");
        assert_eq!(parsed.row, row);
        assert_eq!(parsed.col, col);
        assert_eq!(parsed.xf, xf);
        assert_eq!(parsed.grbit, FormulaGrbit(0));
        assert_eq!(parsed.rgce, rgce_expected);
        assert!(parsed.rgcb.is_empty());
    }

    #[test]
    fn parses_formula_rgce_with_ptgexp_before_continued_ptgstr_token() {
        // Regression test: ensure we consume fixed-size payloads for common tokens (PtgExp/PtgTbl)
        // so the rgce stream stays aligned and we can still detect/skip PtgStr continuation flags.
        let literal = "ABCDE";
        let ptgexp_row = 0x1234u16;
        let ptgexp_col = 0x5678u16;

        let rgce_expected: Vec<u8> = [
            vec![0x01], // PtgExp
            ptgexp_row.to_le_bytes().to_vec(),
            ptgexp_col.to_le_bytes().to_vec(),
            vec![0x17, literal.len() as u8, 0u8], // PtgStr + cch + flags (compressed)
            literal.as_bytes().to_vec(),
        ]
        .concat();

        // Split after the first two characters ("AB") inside the string payload.
        let first_rgce = &rgce_expected[..(5 + 3 + 2)]; // PtgExp (5) + PtgStr header (3) + "AB" (2)
        let remaining_chars = &literal.as_bytes()[2..]; // "CDE"
        let mut continue_payload = Vec::new();
        continue_payload.push(0); // continued segment option flags (compressed)
        continue_payload.extend_from_slice(remaining_chars);

        let row = 1u16;
        let col = 2u16;
        let xf = 3u16;
        let cached_result = 0f64;
        let cce = rgce_expected.len() as u16;

        let mut formula_payload_part1 = Vec::new();
        formula_payload_part1.extend_from_slice(&row.to_le_bytes());
        formula_payload_part1.extend_from_slice(&col.to_le_bytes());
        formula_payload_part1.extend_from_slice(&xf.to_le_bytes());
        formula_payload_part1.extend_from_slice(&cached_result.to_le_bytes());
        formula_payload_part1.extend_from_slice(&0u16.to_le_bytes()); // grbit
        formula_payload_part1.extend_from_slice(&0u32.to_le_bytes()); // chn
        formula_payload_part1.extend_from_slice(&cce.to_le_bytes());
        formula_payload_part1.extend_from_slice(first_rgce);

        let stream = [
            record(RECORD_FORMULA, &formula_payload_part1),
            record(records::RECORD_CONTINUE, &continue_payload),
        ]
        .concat();

        let allows_continuation = |id: u16| id == RECORD_FORMULA;
        let mut iter = records::LogicalBiffRecordIter::new(&stream, allows_continuation);
        let record = iter.next().expect("record").expect("logical record");
        assert_eq!(record.record_id, RECORD_FORMULA);
        assert!(record.is_continued());

        let parsed = parse_biff8_formula_record(&record).expect("parse formula");
        assert_eq!(parsed.grbit, FormulaGrbit(0));
        assert_eq!(parsed.rgce, rgce_expected);
        assert!(parsed.rgcb.is_empty());
    }

    #[test]
    fn parses_formula_rgce_with_unicode_ptgstr_split_across_continue() {
        // Ensure we also handle the non-zero continued-segment option flags (fHighByte=1) used when
        // an uncompressed UTF-16LE ShortXLUnicodeString is continued into a CONTINUE record.
        let literal = "ABCDE";
        let utf16: Vec<u16> = literal.encode_utf16().collect();

        let mut rgce_expected = Vec::new();
        rgce_expected.push(0x17); // PtgStr
        rgce_expected.push(utf16.len() as u8); // cch
        rgce_expected.push(STR_FLAG_HIGH_BYTE); // flags (unicode)
        for unit in &utf16 {
            rgce_expected.extend_from_slice(&unit.to_le_bytes());
        }

        // Split after the first two UTF-16LE code units ("AB") => 4 bytes of character payload.
        let first_rgce_len = 3 + 4;
        let first_rgce = &rgce_expected[..first_rgce_len];
        let remaining_bytes = &rgce_expected[first_rgce_len..];

        let mut continue_payload = Vec::new();
        continue_payload.push(STR_FLAG_HIGH_BYTE); // continued segment option flags (unicode)
        continue_payload.extend_from_slice(remaining_bytes);

        let row = 0u16;
        let col = 0u16;
        let xf = 0u16;
        let cached_result = 0f64;
        let cce = rgce_expected.len() as u16;

        let mut formula_payload_part1 = Vec::new();
        formula_payload_part1.extend_from_slice(&row.to_le_bytes());
        formula_payload_part1.extend_from_slice(&col.to_le_bytes());
        formula_payload_part1.extend_from_slice(&xf.to_le_bytes());
        formula_payload_part1.extend_from_slice(&cached_result.to_le_bytes());
        formula_payload_part1.extend_from_slice(&0u16.to_le_bytes()); // grbit
        formula_payload_part1.extend_from_slice(&0u32.to_le_bytes()); // chn
        formula_payload_part1.extend_from_slice(&cce.to_le_bytes());
        formula_payload_part1.extend_from_slice(first_rgce);

        let stream = [
            record(RECORD_FORMULA, &formula_payload_part1),
            record(records::RECORD_CONTINUE, &continue_payload),
        ]
        .concat();

        let allows_continuation = |id: u16| id == RECORD_FORMULA;
        let mut iter = records::LogicalBiffRecordIter::new(&stream, allows_continuation);
        let record = iter.next().expect("record").expect("logical record");
        assert!(record.is_continued());

        let parsed = parse_biff8_formula_record(&record).expect("parse formula");
        assert_eq!(parsed.rgce, rgce_expected);
        assert!(parsed.rgcb.is_empty());
    }

    #[test]
    fn parses_formula_rgce_with_richtext_ptgstr_split_inside_char_bytes() {
        let literal = "ABCDE";
        let c_run = 1u16;
        let rg_run = [0x11, 0x22, 0x33, 0x44];

        let rgce_expected: Vec<u8> = [
            vec![0x17, literal.len() as u8, STR_FLAG_RICH_TEXT],
            c_run.to_le_bytes().to_vec(),
            literal.as_bytes().to_vec(),
            rg_run.to_vec(),
        ]
        .concat();

        // Split after the first two characters ("AB") inside the character bytes.
        let first_rgce_len = 3 + 2 + 2; // header + cRun + "AB"
        let first_rgce = &rgce_expected[..first_rgce_len];
        let remaining = &rgce_expected[first_rgce_len..];

        let mut continue_payload = Vec::new();
        continue_payload.push(0); // continued segment option flags (compressed)
        continue_payload.extend_from_slice(remaining);

        let row = 0u16;
        let col = 0u16;
        let xf = 0u16;
        let cached_result = 0f64;
        let cce = rgce_expected.len() as u16;

        let mut formula_payload_part1 = Vec::new();
        formula_payload_part1.extend_from_slice(&row.to_le_bytes());
        formula_payload_part1.extend_from_slice(&col.to_le_bytes());
        formula_payload_part1.extend_from_slice(&xf.to_le_bytes());
        formula_payload_part1.extend_from_slice(&cached_result.to_le_bytes());
        formula_payload_part1.extend_from_slice(&0u16.to_le_bytes()); // grbit
        formula_payload_part1.extend_from_slice(&0u32.to_le_bytes()); // chn
        formula_payload_part1.extend_from_slice(&cce.to_le_bytes());
        formula_payload_part1.extend_from_slice(first_rgce);

        let stream = [
            record(RECORD_FORMULA, &formula_payload_part1),
            record(records::RECORD_CONTINUE, &continue_payload),
        ]
        .concat();

        let allows_continuation = |id: u16| id == RECORD_FORMULA;
        let mut iter = records::LogicalBiffRecordIter::new(&stream, allows_continuation);
        let record = iter.next().expect("record").expect("logical record");
        assert!(record.is_continued());

        let parsed = parse_biff8_formula_record(&record).expect("parse formula");
        assert_eq!(parsed.rgce, rgce_expected);
    }

    #[test]
    fn parses_formula_rgce_with_richtext_ptgstr_split_between_crun_bytes() {
        let literal = "ABCDE";
        let c_run = 1u16;
        let rg_run = [0x11, 0x22, 0x33, 0x44];

        let rgce_expected: Vec<u8> = [
            vec![0x17, literal.len() as u8, STR_FLAG_RICH_TEXT],
            c_run.to_le_bytes().to_vec(),
            literal.as_bytes().to_vec(),
            rg_run.to_vec(),
        ]
        .concat();

        // Split between the two bytes of `cRun`.
        let first_rgce_len = 3 + 1; // header + low byte of cRun
        let first_rgce = &rgce_expected[..first_rgce_len];
        let remaining = &rgce_expected[first_rgce_len..];

        let mut continue_payload = Vec::new();
        continue_payload.push(0); // continued segment option flags (compressed)
        continue_payload.extend_from_slice(remaining);

        let row = 0u16;
        let col = 0u16;
        let xf = 0u16;
        let cached_result = 0f64;
        let cce = rgce_expected.len() as u16;

        let mut formula_payload_part1 = Vec::new();
        formula_payload_part1.extend_from_slice(&row.to_le_bytes());
        formula_payload_part1.extend_from_slice(&col.to_le_bytes());
        formula_payload_part1.extend_from_slice(&xf.to_le_bytes());
        formula_payload_part1.extend_from_slice(&cached_result.to_le_bytes());
        formula_payload_part1.extend_from_slice(&0u16.to_le_bytes()); // grbit
        formula_payload_part1.extend_from_slice(&0u32.to_le_bytes()); // chn
        formula_payload_part1.extend_from_slice(&cce.to_le_bytes());
        formula_payload_part1.extend_from_slice(first_rgce);

        let stream = [
            record(RECORD_FORMULA, &formula_payload_part1),
            record(records::RECORD_CONTINUE, &continue_payload),
        ]
        .concat();

        let allows_continuation = |id: u16| id == RECORD_FORMULA;
        let mut iter = records::LogicalBiffRecordIter::new(&stream, allows_continuation);
        let record = iter.next().expect("record").expect("logical record");
        assert!(record.is_continued());

        let parsed = parse_biff8_formula_record(&record).expect("parse formula");
        assert_eq!(parsed.rgce, rgce_expected);
    }

    #[test]
    fn parses_formula_rgce_with_ext_ptgstr_split_inside_ext_bytes() {
        let literal = "ABCDE";
        let ext = [0xDE, 0xAD, 0xBE, 0xEF];

        let rgce_expected: Vec<u8> = [
            vec![0x17, literal.len() as u8, STR_FLAG_EXT],
            (ext.len() as u32).to_le_bytes().to_vec(),
            literal.as_bytes().to_vec(),
            ext.to_vec(),
        ]
        .concat();

        // Split inside the ext bytes.
        let first_rgce_len = 3 + 4 + literal.len() + 2; // header + cbExtRst + chars + first 2 ext bytes
        let first_rgce = &rgce_expected[..first_rgce_len];
        let remaining = &rgce_expected[first_rgce_len..];

        let mut continue_payload = Vec::new();
        continue_payload.push(0); // continued segment option flags (compressed)
        continue_payload.extend_from_slice(remaining);

        let row = 0u16;
        let col = 0u16;
        let xf = 0u16;
        let cached_result = 0f64;
        let cce = rgce_expected.len() as u16;

        let mut formula_payload_part1 = Vec::new();
        formula_payload_part1.extend_from_slice(&row.to_le_bytes());
        formula_payload_part1.extend_from_slice(&col.to_le_bytes());
        formula_payload_part1.extend_from_slice(&xf.to_le_bytes());
        formula_payload_part1.extend_from_slice(&cached_result.to_le_bytes());
        formula_payload_part1.extend_from_slice(&0u16.to_le_bytes()); // grbit
        formula_payload_part1.extend_from_slice(&0u32.to_le_bytes()); // chn
        formula_payload_part1.extend_from_slice(&cce.to_le_bytes());
        formula_payload_part1.extend_from_slice(first_rgce);

        let stream = [
            record(RECORD_FORMULA, &formula_payload_part1),
            record(records::RECORD_CONTINUE, &continue_payload),
        ]
        .concat();

        let allows_continuation = |id: u16| id == RECORD_FORMULA;
        let mut iter = records::LogicalBiffRecordIter::new(&stream, allows_continuation);
        let record = iter.next().expect("record").expect("logical record");
        assert!(record.is_continued());

        let parsed = parse_biff8_formula_record(&record).expect("parse formula");
        assert_eq!(parsed.rgce, rgce_expected);
    }

    #[test]
    fn parses_formula_rgce_with_richtext_and_ext_ptgstr_split_inside_rgrun_bytes() {
        let literal = "ABCDE";
        let c_run = 1u16;
        let rg_run = [0x11, 0x22, 0x33, 0x44];
        let ext = [0x55, 0x66, 0x77, 0x88];

        let rgce_expected: Vec<u8> = [
            vec![0x17, literal.len() as u8, STR_FLAG_RICH_TEXT | STR_FLAG_EXT],
            c_run.to_le_bytes().to_vec(),
            (ext.len() as u32).to_le_bytes().to_vec(),
            literal.as_bytes().to_vec(),
            rg_run.to_vec(),
            ext.to_vec(),
        ]
        .concat();

        // Split inside the `rgRun` bytes.
        let first_rgce_len = 3 + 2 + 4 + literal.len() + 2; // header + cRun + cbExtRst + chars + first 2 rgRun bytes
        let first_rgce = &rgce_expected[..first_rgce_len];
        let remaining = &rgce_expected[first_rgce_len..];

        let mut continue_payload = Vec::new();
        continue_payload.push(0); // continued segment option flags (compressed)
        continue_payload.extend_from_slice(remaining);

        let row = 0u16;
        let col = 0u16;
        let xf = 0u16;
        let cached_result = 0f64;
        let cce = rgce_expected.len() as u16;

        let mut formula_payload_part1 = Vec::new();
        formula_payload_part1.extend_from_slice(&row.to_le_bytes());
        formula_payload_part1.extend_from_slice(&col.to_le_bytes());
        formula_payload_part1.extend_from_slice(&xf.to_le_bytes());
        formula_payload_part1.extend_from_slice(&cached_result.to_le_bytes());
        formula_payload_part1.extend_from_slice(&0u16.to_le_bytes()); // grbit
        formula_payload_part1.extend_from_slice(&0u32.to_le_bytes()); // chn
        formula_payload_part1.extend_from_slice(&cce.to_le_bytes());
        formula_payload_part1.extend_from_slice(first_rgce);

        let stream = [
            record(RECORD_FORMULA, &formula_payload_part1),
            record(records::RECORD_CONTINUE, &continue_payload),
        ]
        .concat();

        let allows_continuation = |id: u16| id == RECORD_FORMULA;
        let mut iter = records::LogicalBiffRecordIter::new(&stream, allows_continuation);
        let record = iter.next().expect("record").expect("logical record");
        assert!(record.is_continued());

        let parsed = parse_biff8_formula_record(&record).expect("parse formula");
        assert_eq!(parsed.rgce, rgce_expected);
    }

    #[test]
    fn unicode_ptgstr_errors_on_mid_code_unit_split() {
        // cch=1, unicode, but split after only 1 byte of the UTF-16LE code unit.
        let rgce_expected = vec![0x17, 1u8, STR_FLAG_HIGH_BYTE, b'A', 0x00];

        let first_rgce = &rgce_expected[..4]; // ptg + cch + flags + first byte of code unit
        let mut continue_payload = Vec::new();
        continue_payload.push(STR_FLAG_HIGH_BYTE); // continued segment option flags (unicode)
        continue_payload.push(0x00); // remaining byte of UTF-16LE code unit

        let row = 0u16;
        let col = 0u16;
        let xf = 0u16;
        let cached_result = 0f64;
        let cce = rgce_expected.len() as u16;

        let mut formula_payload_part1 = Vec::new();
        formula_payload_part1.extend_from_slice(&row.to_le_bytes());
        formula_payload_part1.extend_from_slice(&col.to_le_bytes());
        formula_payload_part1.extend_from_slice(&xf.to_le_bytes());
        formula_payload_part1.extend_from_slice(&cached_result.to_le_bytes());
        formula_payload_part1.extend_from_slice(&0u16.to_le_bytes()); // grbit
        formula_payload_part1.extend_from_slice(&0u32.to_le_bytes()); // chn
        formula_payload_part1.extend_from_slice(&cce.to_le_bytes());
        formula_payload_part1.extend_from_slice(first_rgce);

        let stream = [
            record(RECORD_FORMULA, &formula_payload_part1),
            record(records::RECORD_CONTINUE, &continue_payload),
        ]
        .concat();

        let allows_continuation = |id: u16| id == RECORD_FORMULA;
        let mut iter = records::LogicalBiffRecordIter::new(&stream, allows_continuation);
        let record = iter.next().expect("record").expect("logical record");
        assert!(record.is_continued());

        let err = parse_biff8_formula_record(&record).unwrap_err();
        assert_eq!(err, "string continuation split mid-character");
    }

    #[test]
    fn ptgstr_ext_size_out_of_bounds_errors_without_allocating_unbounded() {
        // Crafted PtgStr that sets fExtSt with a huge cbExtRst. The parser should not try to
        // allocate cbExtRst bytes; it should fail fast and return an error.
        let rgce = vec![0x17, 0u8, STR_FLAG_EXT, 0xFF, 0xFF, 0xFF, 0xFF];
        let payload = formula_payload(0, 0, 0, &rgce);
        let stream = record(RECORD_FORMULA, &payload);

        let allows_continuation = |id: u16| id == RECORD_FORMULA;
        let mut iter = records::LogicalBiffRecordIter::new(&stream, allows_continuation);
        let record = iter.next().expect("record").expect("logical record");

        let err = parse_biff8_formula_record(&record).unwrap_err();
        assert!(err.contains("PtgStr"), "err={err}");
    }

    #[test]
    fn worksheet_formulas_warns_and_continues_on_ptgstr_mid_code_unit_split() {
        // Ensure best-effort worksheet scan continues even if one FORMULA record contains a
        // malformed continued unicode PtgStr.
        //
        // Malformed formula: cch=1, unicode, but the first fragment only has 1 byte of the UTF-16LE
        // code unit.
        let bad_rgce = vec![0x17, 1u8, STR_FLAG_HIGH_BYTE, b'A', 0x00];
        let bad_first_rgce = &bad_rgce[..4];
        let mut bad_continue_payload = Vec::new();
        bad_continue_payload.push(STR_FLAG_HIGH_BYTE); // continued segment option flags (unicode)
        bad_continue_payload.push(0x00); // remaining byte of UTF-16LE code unit

        let bad_row = 0u16;
        let bad_col = 0u16;
        let bad_xf = 0u16;
        let cached_result = 0f64;
        let bad_cce = bad_rgce.len() as u16;

        let mut bad_formula_payload_part1 = Vec::new();
        bad_formula_payload_part1.extend_from_slice(&bad_row.to_le_bytes());
        bad_formula_payload_part1.extend_from_slice(&bad_col.to_le_bytes());
        bad_formula_payload_part1.extend_from_slice(&bad_xf.to_le_bytes());
        bad_formula_payload_part1.extend_from_slice(&cached_result.to_le_bytes());
        bad_formula_payload_part1.extend_from_slice(&0u16.to_le_bytes()); // grbit
        bad_formula_payload_part1.extend_from_slice(&0u32.to_le_bytes()); // chn
        bad_formula_payload_part1.extend_from_slice(&bad_cce.to_le_bytes());
        bad_formula_payload_part1.extend_from_slice(bad_first_rgce);

        // Valid formula after the malformed one.
        let good_rgce: Vec<u8> = vec![0x1E, 0x01, 0x00]; // PtgInt(1)
        let good_payload = formula_payload(0, 1, 0, &good_rgce);

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_FORMULA, &bad_formula_payload_part1),
            record(records::RECORD_CONTINUE, &bad_continue_payload),
            record(RECORD_FORMULA, &good_payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff8_worksheet_formulas(&stream, 0).expect("parse");
        assert!(
            parsed.warnings.iter().any(|w| w
                .message
                .contains("string continuation split mid-character")),
            "warnings={:?}",
            parsed.warnings
        );

        assert!(
            !parsed.formula_cells.contains_key(&CellRef::new(0, 0)),
            "malformed cell should be skipped"
        );
        let good_cell = parsed
            .formula_cells
            .get(&CellRef::new(0, 1))
            .expect("missing good formula cell");
        assert_eq!(good_cell.rgce, good_rgce);
    }

    #[test]
    fn parses_shrfmla_rgce_with_continued_ptgstr_token() {
        // Synthetic SHRFMLA record split across CONTINUE within a PtgStr payload.
        let literal = "ABCDE";
        let rgce_expected: Vec<u8> = [
            vec![0x17, literal.len() as u8, 0u8], // PtgStr + cch + flags (compressed)
            literal.as_bytes().to_vec(),
        ]
        .concat();

        let first_rgce = &rgce_expected[..(3 + 2)]; // ptg + cch + flags + "AB"
        let remaining_chars = &literal.as_bytes()[2..]; // "CDE"
        let mut continue_payload = Vec::new();
        continue_payload.push(0); // continued segment option flags (compressed)
        continue_payload.extend_from_slice(remaining_chars);

        // SHRFMLA record (best-effort): ref (RefU, 6 bytes) + cUse (2) + cce (2) + rgce.
        let rw_first = 0u16;
        let rw_last = 0u16;
        let col_first = 0u8;
        let col_last = 0u8;
        let c_use = 0u16;
        let cce = rgce_expected.len() as u16;

        let mut payload_part1 = Vec::new();
        payload_part1.extend_from_slice(&rw_first.to_le_bytes());
        payload_part1.extend_from_slice(&rw_last.to_le_bytes());
        payload_part1.push(col_first);
        payload_part1.push(col_last);
        payload_part1.extend_from_slice(&c_use.to_le_bytes());
        payload_part1.extend_from_slice(&cce.to_le_bytes());
        payload_part1.extend_from_slice(first_rgce);

        let stream = [
            record(RECORD_SHRFMLA, &payload_part1),
            record(records::RECORD_CONTINUE, &continue_payload),
        ]
        .concat();

        let allows_continuation = |id: u16| id == RECORD_SHRFMLA;
        let mut iter = records::LogicalBiffRecordIter::new(&stream, allows_continuation);
        let record = iter.next().expect("record").expect("logical record");
        assert_eq!(record.record_id, RECORD_SHRFMLA);
        assert!(record.is_continued());

        let parsed = parse_biff8_shrfmla_record(&record).expect("parse SHRFMLA");
        assert_eq!(parsed.rgce, rgce_expected);
        assert!(parsed.rgcb.is_empty());
    }

    #[test]
    fn parses_array_rgce_with_continued_ptgstr_token() {
        // Synthetic ARRAY record split across CONTINUE within a PtgStr payload.
        let literal = "ABCDE";
        let rgce_expected: Vec<u8> = [
            vec![0x17, literal.len() as u8, 0u8], // PtgStr + cch + flags (compressed)
            literal.as_bytes().to_vec(),
        ]
        .concat();

        let first_rgce = &rgce_expected[..(3 + 2)]; // ptg + cch + flags + "AB"
        let remaining_chars = &literal.as_bytes()[2..]; // "CDE"
        let mut continue_payload = Vec::new();
        continue_payload.push(0); // continued segment option flags (compressed)
        continue_payload.extend_from_slice(remaining_chars);

        // ARRAY record (best-effort): ref (RefU, 6 bytes) + reserved/grbit (2) + cce (2) + rgce.
        let rw_first = 0u16;
        let rw_last = 0u16;
        let col_first = 0u8;
        let col_last = 0u8;
        let grbit = 0u16;
        let cce = rgce_expected.len() as u16;

        let mut payload_part1 = Vec::new();
        payload_part1.extend_from_slice(&rw_first.to_le_bytes());
        payload_part1.extend_from_slice(&rw_last.to_le_bytes());
        payload_part1.push(col_first);
        payload_part1.push(col_last);
        payload_part1.extend_from_slice(&grbit.to_le_bytes());
        payload_part1.extend_from_slice(&cce.to_le_bytes());
        payload_part1.extend_from_slice(first_rgce);

        let stream = [
            record(RECORD_ARRAY, &payload_part1),
            record(records::RECORD_CONTINUE, &continue_payload),
        ]
        .concat();

        let allows_continuation = |id: u16| id == RECORD_ARRAY;
        let mut iter = records::LogicalBiffRecordIter::new(&stream, allows_continuation);
        let record = iter.next().expect("record").expect("logical record");
        assert_eq!(record.record_id, RECORD_ARRAY);
        assert!(record.is_continued());

        let parsed = parse_biff8_array_record(&record).expect("parse ARRAY");
        assert_eq!(parsed.rgce, rgce_expected);
        assert!(parsed.rgcb.is_empty());
    }

    #[test]
    fn parses_continued_formula_when_cce_and_rgce_cross_fragment_boundaries() {
        // FORMULA record split across CONTINUE such that the `cce` length field and the `rgce`
        // bytes both cross fragment boundaries.
        let rgce: Vec<u8> = vec![0x1E, 0x03, 0x00]; // PtgInt(3)
        let cce = rgce.len() as u16;

        // FORMULA header: row (2) + col (2) + xf (2) + cached result (8) + grbit (2) + chn (4) +
        // cce (2) + rgce...
        let mut formula_prefix = Vec::new();
        formula_prefix.extend_from_slice(&0u16.to_le_bytes()); // row
        formula_prefix.extend_from_slice(&0u16.to_le_bytes()); // col
        formula_prefix.extend_from_slice(&0u16.to_le_bytes()); // xf
        formula_prefix.extend_from_slice(&[0u8; 8]); // cached result (dummy)
        formula_prefix.extend_from_slice(&0u16.to_le_bytes()); // grbit (dummy)
        formula_prefix.extend_from_slice(&0u32.to_le_bytes()); // chn (dummy)

        let cce_bytes = cce.to_le_bytes();

        // Split so that:
        // - the first CONTINUE boundary occurs after the first byte of cce (so cce crosses),
        // - the second CONTINUE boundary splits the rgce bytes.
        let formula_frag1 = [formula_prefix, vec![cce_bytes[0]]].concat();
        let cont1 = vec![cce_bytes[1], rgce[0]];
        let cont2 = vec![rgce[1], rgce[2]];

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_FORMULA, &formula_frag1),
            record(records::RECORD_CONTINUE, &cont1),
            record(records::RECORD_CONTINUE, &cont2),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff8_worksheet_formulas(&stream, 0).expect("parse");
        assert!(
            parsed.warnings.is_empty(),
            "expected no warnings, got {:?}",
            parsed.warnings
        );

        let cell = CellRef::new(0, 0);
        let formula = parsed.formula_cells.get(&cell).expect("missing FORMULA");
        assert_eq!(formula.grbit, FormulaGrbit(0));
        assert_eq!(formula.rgce, rgce);
    }

    #[test]
    fn parses_continued_shrfmla_when_cce_and_rgce_cross_fragment_boundaries() {
        // SHRFMLA record split across CONTINUE such that the `cce` length field and the `rgce`
        // bytes both cross fragment boundaries.
        let rgce: Vec<u8> = vec![0x1E, 0x01, 0x00]; // PtgInt(1)
        let cce = rgce.len() as u16;

        // SHRFMLA header: RefU (6 bytes) + cUse (2 bytes) + cce (2 bytes) + rgce...
        let mut shrfmla_prefix = Vec::new();
        shrfmla_prefix.extend_from_slice(&0u16.to_le_bytes()); // rwFirst
        shrfmla_prefix.extend_from_slice(&0u16.to_le_bytes()); // rwLast
        shrfmla_prefix.push(0); // colFirst
        shrfmla_prefix.push(0); // colLast
        shrfmla_prefix.extend_from_slice(&0u16.to_le_bytes()); // cUse

        let cce_bytes = cce.to_le_bytes();

        // Split so that:
        // - the first CONTINUE boundary occurs after the first byte of cce (so cce crosses),
        // - the second CONTINUE boundary splits the rgce bytes.
        let shrfmla_frag1 = [shrfmla_prefix, vec![cce_bytes[0]]].concat();
        let cont1 = vec![cce_bytes[1], rgce[0]];
        let cont2 = vec![rgce[1], rgce[2]];

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_SHRFMLA, &shrfmla_frag1),
            record(records::RECORD_CONTINUE, &cont1),
            record(records::RECORD_CONTINUE, &cont2),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff8_worksheet_formulas(&stream, 0).expect("parse");
        assert!(
            parsed.warnings.is_empty(),
            "expected no warnings, got {:?}",
            parsed.warnings
        );

        let anchor = CellRef::new(0, 0);
        let shrfmla = parsed.shrfmla.get(&anchor).expect("missing SHRFMLA");
        assert_eq!(shrfmla.range, (anchor, anchor));
        assert_eq!(shrfmla.rgce, rgce);
        assert!(shrfmla.rgcb.is_empty());
    }

    #[test]
    fn parses_continued_array_when_cce_and_rgce_cross_fragment_boundaries() {
        // ARRAY record split across CONTINUE such that the `cce` length field and the `rgce` bytes
        // both cross fragment boundaries.
        let rgce: Vec<u8> = vec![0x1E, 0x02, 0x00]; // PtgInt(2)
        let cce = rgce.len() as u16;

        // ARRAY header: RefU (6 bytes) + reserved (2 bytes) + cce (2 bytes) + rgce...
        let mut array_prefix = Vec::new();
        array_prefix.extend_from_slice(&0u16.to_le_bytes()); // rwFirst
        array_prefix.extend_from_slice(&0u16.to_le_bytes()); // rwLast
        array_prefix.push(0); // colFirst
        array_prefix.push(0); // colLast
        array_prefix.extend_from_slice(&0u16.to_le_bytes()); // reserved

        let cce_bytes = cce.to_le_bytes();

        // Split so that:
        // - the first CONTINUE boundary occurs after the first byte of cce (so cce crosses),
        // - the second CONTINUE boundary splits the rgce bytes.
        let array_frag1 = [array_prefix, vec![cce_bytes[0]]].concat();
        let cont1 = vec![cce_bytes[1], rgce[0]];
        let cont2 = vec![rgce[1], rgce[2]];

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_ARRAY, &array_frag1),
            record(records::RECORD_CONTINUE, &cont1),
            record(records::RECORD_CONTINUE, &cont2),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff8_worksheet_formulas(&stream, 0).expect("parse");
        assert!(
            parsed.warnings.is_empty(),
            "expected no warnings, got {:?}",
            parsed.warnings
        );

        let anchor = CellRef::new(0, 0);
        let array = parsed.array.get(&anchor).expect("missing ARRAY");
        assert_eq!(array.range, (anchor, anchor));
        assert_eq!(array.rgce, rgce);
        assert!(array.rgcb.is_empty());
    }

    fn formula_payload(row: u16, col: u16, grbit: u16, rgce: &[u8]) -> Vec<u8> {
        let mut out = Vec::new();
        out.extend_from_slice(&row.to_le_bytes());
        out.extend_from_slice(&col.to_le_bytes());
        out.extend_from_slice(&0u16.to_le_bytes()); // ixfe
        out.extend_from_slice(&[0u8; 8]); // cached result (dummy)
        out.extend_from_slice(&grbit.to_le_bytes());
        out.extend_from_slice(&0u32.to_le_bytes()); // chn (dummy)
        out.extend_from_slice(&(rgce.len() as u16).to_le_bytes()); // cce
        out.extend_from_slice(rgce);
        out
    }

    fn shrfmla_payload(rw_first: u16, rw_last: u16, col_first: u16, col_last: u16) -> Vec<u8> {
        let rgce = vec![0x1E, 0x01, 0x00]; // PtgInt(1) (dummy)
        let mut out = Vec::new();
        out.extend_from_slice(&rw_first.to_le_bytes());
        out.extend_from_slice(&rw_last.to_le_bytes());
        out.extend_from_slice(&col_first.to_le_bytes());
        out.extend_from_slice(&col_last.to_le_bytes());
        out.extend_from_slice(&0u16.to_le_bytes()); // cUse (dummy)
        out.extend_from_slice(&(rgce.len() as u16).to_le_bytes());
        out.extend_from_slice(&rgce);
        out
    }

    fn array_payload(rw_first: u16, rw_last: u16, col_first: u16, col_last: u16) -> Vec<u8> {
        let rgce = vec![0x1E, 0x02, 0x00]; // PtgInt(2) (dummy)
        let mut out = Vec::new();
        out.extend_from_slice(&rw_first.to_le_bytes());
        out.extend_from_slice(&rw_last.to_le_bytes());
        out.extend_from_slice(&col_first.to_le_bytes());
        out.extend_from_slice(&col_last.to_le_bytes());
        out.extend_from_slice(&0u16.to_le_bytes()); // reserved
        out.extend_from_slice(&(rgce.len() as u16).to_le_bytes());
        out.extend_from_slice(&rgce);
        out
    }

    fn table_payload(rw_first: u16, rw_last: u16, col_first: u16, col_last: u16) -> Vec<u8> {
        let mut out = Vec::new();
        out.extend_from_slice(&rw_first.to_le_bytes());
        out.extend_from_slice(&rw_last.to_le_bytes());
        out.extend_from_slice(&col_first.to_le_bytes());
        out.extend_from_slice(&col_last.to_le_bytes());
        out
    }

    fn ptgexp(base_row: u16, base_col: u16) -> Vec<u8> {
        let mut out = Vec::new();
        out.push(0x01);
        out.extend_from_slice(&base_row.to_le_bytes());
        out.extend_from_slice(&base_col.to_le_bytes());
        out
    }

    fn ptgtbl(base_row: u16, base_col: u16) -> Vec<u8> {
        let mut out = Vec::new();
        out.push(0x02);
        out.extend_from_slice(&base_row.to_le_bytes());
        out.extend_from_slice(&base_col.to_le_bytes());
        out
    }

    #[test]
    fn resolves_ptgexp_ptgtbl_using_grbit_hints_and_warns_on_mismatch() {
        // Build a synthetic sheet substream with:
        // - SHRFMLA + ARRAY definitions anchored at A1 (0,0) (ambiguous without flags)
        // - TABLE definition anchored at A4 (3,0)
        // - Formula cells:
        //   A2: PtgExp(A1) + fShrFmla -> resolve as shared (SHRFMLA)
        //   A3: PtgExp(A1) + fArray  -> resolve as array (ARRAY)
        //   A4: PtgTbl(A4) + fTbl    -> resolve as table (TABLE)
        //   A5: PtgExp(A1) + fTbl    -> warn mismatch and missing TABLE(A1), then fall back best-effort
        let a1 = (0u16, 0u16);
        let a4 = (3u16, 0u16);

        let stream = [
            record(RECORD_SHRFMLA, &shrfmla_payload(0, 10, 0, 0)),
            record(RECORD_ARRAY, &array_payload(0, 10, 0, 0)),
            record(RECORD_TABLE, &table_payload(a4.0, a4.0, a4.1, a4.1)),
            record(
                RECORD_FORMULA,
                &formula_payload(1, 0, FormulaGrbit::F_SHR_FMLA, &ptgexp(a1.0, a1.1)),
            ),
            record(
                RECORD_FORMULA,
                &formula_payload(2, 0, FormulaGrbit::F_ARRAY, &ptgexp(a1.0, a1.1)),
            ),
            record(
                RECORD_FORMULA,
                &formula_payload(3, 0, FormulaGrbit::F_TBL, &ptgtbl(a4.0, a4.1)),
            ),
            record(
                RECORD_FORMULA,
                &formula_payload(4, 0, FormulaGrbit::F_TBL, &ptgexp(a1.0, a1.1)),
            ),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff8_worksheet_formulas(&stream, 0).expect("parse");

        let mut warnings = Vec::new();

        let a2 = parsed.formula_cells.get(&CellRef::new(1, 0)).unwrap();
        assert_eq!(
            resolve_ptgexp_or_ptgtbl_best_effort(&parsed, a2, &mut warnings),
            PtgReferenceResolution::Shared {
                base: CellRef::new(0, 0)
            }
        );

        let a3 = parsed.formula_cells.get(&CellRef::new(2, 0)).unwrap();
        assert_eq!(
            resolve_ptgexp_or_ptgtbl_best_effort(&parsed, a3, &mut warnings),
            PtgReferenceResolution::Array {
                base: CellRef::new(0, 0)
            }
        );

        let a4_cell = parsed.formula_cells.get(&CellRef::new(3, 0)).unwrap();
        assert_eq!(
            resolve_ptgexp_or_ptgtbl_best_effort(&parsed, a4_cell, &mut warnings),
            PtgReferenceResolution::Table {
                base: CellRef::new(3, 0)
            }
        );

        let a5 = parsed.formula_cells.get(&CellRef::new(4, 0)).unwrap();
        let res = resolve_ptgexp_or_ptgtbl_best_effort(&parsed, a5, &mut warnings);
        // Best-effort fallback: `PtgExp(A1)` can still resolve via SHRFMLA/ARRAY.
        assert_eq!(
            res,
            PtgReferenceResolution::Shared {
                base: CellRef::new(0, 0)
            }
        );

        let warning_text = warnings
            .iter()
            .map(|w| w.message.as_str())
            .collect::<Vec<_>>()
            .join("\n");
        assert!(
            warning_text.contains("grbit.fTbl set but rgce does not start with PtgTbl"),
            "expected mismatch warning, got:\n{warning_text}"
        );
        assert!(
            warning_text.contains("no TABLE record was found for base A1"),
            "expected missing-TABLE warning, got:\n{warning_text}"
        );
    }

    #[test]
    fn resolves_ptgexp_when_base_cell_is_not_range_anchor() {
        // Regression: some `.xls` producers point PtgExp at a non-anchor cell inside the shared
        // range. We should still resolve it by scanning ranges.
        //
        // Shared range: B1:B2 (anchor = B1).
        // Formula cell: B2 has PtgExp(B2) + fShrFmla.
        let stream = [
            record(RECORD_SHRFMLA, &shrfmla_payload(0, 1, 1, 1)),
            record(
                RECORD_FORMULA,
                &formula_payload(1, 1, FormulaGrbit::F_SHR_FMLA, &ptgexp(1, 1)),
            ),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff8_worksheet_formulas(&stream, 0).expect("parse");

        let mut warnings = Vec::new();
        let b2 = parsed.formula_cells.get(&CellRef::new(1, 1)).unwrap();
        assert_eq!(
            resolve_ptgexp_or_ptgtbl_best_effort(&parsed, b2, &mut warnings),
            PtgReferenceResolution::Shared {
                base: CellRef::new(0, 1)
            }
        );
        assert!(
            warnings.is_empty(),
            "expected no warnings, got: {warnings:?}"
        );
    }

    #[test]
    fn ptgexp_wide_payload_warnings_are_bounded() {
        // Emit many non-canonical PtgExp tokens whose payload is too short to contain any
        // in-bounds coordinate candidates, and ensure warnings are capped.
        let mut stream = Vec::new();
        for idx in 0..(MAX_WARNINGS_PER_SHEET + 25) {
            // PtgExp with a 1-byte payload (non-canonical; canonical BIFF8 is 4 bytes).
            let rgce = [0x01u8, 0x00u8];
            stream.extend_from_slice(&record(
                RECORD_FORMULA,
                &formula_payload(idx as u16, 0, 0, &rgce),
            ));
        }
        stream.extend_from_slice(&record(records::RECORD_EOF, &[]));

        let ctx = rgce::RgceDecodeContext {
            codepage: 1252,
            sheet_names: &[],
            externsheet: &[],
            supbooks: &[],
            defined_names: &[],
        };

        let parsed = parse_biff8_worksheet_ptgexp_formulas(&stream, 0, &ctx).expect("parse");
        assert_eq!(
            parsed.warnings.len(),
            MAX_WARNINGS_PER_SHEET + 1,
            "warnings={:?}",
            parsed.warnings
        );
        assert_eq!(
            parsed.warnings.last().map(|w| w.as_str()),
            Some(WARNINGS_SUPPRESSED_MESSAGE)
        );
    }
}

/// Decoder for BIFF8 worksheet formula `rgce` streams.
///
/// This wraps the BIFF8 `rgce` decoder while reusing the same workbook-global context construction
/// as defined-name decoding (SUPBOOK + EXTERNSHEET + ordered NAME metadata for `PtgName`).
pub(crate) struct WorksheetFormulaDecoder {
    tables: super::workbook_context::BiffWorkbookContextTables,
}

impl WorksheetFormulaDecoder {
    pub(crate) fn new(
        workbook_stream: &[u8],
        biff: super::BiffVersion,
        codepage: u16,
        sheet_names: &[String],
    ) -> Self {
        let tables = super::workbook_context::build_biff_workbook_context_tables(
            workbook_stream,
            biff,
            codepage,
            sheet_names,
        );
        Self { tables }
    }

    pub(crate) fn warnings(&self) -> &[String] {
        &self.tables.warnings
    }

    pub(crate) fn rgce_decode_context<'a>(
        &'a self,
        sheet_names: &'a [String],
    ) -> super::rgce::RgceDecodeContext<'a> {
        self.tables.rgce_decode_context(sheet_names)
    }

    pub(crate) fn decode_rgce(
        &self,
        rgce_bytes: &[u8],
        sheet_names: &[String],
        base: super::rgce::CellCoord,
    ) -> super::rgce::DecodeRgceResult {
        let ctx = self.tables.rgce_decode_context(sheet_names);
        super::rgce::decode_biff8_rgce_with_base(rgce_bytes, &ctx, Some(base))
    }
}
