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

use super::records;

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

        (count == 1).then_some(out.expect("count==1 implies out set"))
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
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct ParsedSharedFormulaRecord {
    pub(crate) rgce: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct ParsedArrayRecord {
    pub(crate) rgce: Vec<u8>,
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

    Ok(ParsedFormulaRecord {
        row,
        col,
        xf,
        grbit,
        rgce,
    })
}

pub(crate) fn parse_biff8_shrfmla_record(
    record: &records::LogicalBiffRecord<'_>,
) -> Result<ParsedSharedFormulaRecord, String> {
    let fragments: Vec<&[u8]> = record.fragments().collect();
    let cursor = FragmentCursor::new(&fragments, 0, 0);

    // SHRFMLA layouts vary slightly between producers (RefU vs Ref8 for the shared range). Try a
    // small set of plausible BIFF8 layouts.
    // Layout A: RefU (6) + cUse (2) + cce (2).
    let mut c = cursor.clone();
    if let Ok(rgce) = parse_shrfmla_with_refu(&mut c) {
        return Ok(ParsedSharedFormulaRecord { rgce });
    }
    // Layout B: Ref8 (8) + cUse (2) + cce (2).
    let mut c = cursor;
    if let Ok(rgce) = parse_shrfmla_with_ref8(&mut c) {
        return Ok(ParsedSharedFormulaRecord { rgce });
    }

    Err("unrecognized SHRFMLA record layout".to_string())
}

pub(crate) fn parse_biff8_array_record(
    record: &records::LogicalBiffRecord<'_>,
) -> Result<ParsedArrayRecord, String> {
    let fragments: Vec<&[u8]> = record.fragments().collect();
    let cursor = FragmentCursor::new(&fragments, 0, 0);

    // ARRAY layouts vary slightly (RefU vs Ref8). Try both.
    {
        let mut c = cursor.clone();
        if let Ok(rgce) = parse_array_with_refu(&mut c) {
            return Ok(ParsedArrayRecord { rgce });
        }
    }
    {
        let mut c = cursor;
        if let Ok(rgce) = parse_array_with_ref8(&mut c) {
            return Ok(ParsedArrayRecord { rgce });
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
}

/// Minimal parsed representation of a BIFF8 SHRFMLA record.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Biff8ShrFmlaRecord {
    pub(crate) range: (CellRef, CellRef),
    pub(crate) rgce: Vec<u8>,
}

/// Minimal parsed representation of a BIFF8 ARRAY record.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Biff8ArrayRecord {
    pub(crate) range: (CellRef, CellRef),
    pub(crate) rgce: Vec<u8>,
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

fn warn(warnings: &mut Vec<crate::ImportWarning>, msg: impl Into<String>) {
    warnings.push(crate::ImportWarning {
        message: msg.into(),
    })
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
    let mut iter = records::LogicalBiffRecordIter::from_offset(workbook_stream, start, allows_continuation)?;

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
                    let Some(cell) = parse_cell_ref_u16(parsed_formula.row, parsed_formula.col) else {
                        continue;
                    };
                    out.formula_cells.insert(
                        cell,
                        Biff8FormulaCell {
                            cell,
                            grbit: parsed_formula.grbit,
                            rgce: parsed_formula.rgce,
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
                let Some(range) = parse_ref_any_best_effort(record.data.as_ref()) else {
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
                match parse_biff8_shrfmla_record(&record) {
                    Ok(parsed) => {
                        out.shrfmla.insert(
                            anchor,
                            Biff8ShrFmlaRecord {
                                range,
                                rgce: parsed.rgce,
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
                }
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
                            },
                        );
                    }
                    Err(err) => warn(
                        &mut out.warnings,
                        format!("failed to parse ARRAY record at offset {}: {err}", record.offset),
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
    Exp { base: CellRef },
    /// `PtgTbl` (0x02): points at a data table definition.
    Tbl { base: CellRef },
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
    Shared { base: CellRef },
    Array { base: CellRef },
    Table { base: CellRef },
    /// `PtgExp`/`PtgTbl` present but no suitable backing record could be found.
    Unresolved,
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
                if parsed.shrfmla.contains_key(&base) {
                    return PtgReferenceResolution::Shared { base };
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
                if parsed.array.contains_key(&base) {
                    return PtgReferenceResolution::Array { base };
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
                if parsed.table.contains_key(&base) {
                    return PtgReferenceResolution::Table { base };
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

fn parse_shrfmla_with_refu(cursor: &mut FragmentCursor<'_>) -> Result<Vec<u8>, String> {
    // ref (rwFirst:u16, rwLast:u16, colFirst:u8, colLast:u8)
    cursor.skip_bytes(2 + 2 + 1 + 1)?;
    // cUse
    cursor.skip_bytes(2)?;
    let cce = cursor.read_u16_le()? as usize;
    cursor.read_biff8_rgce(cce)
}

fn parse_shrfmla_with_ref8(cursor: &mut FragmentCursor<'_>) -> Result<Vec<u8>, String> {
    // ref (rwFirst:u16, rwLast:u16, colFirst:u16, colLast:u16)
    cursor.skip_bytes(8)?;
    // cUse
    cursor.skip_bytes(2)?;
    let cce = cursor.read_u16_le()? as usize;
    cursor.read_biff8_rgce(cce)
}

fn parse_array_with_refu(cursor: &mut FragmentCursor<'_>) -> Result<Vec<u8>, String> {
    // ref (rwFirst:u16, rwLast:u16, colFirst:u8, colLast:u8)
    cursor.skip_bytes(2 + 2 + 1 + 1)?;
    // reserved
    cursor.skip_bytes(2)?;
    let cce = cursor.read_u16_le()? as usize;
    cursor.read_biff8_rgce(cce)
}

fn parse_array_with_ref8(cursor: &mut FragmentCursor<'_>) -> Result<Vec<u8>, String> {
    // ref (rwFirst:u16, rwLast:u16, colFirst:u16, colLast:u16)
    cursor.skip_bytes(8)?;
    // reserved
    cursor.skip_bytes(2)?;
    let cce = cursor.read_u16_le()? as usize;
    cursor.read_biff8_rgce(cce)
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
                    let bytes = self.read_bytes(4)?;
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

                    let richtext_runs = if (flags & STR_FLAG_RICH_TEXT) != 0 {
                        let v = self.read_u16_le()?;
                        out.extend_from_slice(&v.to_le_bytes());
                        v as usize
                    } else {
                        0
                    };

                    let ext_size = if (flags & STR_FLAG_EXT) != 0 {
                        let v = self.read_u32_le()?;
                        out.extend_from_slice(&v.to_le_bytes());
                        v as usize
                    } else {
                        0
                    };

                    let mut is_unicode = (flags & STR_FLAG_HIGH_BYTE) != 0;
                    let mut remaining_chars = cch;

                    while remaining_chars > 0 {
                        if self.remaining_in_fragment() == 0 {
                            self.advance_fragment()?;
                            // Continued-segment option flags byte (fHighByte).
                            let cont_flags = self.read_u8()?;
                            is_unicode = (cont_flags & STR_FLAG_HIGH_BYTE) != 0;
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
                        let bytes = self.read_exact_from_current(take_bytes)?;
                        out.extend_from_slice(bytes);
                        remaining_chars -= take_chars;
                    }

                    let richtext_bytes = richtext_runs
                        .checked_mul(4)
                        .ok_or_else(|| "rich text run count overflow".to_string())?;
                    if richtext_bytes + ext_size > 0 {
                        let extra = self.read_bytes(richtext_bytes + ext_size)?;
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
}

/// Decoder for BIFF8 worksheet formula `rgce` streams.
///
/// This wraps the BIFF8 `rgce` decoder while reusing the same workbook-global context construction
/// as defined-name decoding (SUPBOOK + EXTERNSHEET + ordered NAME metadata for `PtgName`).
pub(crate) struct WorksheetFormulaDecoder<'a> {
    tables: super::workbook_context::BiffWorkbookContextTables,
    sheet_names: &'a [String],
}

impl<'a> WorksheetFormulaDecoder<'a> {
    pub(crate) fn new(
        workbook_stream: &[u8],
        biff: super::BiffVersion,
        codepage: u16,
        sheet_names: &'a [String],
    ) -> Self {
        let tables = super::workbook_context::build_biff_workbook_context_tables(
            workbook_stream,
            biff,
            codepage,
            sheet_names,
        );
        Self { tables, sheet_names }
    }

    pub(crate) fn warnings(&self) -> &[String] {
        &self.tables.warnings
    }

    pub(crate) fn decode_rgce(
        &self,
        rgce_bytes: &[u8],
        base: super::rgce::CellCoord,
    ) -> super::rgce::DecodeRgceResult {
        let ctx = self.tables.rgce_decode_context(self.sheet_names);
        super::rgce::decode_biff8_rgce_with_base(rgce_bytes, &ctx, Some(base))
    }
}
