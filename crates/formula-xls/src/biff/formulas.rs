//! BIFF8 worksheet formula helpers.
//!
//! This module contains best-effort BIFF8 worksheet-formula fallbacks/overrides used by the
//! `.xls` importer:
//! some non-standard producers emit follower-cell `FORMULA` records whose `rgce` consists of a
//! single `PtgExp` token (pointing at a "base" cell), but omit the corresponding `SHRFMLA`/`ARRAY`
//! definition record.
//!
//! In this case we attempt to recover a per-cell `rgce` by materializing the base cell's own
//! `FORMULA.rgce` token stream (when present and non-`PtgExp`) across the row/col delta.
//!
//! The materializer only needs to adjust BIFF8 reference ptgs that embed absolute coordinates plus
//! relative flags:
//! - `PtgRef` / `PtgArea` (and class variants)
//! - `PtgRef3d` / `PtgArea3d` (and class variants)
//!
//! Relative-offset ptgs (`PtgRefN` / `PtgAreaN`) are copied verbatim because they are interpreted
//! relative to the *current* formula cell at decode time.
//!
//! Additionally, calamine’s `.xls` formula decoder can mis-handle relative flags in 3D area
//! references (`PtgArea3d`) by treating the high flag bits of the column fields as part of the
//! column index. When we detect these patterns, we re-decode formulas directly from the BIFF
//! worksheet stream and override calamine’s string output.

use std::borrow::Cow;
use std::collections::{HashMap, HashSet};

use formula_model::{CellRef, EXCEL_MAX_COLS, EXCEL_MAX_ROWS};

use super::{records, rgce, worksheet_formulas};
use super::worksheet_formulas::FormulaMembershipHint;

// BIFF8 limits.
const BIFF8_MAX_ROW0: i64 = u16::MAX as i64; // 0..=65535
const BIFF8_MAX_COL0: i64 = 0x3FFF; // 14-bit field (some writers use 0x3FFF sentinels)

const COL_INDEX_MASK: u16 = 0x3FFF;
const ROW_RELATIVE_BIT: u16 = 0x4000;
const COL_RELATIVE_BIT: u16 = 0x8000;
const RELATIVE_MASK: u16 = 0xC000;

/// Cap warnings collected during best-effort shared/array formula recovery so a crafted `.xls`
/// cannot allocate an unbounded number of warning strings.
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

#[derive(Debug, Default)]
pub(crate) struct PtgExpFallbackResult {
    /// Recovered formulas (`CellRef` -> formula text without a leading `=`).
    pub(crate) formulas: HashMap<CellRef, String>,
    pub(crate) warnings: Vec<String>,
}

#[derive(Debug, Default)]
pub(crate) struct SheetFormulaOverrides {
    pub(crate) formulas: HashMap<CellRef, String>,
    pub(crate) warnings: Vec<String>,
}

#[derive(Clone, Copy)]
struct RangeMatch<'a, T> {
    anchor: CellRef,
    range: (CellRef, CellRef),
    record: &'a T,
}

fn format_range(range: (CellRef, CellRef)) -> String {
    let (start, end) = range;
    if start == end {
        start.to_a1()
    } else {
        format!("{}:{}", start.to_a1(), end.to_a1())
    }
}

fn format_cell_list(cells: &[CellRef]) -> String {
    let mut out = String::new();
    for (i, cell) in cells.iter().enumerate() {
        if i > 0 {
            out.push_str(", ");
        }
        out.push_str(&cell.to_a1());
    }
    out
}

fn range_area(range: (CellRef, CellRef)) -> u64 {
    let (start, end) = range;
    let rows = end.row.saturating_sub(start.row).saturating_add(1) as u64;
    let cols = end.col.saturating_sub(start.col).saturating_add(1) as u64;
    rows.saturating_mul(cols)
}

fn manhattan_distance(a: CellRef, b: CellRef) -> u64 {
    (a.row.abs_diff(b.row) as u64).saturating_add(a.col.abs_diff(b.col) as u64)
}

fn choose_best_range_match<'a, T>(
    matches: &[RangeMatch<'a, T>],
    current: CellRef,
    master_candidates: &[CellRef],
) -> usize {
    // Deterministic tie-breaking:
    // 1) smallest range area
    // 2) closest range start to master candidate (or current cell if no master candidates)
    // 3) range start/end ordering
    let mut best_idx: Option<usize> = None;
    let mut best_key: Option<(u64, u64, u32, u32, u32, u32)> = None;

    for (i, m) in matches.iter().enumerate() {
        let area = range_area(m.range);
        let dist = if master_candidates.is_empty() {
            manhattan_distance(m.range.0, current)
        } else {
            master_candidates
                .iter()
                .map(|&c| manhattan_distance(m.range.0, c))
                .min()
                .unwrap_or_else(|| manhattan_distance(m.range.0, current))
        };
        let key = (
            area,
            dist,
            m.range.0.row,
            m.range.0.col,
            m.range.1.row,
            m.range.1.col,
        );

        match best_key.as_ref() {
            None => {
                best_key = Some(key);
                best_idx = Some(i);
            }
            Some(best) if key < *best => {
                best_key = Some(key);
                best_idx = Some(i);
            }
            _ => {}
        }
    }

    best_idx.unwrap_or(0)
}

fn select_ptgexp_backing_record<'a, T>(
    records: &'a HashMap<CellRef, T>,
    current_cell: CellRef,
    master_candidates: &[CellRef],
    warnings: &mut Vec<String>,
    kind: &str,
    get_range: fn(&T) -> (CellRef, CellRef),
) -> Option<RangeMatch<'a, T>> {
    // (a) Exact-key match on master cell when the producer uses it as the record key.
    let mut exact_matches: Vec<RangeMatch<'a, T>> = Vec::new();
    for &master in master_candidates {
        if let Some(record) = records.get(&master) {
            let range = get_range(record);
            if range_contains(range, current_cell) {
                exact_matches.push(RangeMatch {
                    anchor: master,
                    range,
                    record,
                });
            }
        }
    }
    if !exact_matches.is_empty() {
        let exact_len = exact_matches.len();
        let best_idx = choose_best_range_match(&exact_matches, current_cell, master_candidates);
        let selected = exact_matches.swap_remove(best_idx);
        if exact_len > 1 {
            push_warning_bounded(
                warnings,
                format!(
                    "ambiguous {kind} match for PtgExp in {} (masters: {}): {} records keyed by master; chose range {}",
                    current_cell.to_a1(),
                    format_cell_list(master_candidates),
                    exact_len,
                    format_range(selected.range),
                ),
            );
        }
        return Some(selected);
    }

    // (b) Range containment match: current cell and referenced master are both inside the record's range.
    if !master_candidates.is_empty() {
        let mut matches: Vec<RangeMatch<'a, T>> = Vec::new();
        for (&anchor, record) in records {
            let range = get_range(record);
            if !range_contains(range, current_cell) {
                continue;
            }
            if master_candidates.iter().any(|&m| range_contains(range, m)) {
                matches.push(RangeMatch {
                    anchor,
                    range,
                    record,
                });
            }
        }

        if matches.len() == 1 {
            return matches.pop();
        }
        if matches.len() > 1 {
            let match_len = matches.len();
            let best_idx = choose_best_range_match(&matches, current_cell, master_candidates);
            let selected = matches.swap_remove(best_idx);
            push_warning_bounded(
                warnings,
                format!(
                    "ambiguous {kind} match for PtgExp in {} (masters: {}): matched {} ranges; chose range {}",
                    current_cell.to_a1(),
                    format_cell_list(master_candidates),
                    match_len,
                    format_range(selected.range),
                ),
            );
            return Some(selected);
        }
    }

    // (d) Fallback: ignore master cell; use current-cell containment only.
    let mut matches: Vec<RangeMatch<'a, T>> = Vec::new();
    for (&anchor, record) in records {
        let range = get_range(record);
        if range_contains(range, current_cell) {
            matches.push(RangeMatch {
                anchor,
                range,
                record,
            });
        }
    }

    if matches.is_empty() {
        return None;
    }
    if matches.len() == 1 {
        return matches.pop();
    }

    let match_len = matches.len();
    let best_idx = choose_best_range_match(&matches, current_cell, master_candidates);
    let selected = matches.swap_remove(best_idx);
    push_warning_bounded(
        warnings,
        format!(
            "ambiguous {kind} match for PtgExp in {}: {} ranges contain the cell (no usable master); chose range {}",
            current_cell.to_a1(),
            match_len,
            format_range(selected.range),
        ),
    );
    Some(selected)
}

/// Recover shared/array formulas referenced via `PtgExp` by resolving their corresponding
/// `SHRFMLA`/`ARRAY` definition records.
///
/// This is best-effort: malformed records are skipped with warnings.
pub(crate) fn recover_ptgexp_formulas_from_shrfmla_and_array(
    workbook_stream: &[u8],
    sheet_offset: usize,
    ctx: &rgce::RgceDecodeContext<'_>,
) -> Result<PtgExpFallbackResult, String> {
    let parsed = worksheet_formulas::parse_biff8_worksheet_formulas(workbook_stream, sheet_offset)?;

    let mut warnings: Vec<String> = parsed.warnings.into_iter().map(|w| w.message).collect();

    // Track all FORMULA cells so array formulas can be applied to every cell in the group range.
    let formula_cells: Vec<CellRef> = parsed.formula_cells.keys().copied().collect();

    // Track FORMULA cells whose `rgce` is `PtgExp`, so we can resolve them after scanning.
    let mut ptgexp_cells: Vec<(CellRef, Vec<CellRef>)> = Vec::new();
    for (&cell_ref, cell) in &parsed.formula_cells {
        let Some((base_row, base_col, candidates)) = parse_ptg_exp_master_cell_candidates(&cell.rgce)
        else {
            continue;
        };
        if candidates.is_empty() {
            push_warning_bounded(
                &mut warnings,
                format!(
                "skipping out-of-bounds PtgExp reference in {} -> ({base_row},{base_col})",
                cell_ref.to_a1()
                ),
            );
            continue;
        }
        ptgexp_cells.push((cell_ref, candidates));
    }

    let mut recovered: HashMap<CellRef, String> = HashMap::new();
    let mut shrfmla_analysis_by_base: HashMap<CellRef, Option<rgce::Biff8SharedFormulaRgceAnalysis>> =
        HashMap::new();
    let mut decoded_arrays: HashMap<CellRef, String> = HashMap::new();
    let mut applied_arrays: HashSet<CellRef> = HashSet::new();

    for (cell, base_candidates) in ptgexp_cells {
        // Shared formulas are decoded relative to the *current* cell.
        if let Some(selected) = select_ptgexp_backing_record(
            &parsed.shrfmla,
            cell,
            &base_candidates,
            &mut warnings,
            "SHRFMLA",
            |r| r.range,
        ) {
            let anchor = selected.anchor;
            let def = selected.record;

            let base_cell = rgce::CellCoord::new(anchor.row, anchor.col);
            let target_cell = rgce::CellCoord::new(cell.row, cell.col);

            let analysis = shrfmla_analysis_by_base
                .entry(anchor)
                .or_insert_with(|| rgce::analyze_biff8_shared_formula_rgce(&def.rgce).ok());

            let delta_is_zero = cell == anchor;
            let needs_materialization = analysis
                .as_ref()
                .is_some_and(|analysis| {
                    !analysis.has_refn_or_arean
                        || (!delta_is_zero && analysis.has_abs_refs_with_relative_flags)
                })
                // Best-effort fallback: if analysis failed or was inconclusive, attempt
                // materialization for follower cells (delta != 0) and fall back on failure.
                || (analysis.is_none() && !delta_is_zero);

            let rgce_to_decode: Cow<'_, [u8]> = if needs_materialization {
                match rgce::materialize_biff8_shared_formula_rgce(&def.rgce, base_cell, target_cell)
                {
                    Ok(v) => Cow::Owned(v),
                    Err(err) => {
                        push_warning_bounded(
                            &mut warnings,
                            format!(
                                "failed to materialize shared formula at {} (base {}, range {}): {err}",
                                cell.to_a1(),
                                anchor.to_a1(),
                                format_range(selected.range),
                            ),
                        );
                        Cow::Borrowed(&def.rgce)
                    }
                }
            } else {
                Cow::Borrowed(&def.rgce)
            };

            let decoded = rgce::decode_biff8_rgce_with_base_and_rgcb(
                &rgce_to_decode,
                &def.rgcb,
                ctx,
                Some(target_cell),
            );
            for w in decoded.warnings {
                push_warning_bounded(
                    &mut warnings,
                    format!(
                        "failed to fully decode shared formula at {} (range {}): {w}",
                        cell.to_a1(),
                        format_range(selected.range),
                    ),
                );
            }
            if !decoded.text.trim().is_empty() {
                recovered.insert(cell, decoded.text);
            }
            continue;
        }

        // Array formulas are decoded relative to the *array base* cell, and the same formula text
        // is displayed for every cell in the group.
        if let Some(selected) = select_ptgexp_backing_record(
            &parsed.array,
            cell,
            &base_candidates,
            &mut warnings,
            "ARRAY",
            |r| r.range,
        ) {
            let anchor = selected.anchor;
            let def = selected.record;
            let text = if let Some(existing) = decoded_arrays.get(&anchor) {
                existing.clone()
            } else {
                let base_coord = rgce::CellCoord::new(anchor.row, anchor.col);
                let decoded = rgce::decode_biff8_rgce_with_base_and_rgcb(
                    &def.rgce,
                    &def.rgcb,
                    ctx,
                    Some(base_coord),
                );
                for w in decoded.warnings {
                    push_warning_bounded(
                        &mut warnings,
                        format!("failed to fully decode array formula base {}: {w}", anchor.to_a1()),
                    );
                }
                let text = decoded.text;
                decoded_arrays.insert(anchor, text.clone());
                text
            };

            if text.trim().is_empty() {
                continue;
            }

            if applied_arrays.insert(anchor) {
                for &target in &formula_cells {
                    if range_contains(def.range, target) {
                        recovered.insert(target, text.clone());
                    }
                }
            }
            continue;
        }

        push_warning_bounded(
            &mut warnings,
            format!(
                "unresolved PtgExp reference in {} -> {}",
                cell.to_a1(),
                format_cell_list(&base_candidates)
            ),
        );
    }

    Ok(PtgExpFallbackResult {
        formulas: recovered,
        warnings,
    })
}

/// Best-effort recovery of `PtgExp`-only formulas by materializing from the referenced base cell's
/// own `FORMULA.rgce`.
///
/// This is intended as a fallback for malformed BIFF8 sheets where `SHRFMLA` / `ARRAY` records are
/// missing or corrupt.
pub(crate) fn recover_ptgexp_formulas_from_base_cell(
    workbook_stream: &[u8],
    sheet_offset: usize,
    ctx: &rgce::RgceDecodeContext<'_>,
) -> Result<PtgExpFallbackResult, String> {
    let allows_continuation = |id: u16| {
        id == worksheet_formulas::RECORD_FORMULA
            || id == worksheet_formulas::RECORD_SHRFMLA
            || id == worksheet_formulas::RECORD_ARRAY
    };
    let mut iter = records::LogicalBiffRecordIter::from_offset(
        workbook_stream,
        sheet_offset,
        allows_continuation,
    )?;

    // Collect all cell formula rgce bytes first so PtgExp followers can reference bases that
    // appear later in the stream.
    let mut rgce_by_cell: HashMap<(u32, u32), (Vec<u8>, Vec<u8>)> = HashMap::new();
    let mut shrfmla_by_cell: HashMap<(u32, u32), (Vec<u8>, Vec<u8>)> = HashMap::new();
    let mut array_ranges_by_cell: HashMap<(u32, u32), (CellRef, CellRef)> = HashMap::new();
    let mut ptgexp_cells: Vec<(u32, u32, u32, u32, worksheet_formulas::FormulaGrbit)> = Vec::new();
    let mut warnings: Vec<String> = Vec::new();

    // Ref8 columns can carry flags in their high bits; mask down to the 14-bit payload.
    const REF8_COL_MASK: u16 = 0x3FFF;
    let parse_ref_range = |data: &[u8]| -> Option<(CellRef, CellRef)> {
        // Prefer Ref8 when it decodes to classic `.xls` column bounds (<=255). Some producers store
        // Ref8 even when RefU would suffice.
        if let Some(chunk) = data.get(0..8) {
            let rw_first = u16::from_le_bytes([chunk[0], chunk[1]]);
            let rw_last = u16::from_le_bytes([chunk[2], chunk[3]]);
            let col_first_raw = u16::from_le_bytes([chunk[4], chunk[5]]);
            let col_last_raw = u16::from_le_bytes([chunk[6], chunk[7]]);
            let col_first = col_first_raw & REF8_COL_MASK;
            let col_last = col_last_raw & REF8_COL_MASK;
            if rw_first <= rw_last && col_first <= col_last && col_first <= 0x00FF && col_last <= 0x00FF
            {
                return Some((
                    CellRef::new(rw_first as u32, col_first as u32),
                    CellRef::new(rw_last as u32, col_last as u32),
                ));
            }
        }

        if let Some(chunk) = data.get(0..6) {
            let rw_first = u16::from_le_bytes([chunk[0], chunk[1]]);
            let rw_last = u16::from_le_bytes([chunk[2], chunk[3]]);
            let col_first = chunk[4] as u16;
            let col_last = chunk[5] as u16;
            if rw_first <= rw_last && col_first <= col_last {
                return Some((
                    CellRef::new(rw_first as u32, col_first as u32),
                    CellRef::new(rw_last as u32, col_last as u32),
                ));
            }
        }

        if let Some(chunk) = data.get(0..8) {
            let rw_first = u16::from_le_bytes([chunk[0], chunk[1]]);
            let rw_last = u16::from_le_bytes([chunk[2], chunk[3]]);
            let col_first_raw = u16::from_le_bytes([chunk[4], chunk[5]]);
            let col_last_raw = u16::from_le_bytes([chunk[6], chunk[7]]);
            let col_first = col_first_raw & REF8_COL_MASK;
            let col_last = col_last_raw & REF8_COL_MASK;
            if rw_first <= rw_last && col_first <= col_last {
                return Some((
                    CellRef::new(rw_first as u32, col_first as u32),
                    CellRef::new(rw_last as u32, col_last as u32),
                ));
            }
        }

        None
    };

    while let Some(next) = iter.next() {
        let record = match next {
            Ok(r) => r,
            Err(err) => {
                push_warning_bounded(&mut warnings, format!("malformed BIFF record in worksheet stream: {err}"));
                break;
            }
        };

        if record.offset != sheet_offset && records::is_bof_record(record.record_id) {
            break;
        }
        if record.record_id == records::RECORD_EOF {
            break;
        }
        match record.record_id {
            worksheet_formulas::RECORD_FORMULA => {
                let parsed = match worksheet_formulas::parse_biff8_formula_record(&record) {
                    Ok(parsed) => parsed,
                    Err(err) => {
                        push_warning_bounded(
                            &mut warnings,
                            format!(
                                "failed to parse FORMULA record at offset {} in worksheet stream: {err}",
                                record.offset
                            ),
                        );
                        continue;
                    }
                };
                let row = parsed.row as u32;
                let col = parsed.col as u32;
                let grbit = parsed.grbit;
                let rgce = parsed.rgce;
                let rgcb = parsed.rgcb;

                if let Some((base_row, base_col)) = parse_ptg_exp(&rgce) {
                    ptgexp_cells.push((row, col, base_row, base_col, grbit));
                }
                rgce_by_cell.insert((row, col), (rgce, rgcb));
            }
            worksheet_formulas::RECORD_SHRFMLA => {
                let Some((start, _end)) = parse_ref_range(record.data.as_ref()) else {
                    push_warning_bounded(
                        &mut warnings,
                        format!(
                            "failed to parse SHRFMLA range header at offset {} (len={})",
                            record.offset,
                            record.data.len()
                        ),
                    );
                    continue;
                };
                let row = start.row;
                let col = start.col;

                let parsed = match worksheet_formulas::parse_biff8_shrfmla_record(&record) {
                    Ok(parsed) => parsed,
                    Err(err) => {
                        push_warning_bounded(
                            &mut warnings,
                            format!(
                                "failed to parse SHRFMLA record at offset {} in worksheet stream: {err}",
                                record.offset
                            ),
                        );
                        continue;
                    }
                };
                shrfmla_by_cell.insert((row, col), (parsed.rgce, parsed.rgcb));
            }
            worksheet_formulas::RECORD_ARRAY => {
                let Some((start, end)) = parse_ref_range(record.data.as_ref()) else {
                    // Best-effort: ARRAY records vary between producers; if we cannot parse the range
                    // header, treat it as unknown and fall back on warnings for unresolved PtgExp.
                    continue;
                };
                array_ranges_by_cell.insert((start.row, start.col), (start, end));
            }
            _ => {}
        }
    }

    let mut recovered: HashMap<CellRef, String> = HashMap::new();

    for (row, col, base_row, base_col, grbit) in ptgexp_cells {
        let cell_ref = CellRef::new(row, col);
        let base_cell_ref = CellRef::new(base_row, base_col);
        let Some((base_rgce, base_rgcb)) = rgce_by_cell.get(&(base_row, base_col)) else {
            push_warning_bounded(
                &mut warnings,
                format!(
                    "failed to recover shared formula at {}: base cell ({},{}) has no FORMULA record",
                    cell_ref.to_a1(),
                    base_row,
                    base_col
                ),
            );
            continue;
        };

        let mut base_rgce_bytes: &[u8] = base_rgce;
        let mut base_rgcb_bytes: &[u8] = base_rgcb;
        if base_rgce_bytes.first().copied() == Some(0x01) {
            // Base cell stores PtgExp; attempt to resolve via SHRFMLA before giving up.
            if let Some((shared_rgce, shared_rgcb)) = shrfmla_by_cell.get(&(base_row, base_col)) {
                base_rgce_bytes = shared_rgce;
                base_rgcb_bytes = shared_rgcb;
            } else {
                // If this cell is part of a well-formed ARRAY formula group, the base cell's rgce is
                // also `PtgExp` and the real formula text lives in an `ARRAY` record. The base-cell
                // fallback cannot recover that rgce, so suppress the misleading "missing ARRAY"
                // warning.
                if array_ranges_by_cell
                    .get(&(base_row, base_col))
                    .is_some_and(|&range| range_contains(range, cell_ref))
                {
                    continue;
                }

                let expected = match grbit.membership_hint() {
                    Some(FormulaMembershipHint::Shared) => "missing SHRFMLA definition",
                    Some(FormulaMembershipHint::Array) => "missing ARRAY definition",
                    Some(FormulaMembershipHint::Table) => {
                        "unexpected fTbl set (expected TABLE definition)"
                    }
                    None => "missing SHRFMLA/ARRAY definition",
                };
                push_warning_bounded(
                    &mut warnings,
                    format!(
                        "failed to recover shared formula at {}: base cell {} stores PtgExp ({expected})",
                        cell_ref.to_a1(),
                        base_cell_ref.to_a1()
                    ),
                );
                continue;
            }
        }

        // Array formulas are anchored at the base cell and display the same formula text for every
        // cell in the group. If ARRAY definition records are missing but the base cell stores a
        // full `FORMULA.rgce` token stream, recover follower cells by decoding the base cell's rgce
        // in the base cell coordinate space (without materializing per-cell deltas).
        if grbit.membership_hint() == Some(FormulaMembershipHint::Array) {
            let base_coord = rgce::CellCoord::new(base_row, base_col);
            let decoded = rgce::decode_biff8_rgce_with_base_and_rgcb(
                base_rgce_bytes,
                base_rgcb_bytes,
                ctx,
                Some(base_coord),
            );
            if !decoded.warnings.is_empty() {
                for w in decoded.warnings {
                    push_warning_bounded(
                        &mut warnings,
                        format!(
                            "failed to fully decode recovered array formula at {}: {w}",
                            CellRef::new(row, col).to_a1()
                        ),
                    );
                }
            }

            if decoded.text.is_empty() {
                push_warning_bounded(
                    &mut warnings,
                    format!(
                        "failed to recover array formula at {}: decoded rgce produced empty text",
                        CellRef::new(row, col).to_a1()
                    ),
                );
                continue;
            }

            recovered.insert(CellRef::new(row, col), decoded.text);
            continue;
        }

        let Some(materialized) =
            materialize_biff8_rgce(base_rgce_bytes, base_row, base_col, row, col)
        else {
            push_warning_bounded(
                &mut warnings,
                format!(
                    "failed to recover shared formula at {}: could not materialize base rgce from {} (unsupported or malformed tokens)",
                    CellRef::new(row, col).to_a1(),
                    CellRef::new(base_row, base_col).to_a1()
                ),
            );
            continue;
        };

        let base_coord = rgce::CellCoord::new(row, col);
        let decoded = rgce::decode_biff8_rgce_with_base_and_rgcb(
            &materialized,
            base_rgcb_bytes,
            ctx,
            Some(base_coord),
        );
        if !decoded.warnings.is_empty() {
            for w in decoded.warnings {
                push_warning_bounded(
                    &mut warnings,
                    format!(
                        "failed to fully decode recovered shared formula at {}: {w}",
                        CellRef::new(row, col).to_a1()
                    ),
                );
            }
        }

        if decoded.text.is_empty() {
            // Avoid replacing an existing formula with an empty string.
            push_warning_bounded(
                &mut warnings,
                format!(
                    "failed to recover shared formula at {}: decoded rgce produced empty text",
                    CellRef::new(row, col).to_a1()
                ),
            );
            continue;
        }

        recovered.insert(CellRef::new(row, col), decoded.text);
    }

    Ok(PtgExpFallbackResult {
        formulas: recovered,
        warnings,
    })
}

/// Best-effort parse + materialization of BIFF8 worksheet formulas used as an override when
/// calamine mis-decodes certain token streams (notably `PtgArea3d` with relative flags).
///
/// The returned formulas are a map of cell -> decoded formula text (no leading `=`) that should
/// replace calamine’s output.
pub(crate) fn parse_biff8_sheet_formula_overrides(
    workbook_stream: &[u8],
    start: usize,
    ctx: &rgce::RgceDecodeContext<'_>,
) -> Result<SheetFormulaOverrides, String> {
    let parsed = worksheet_formulas::parse_biff8_worksheet_formulas(workbook_stream, start)?;

    let mut out = SheetFormulaOverrides::default();
    let mut selection_warnings: Vec<String> = Vec::new();

    for (cell_ref, cell) in parsed.formula_cells {
        // Shared formula reference (PtgExp).
        if let Some((_base_row, _base_col, base_candidates)) =
            parse_ptg_exp_master_cell_candidates(&cell.rgce)
        {
            selection_warnings.clear();
            let Some(selected) = select_ptgexp_backing_record(
                &parsed.shrfmla,
                cell_ref,
                &base_candidates,
                &mut selection_warnings,
                "SHRFMLA",
                |r| r.range,
            ) else {
                continue;
            };
            let shrfmla = selected.record;

            let Some(materialized) = materialize_biff8_rgce(
                &shrfmla.rgce,
                selected.range.0.row,
                selected.range.0.col,
                cell_ref.row,
                cell_ref.col,
            ) else {
                continue;
            };

            let decoded = decode_formula_text_best_effort(
                &materialized,
                &shrfmla.rgcb,
                cell_ref,
                ctx,
            );
            if decoded.warnings.is_empty()
                && !decoded.text.is_empty()
                && decoded.text != "#UNKNOWN!"
            {
                out.formulas.insert(cell_ref, decoded.text);
            } else {
                // Surface array-constant decode failures for better diagnostics.
                for w in decoded.warnings {
                    if w.contains("PtgArray") {
                        out.warnings.push(format!("cell {}: {w}", cell_ref.to_a1()));
                    }
                }
            }
            continue;
        }

        // Non-shared formulas: only override when we detect a 3D area token that uses relative
        // flags (these are easy to mis-decode if the high bits of the column fields are treated as
        // part of the column index).
        let has_area3d_relative = rgce_contains_area3d_relative_flags(&cell.rgce);
        let has_ptgarray = rgce_contains_ptgarray(&cell.rgce);
        if !has_area3d_relative && !has_ptgarray {
            continue;
        }

        let decoded = decode_formula_text_best_effort(&cell.rgce, &cell.rgcb, cell_ref, ctx);
        if decoded.warnings.is_empty() && !decoded.text.is_empty() && decoded.text != "#UNKNOWN!" {
            out.formulas.insert(cell_ref, decoded.text);
        } else if has_ptgarray {
            for w in decoded.warnings {
                if w.contains("PtgArray") {
                    out.warnings.push(format!("cell {}: {w}", cell_ref.to_a1()));
                }
            }
        }
    }

    Ok(out)
}

fn decode_formula_text_best_effort(
    rgce_bytes: &[u8],
    rgcb: &[u8],
    cell_ref: CellRef,
    ctx: &rgce::RgceDecodeContext<'_>,
) -> rgce::DecodeRgceResult {
    let base = rgce::CellCoord::new(cell_ref.row, cell_ref.col);
    rgce::decode_biff8_rgce_with_base_and_rgcb(rgce_bytes, rgcb, ctx, Some(base))
}

fn range_contains(range: (CellRef, CellRef), cell: CellRef) -> bool {
    let (start, end) = range;
    cell.row >= start.row && cell.row <= end.row && cell.col >= start.col && cell.col <= end.col
}

fn rgce_contains_ptgarray(rgce_bytes: &[u8]) -> bool {
    // Best-effort scan: stay aligned for common fixed-width ptgs and bail on unknown tokens.
    let mut i = 0usize;
    while i < rgce_bytes.len() {
        let ptg = rgce_bytes[i];
        i += 1;

        match ptg {
            // PtgArray (array constant): [unused: 7 bytes] + values stored in trailing rgcb.
            0x20 | 0x40 | 0x60 => return true,

            // PtgExp / PtgTbl: [rw:u16][col:u16]
            0x01 | 0x02 => i = i.saturating_add(4),

            // Fixed-width/no-payload operators + PtgParen.
            0x03..=0x16 | 0x2F => {}

            // PtgAttr: [grbit:u8][wAttr:u16] (+ optional jump table for tAttrChoose)
            0x19 => {
                if i + 3 > rgce_bytes.len() {
                    return false;
                }
                let grbit = rgce_bytes[i];
                let w_attr = u16::from_le_bytes([rgce_bytes[i + 1], rgce_bytes[i + 2]]) as usize;
                i += 3;

                const T_ATTR_CHOOSE: u8 = 0x04;
                if grbit & T_ATTR_CHOOSE != 0 {
                    let needed = w_attr.saturating_mul(2);
                    i = i.saturating_add(needed);
                }
            }

            // PtgStr: [cch:u8][flags:u8][chars...]
            0x17 => {
                if i + 2 > rgce_bytes.len() {
                    return false;
                }
                let cch = rgce_bytes[i] as usize;
                let flags = rgce_bytes[i + 1];
                i += 2;
                let bytes = if flags & 0x01 != 0 { cch.saturating_mul(2) } else { cch };
                i = i.saturating_add(bytes);
            }

            // PtgErr / PtgBool
            0x1C | 0x1D => i = i.saturating_add(1),
            // PtgInt
            0x1E => i = i.saturating_add(2),
            // PtgNum
            0x1F => i = i.saturating_add(8),

            // PtgFunc
            0x21 | 0x41 | 0x61 => i = i.saturating_add(2),
            // PtgFuncVar
            0x22 | 0x42 | 0x62 => i = i.saturating_add(3),
            // PtgName
            0x23 | 0x43 | 0x63 => i = i.saturating_add(6),

            // PtgRef
            0x24 | 0x44 | 0x64 => i = i.saturating_add(4),
            // PtgArea
            0x25 | 0x45 | 0x65 => i = i.saturating_add(8),

            // PtgMem* tokens: [cce:u16][rgce:cce bytes]
            0x26 | 0x46 | 0x66 | 0x27 | 0x47 | 0x67 | 0x28 | 0x48 | 0x68 | 0x29 | 0x49 | 0x69
            | 0x2E | 0x4E | 0x6E => {
                if i + 2 > rgce_bytes.len() {
                    return false;
                }
                let cce = u16::from_le_bytes([rgce_bytes[i], rgce_bytes[i + 1]]) as usize;
                i = i.saturating_add(2 + cce);
            }

            // 3D references: PtgRef3d / PtgArea3d.
            0x3A | 0x5A | 0x7A => i = i.saturating_add(6),
            0x3B | 0x5B | 0x7B => i = i.saturating_add(10),

            // Unknown/unsupported token; bail so we don't mis-scan.
            _ => return false,
        }

        if i > rgce_bytes.len() {
            return false;
        }
    }

    false
}

fn rgce_contains_area3d_relative_flags(rgce_bytes: &[u8]) -> bool {
    // Best-effort scan: parse only a subset of tokens needed to find PtgArea3d while staying
    // aligned for common fixed-width ptgs.
    let mut i = 0usize;
    while i < rgce_bytes.len() {
        let ptg = rgce_bytes[i];
        i += 1;

        match ptg {
            // PtgExp / PtgTbl: [rw:u16][col:u16]
            0x01 | 0x02 => i = i.saturating_add(4),

            // Fixed-width/no-payload operators + PtgParen.
            0x03..=0x16 | 0x2F => {}

            // PtgAttr: [grbit:u8][wAttr:u16] (+ optional jump table for tAttrChoose)
            0x19 => {
                if i + 3 > rgce_bytes.len() {
                    return false;
                }
                let grbit = rgce_bytes[i];
                let w_attr = u16::from_le_bytes([rgce_bytes[i + 1], rgce_bytes[i + 2]]) as usize;
                i += 3;

                const T_ATTR_CHOOSE: u8 = 0x04;
                if grbit & T_ATTR_CHOOSE != 0 {
                    let needed = w_attr.saturating_mul(2);
                    i = i.saturating_add(needed);
                }
            }

            // PtgStr: [cch:u8][flags:u8][chars...]
            0x17 => {
                if i + 2 > rgce_bytes.len() {
                    return false;
                }
                let cch = rgce_bytes[i] as usize;
                let flags = rgce_bytes[i + 1];
                i += 2;
                let bytes = if flags & 0x01 != 0 { cch.saturating_mul(2) } else { cch };
                i = i.saturating_add(bytes);
                // Ignore rich/ext segments; this is best-effort and our fixtures emit none.
            }

            // PtgErr / PtgBool
            0x1C | 0x1D => i = i.saturating_add(1),
            // PtgInt
            0x1E => i = i.saturating_add(2),
            // PtgNum
            0x1F => i = i.saturating_add(8),
            // PtgArray
            0x20 | 0x40 | 0x60 => i = i.saturating_add(7),
            // PtgFunc
            0x21 | 0x41 | 0x61 => i = i.saturating_add(2),
            // PtgFuncVar
            0x22 | 0x42 | 0x62 => i = i.saturating_add(3),
            // PtgName
            0x23 | 0x43 | 0x63 => i = i.saturating_add(6),

            // PtgRef: [row:u16][col+flags:u16]
            0x24 | 0x44 | 0x64 => i = i.saturating_add(4),
            // PtgArea: [row1:u16][row2:u16][col1+flags:u16][col2+flags:u16]
            0x25 | 0x45 | 0x65 => i = i.saturating_add(8),

            // PtgMem* tokens: [cce:u16][rgce:cce bytes]
            0x26 | 0x46 | 0x66 | 0x27 | 0x47 | 0x67 | 0x28 | 0x48 | 0x68 | 0x29 | 0x49 | 0x69
            | 0x2E | 0x4E | 0x6E => {
                if i + 2 > rgce_bytes.len() {
                    return false;
                }
                let cce = u16::from_le_bytes([rgce_bytes[i], rgce_bytes[i + 1]]) as usize;
                i = i.saturating_add(2 + cce);
            }

            // PtgRef3d: [ixti:u16][row:u16][col+flags:u16]
            0x3A | 0x5A | 0x7A => i = i.saturating_add(6),

            // PtgArea3d: [ixti:u16][row1:u16][row2:u16][col1+flags:u16][col2+flags:u16]
            0x3B | 0x5B | 0x7B => {
                if i + 10 > rgce_bytes.len() {
                    return false;
                }
                let col1_off = i + 6;
                let col2_off = i + 8;
                let col1 = u16::from_le_bytes([rgce_bytes[col1_off], rgce_bytes[col1_off + 1]]);
                let col2 = u16::from_le_bytes([rgce_bytes[col2_off], rgce_bytes[col2_off + 1]]);
                if (col1 & RELATIVE_MASK) != 0 || (col2 & RELATIVE_MASK) != 0 {
                    return true;
                }
                i = i.saturating_add(10);
            }

            // Unknown/unsupported token; bail so we don't mis-scan.
            _ => return false,
        }

        if i > rgce_bytes.len() {
            return false;
        }
    }

    false
}

/// Materialize a BIFF8 `rgce` token stream from a base cell into a target cell by applying the
/// row/col delta to tokens that embed absolute coordinates plus relative flags.
///
/// This is used by:
/// - the `PtgExp` fallback (when `SHRFMLA`/`ARRAY` definition records are missing)
/// - shared-formula decoding (when we need to expand `SHRFMLA` token streams into per-cell `rgce`)
pub(crate) fn materialize_biff8_rgce_from_base(
    base_rgce: &[u8],
    base_cell: CellRef,
    target_cell: CellRef,
) -> Option<Vec<u8>> {
    materialize_biff8_rgce(
        base_rgce,
        base_cell.row,
        base_cell.col,
        target_cell.row,
        target_cell.col,
    )
}

fn parse_ptg_exp(rgce: &[u8]) -> Option<(u32, u32)> {
    // BIFF8 PtgExp: [0x01][rw: u16][col: u16]
    if rgce.first().copied()? != 0x01 {
        return None;
    }
    if rgce.len() < 5 {
        return None;
    }
    let row = u16::from_le_bytes([rgce[1], rgce[2]]) as u32;
    let col = u16::from_le_bytes([rgce[3], rgce[4]]) as u32;
    Some((row, col))
}

fn parse_ptg_exp_master_cell_candidates(rgce: &[u8]) -> Option<(u32, u32, Vec<CellRef>)> {
    let (row, col) = parse_ptg_exp(rgce)?;

    let mut candidates: Vec<CellRef> = Vec::new();
    if row < EXCEL_MAX_ROWS && col < EXCEL_MAX_COLS {
        candidates.push(CellRef::new(row, col));
    }

    // Best-effort: some `.xls` producers appear to swap row/col ordering in the PtgExp payload.
    // Keep a second candidate when it is in-bounds (for the model) and distinct.
    if col < EXCEL_MAX_ROWS && row < EXCEL_MAX_COLS {
        let swapped = CellRef::new(col, row);
        if candidates.first().copied() != Some(swapped) {
            candidates.push(swapped);
        }
    }

    Some((row, col, candidates))
}

fn cell_in_bounds(row: i64, col: i64) -> bool {
    row >= 0 && row <= BIFF8_MAX_ROW0 && col >= 0 && col <= BIFF8_MAX_COL0
}

fn sign_extend_14(v: u16) -> i16 {
    debug_assert!(v <= COL_INDEX_MASK);
    // 14-bit two's complement. If bit13 is set, treat as negative.
    if (v & 0x2000) != 0 {
        (v | 0xC000) as i16
    } else {
        v as i16
    }
}

fn pack_col_with_flags(col0: u16, flags: u16) -> u16 {
    (col0 & COL_INDEX_MASK) | (flags & RELATIVE_MASK)
}

fn adjust_row_col(row0: u16, col_field: u16, delta_row: i64, delta_col: i64) -> Option<(u16, u16)> {
    let row_rel = (col_field & ROW_RELATIVE_BIT) != 0;
    let col_rel = (col_field & COL_RELATIVE_BIT) != 0;
    let col0 = (col_field & COL_INDEX_MASK) as i64;
    let row0_i64 = row0 as i64;

    let new_row = if row_rel {
        row0_i64 + delta_row
    } else {
        row0_i64
    };
    let new_col = if col_rel { col0 + delta_col } else { col0 };

    if !cell_in_bounds(new_row, new_col) {
        return None;
    }

    let new_row_u16 = new_row as u16;
    let new_col_u16 = pack_col_with_flags(new_col as u16, col_field);
    Some((new_row_u16, new_col_u16))
}

/// Materialize a BIFF8 `rgce` token stream from a base cell into the token stream that would
/// appear in a target cell after applying shared-formula row/col deltas.
///
/// This adjusts `PtgRef`/`PtgArea` (and 3D variants) that use relative flags while preserving
/// relative-offset tokens (`PtgRefN`/`PtgAreaN`) verbatim (those are resolved relative to the
/// current formula cell at decode time).
///
/// Returns `None` when the token stream contains an unsupported or malformed token whose payload
/// length cannot be determined.
pub(crate) fn materialize_biff8_rgce(
    base: &[u8],
    base_row: u32,
    base_col: u32,
    row: u32,
    col: u32,
) -> Option<Vec<u8>> {
    let delta_row = row as i64 - base_row as i64;
    let delta_col = col as i64 - base_col as i64;
    let cell_row = row as i64;
    let cell_col = col as i64;

    let mut out = Vec::with_capacity(base.len());
    let mut i = 0usize;
    while i < base.len() {
        let ptg = *base.get(i)?;
        i += 1;

        match ptg {
            // PtgExp / PtgTbl: not expected in a base-cell formula for this fallback.
            0x01 | 0x02 => return None,

            // Fixed-width / no-payload tokens.
            0x03..=0x16 | 0x2F => out.push(ptg),

            // PtgStr (ShortXLUnicodeString): variable.
            0x17 => {
                let len = biff8_short_unicode_string_len(base.get(i..)?)?;
                out.push(ptg);
                out.extend_from_slice(base.get(i..i + len)?);
                i += len;
            }

            // PtgExtend / PtgExtendV / PtgExtendA (ptg=0x18 variants).
            //
            // Structured references (tables) use an `etpg` subtype byte (`0x19` = PtgList) followed
            // by a fixed 12-byte payload. Other (unsupported) subtypes appear in the wild with an
            // opaque 5-byte payload; copy them verbatim so the rgce stream stays aligned.
            0x18 | 0x38 | 0x58 | 0x78 => {
                let etpg = *base.get(i)?;
                out.push(ptg);
                out.push(etpg);
                i += 1;

                if etpg == 0x19 {
                    out.extend_from_slice(base.get(i..i + 12)?);
                    i += 12;
                } else {
                    out.extend_from_slice(base.get(i..i + 4)?);
                    i += 4;
                }
            }

            // PtgAttr: [grbit: u8][wAttr: u16] (+ optional jump table for tAttrChoose)
            0x19 => {
                if i + 3 > base.len() {
                    return None;
                }
                out.push(ptg);
                let grbit = base[i];
                let w_attr = u16::from_le_bytes([base[i + 1], base[i + 2]]);
                out.extend_from_slice(&base[i..i + 3]);
                i += 3;

                const T_ATTR_CHOOSE: u8 = 0x04;
                if grbit & T_ATTR_CHOOSE != 0 {
                    let needed = (w_attr as usize).checked_mul(2)?;
                    out.extend_from_slice(base.get(i..i + needed)?);
                    i += needed;
                }
            }

            // PtgErr / PtgBool: 1 byte.
            0x1C | 0x1D => {
                out.push(ptg);
                out.push(*base.get(i)?);
                i += 1;
            }

            // PtgInt: 2 bytes.
            0x1E => {
                out.push(ptg);
                out.extend_from_slice(base.get(i..i + 2)?);
                i += 2;
            }

            // PtgNum: 8 bytes.
            0x1F => {
                out.push(ptg);
                out.extend_from_slice(base.get(i..i + 8)?);
                i += 8;
            }

            // PtgArray: 7 bytes of unused/reserved payload (array data stored elsewhere).
            0x20 | 0x40 | 0x60 => {
                out.push(ptg);
                out.extend_from_slice(base.get(i..i + 7)?);
                i += 7;
            }

            // PtgFunc: [iftab: u16]
            0x21 | 0x41 | 0x61 => {
                out.push(ptg);
                out.extend_from_slice(base.get(i..i + 2)?);
                i += 2;
            }

            // PtgFuncVar: [argc: u8][iftab: u16]
            0x22 | 0x42 | 0x62 => {
                out.push(ptg);
                out.extend_from_slice(base.get(i..i + 3)?);
                i += 3;
            }

            // PtgName: [name_id: u32][reserved: u16]
            0x23 | 0x43 | 0x63 => {
                out.push(ptg);
                out.extend_from_slice(base.get(i..i + 6)?);
                i += 6;
            }

            // PtgRef: [rw: u16][col: u16]
            0x24 | 0x44 | 0x64 => {
                let row0 = u16::from_le_bytes(base.get(i..i + 2)?.try_into().ok()?);
                let col_field = u16::from_le_bytes(base.get(i + 2..i + 4)?.try_into().ok()?);
                let payload = base.get(i..i + 4)?;

                if let Some((new_row, new_col_field)) =
                    adjust_row_col(row0, col_field, delta_row, delta_col)
                {
                    out.push(ptg);
                    out.extend_from_slice(&new_row.to_le_bytes());
                    out.extend_from_slice(&new_col_field.to_le_bytes());
                } else {
                    // Out-of-bounds references materialize as PtgRefErr*.
                    out.push(ptg.saturating_add(0x06));
                    out.extend_from_slice(payload);
                }

                i += 4;
            }

            // PtgArea: [rwFirst: u16][rwLast: u16][colFirst: u16][colLast: u16]
            0x25 | 0x45 | 0x65 => {
                let row1 = u16::from_le_bytes(base.get(i..i + 2)?.try_into().ok()?);
                let row2 = u16::from_le_bytes(base.get(i + 2..i + 4)?.try_into().ok()?);
                let col1 = u16::from_le_bytes(base.get(i + 4..i + 6)?.try_into().ok()?);
                let col2 = u16::from_le_bytes(base.get(i + 6..i + 8)?.try_into().ok()?);
                let payload = base.get(i..i + 8)?;

                let adjusted1 = adjust_row_col(row1, col1, delta_row, delta_col);
                let adjusted2 = adjust_row_col(row2, col2, delta_row, delta_col);
                if let (Some((new_row1, new_col1)), Some((new_row2, new_col2))) =
                    (adjusted1, adjusted2)
                {
                    out.push(ptg);
                    out.extend_from_slice(&new_row1.to_le_bytes());
                    out.extend_from_slice(&new_row2.to_le_bytes());
                    out.extend_from_slice(&new_col1.to_le_bytes());
                    out.extend_from_slice(&new_col2.to_le_bytes());
                } else {
                    out.push(ptg.saturating_add(0x06));
                    out.extend_from_slice(payload);
                }

                i += 8;
            }

            // PtgMem* tokens: [cce: u16][rgce: cce bytes]
            0x26 | 0x46 | 0x66 | 0x27 | 0x47 | 0x67 | 0x28 | 0x48 | 0x68 | 0x29 | 0x49 | 0x69
            | 0x2E | 0x4E | 0x6E => {
                if i + 2 > base.len() {
                    return None;
                }
                let cce = u16::from_le_bytes([base[i], base[i + 1]]) as usize;
                out.push(ptg);
                out.extend_from_slice(&base[i..i + 2]);
                i += 2;
                out.extend_from_slice(base.get(i..i + cce)?);
                i += cce;
            }

            // PtgRefErr / PtgRefErrN: [rw: u16][col: u16]
            0x2A | 0x4A | 0x6A => {
                out.push(ptg);
                out.extend_from_slice(base.get(i..i + 4)?);
                i += 4;
            }

            // PtgAreaErr / PtgAreaErrN: [rwFirst: u16][rwLast: u16][colFirst: u16][colLast: u16]
            0x2B | 0x4B | 0x6B => {
                out.push(ptg);
                out.extend_from_slice(base.get(i..i + 8)?);
                i += 8;
            }

            // PtgRefN: offsets are resolved relative to the *current* formula cell at decode time.
            // Keep the token verbatim unless the resulting reference would be out-of-bounds, in
            // which case Excel materializes it as `PtgRefErr*`.
            0x2C | 0x4C | 0x6C => {
                let payload = base.get(i..i + 4)?;
                let row_raw = u16::from_le_bytes(payload.get(0..2)?.try_into().ok()?);
                let col_field = u16::from_le_bytes(payload.get(2..4)?.try_into().ok()?);

                let row_rel = (col_field & ROW_RELATIVE_BIT) != 0;
                let col_rel = (col_field & COL_RELATIVE_BIT) != 0;
                let col_raw = col_field & COL_INDEX_MASK;

                let abs_row = if row_rel {
                    cell_row.saturating_add(row_raw as i16 as i64)
                } else {
                    row_raw as i64
                };
                let abs_col = if col_rel {
                    cell_col.saturating_add(sign_extend_14(col_raw) as i64)
                } else {
                    col_raw as i64
                };

                if cell_in_bounds(abs_row, abs_col) {
                    out.push(ptg);
                } else {
                    // Excel materializes out-of-bounds `PtgRefN` / `PtgAreaN` refs in shared formulas
                    // as `#REF!` error ptgs (`PtgRefErr*` / `PtgAreaErr*`).
                    out.push(ptg.saturating_sub(0x02));
                }
                out.extend_from_slice(payload);
                i += 4;
            }

            // PtgAreaN: like `PtgRefN`, keep verbatim unless out-of-bounds (then materialize as
            // `PtgAreaErr*`).
            0x2D | 0x4D | 0x6D => {
                let payload = base.get(i..i + 8)?;
                let row1_raw = u16::from_le_bytes(payload.get(0..2)?.try_into().ok()?);
                let row2_raw = u16::from_le_bytes(payload.get(2..4)?.try_into().ok()?);
                let col1_field = u16::from_le_bytes(payload.get(4..6)?.try_into().ok()?);
                let col2_field = u16::from_le_bytes(payload.get(6..8)?.try_into().ok()?);

                let row1_rel = (col1_field & ROW_RELATIVE_BIT) != 0;
                let col1_rel = (col1_field & COL_RELATIVE_BIT) != 0;
                let row2_rel = (col2_field & ROW_RELATIVE_BIT) != 0;
                let col2_rel = (col2_field & COL_RELATIVE_BIT) != 0;

                let col1_raw = col1_field & COL_INDEX_MASK;
                let col2_raw = col2_field & COL_INDEX_MASK;

                let abs_row1 = if row1_rel {
                    cell_row.saturating_add(row1_raw as i16 as i64)
                } else {
                    row1_raw as i64
                };
                let abs_row2 = if row2_rel {
                    cell_row.saturating_add(row2_raw as i16 as i64)
                } else {
                    row2_raw as i64
                };

                let abs_col1 = if col1_rel {
                    cell_col.saturating_add(sign_extend_14(col1_raw) as i64)
                } else {
                    col1_raw as i64
                };
                let abs_col2 = if col2_rel {
                    cell_col.saturating_add(sign_extend_14(col2_raw) as i64)
                } else {
                    col2_raw as i64
                };

                if cell_in_bounds(abs_row1, abs_col1) && cell_in_bounds(abs_row2, abs_col2) {
                    out.push(ptg);
                } else {
                    out.push(ptg.saturating_sub(0x02));
                }
                out.extend_from_slice(payload);
                i += 8;
            }

            // PtgNameX: [ixti: u16][iname: u16][reserved: u16]
            0x39 | 0x59 | 0x79 => {
                out.push(ptg);
                out.extend_from_slice(base.get(i..i + 6)?);
                i += 6;
            }

            // PtgRef3d: [ixti: u16][rw: u16][col: u16]
            0x3A | 0x5A | 0x7A => {
                let ixti = u16::from_le_bytes(base.get(i..i + 2)?.try_into().ok()?);
                let row0 = u16::from_le_bytes(base.get(i + 2..i + 4)?.try_into().ok()?);
                let col_field = u16::from_le_bytes(base.get(i + 4..i + 6)?.try_into().ok()?);
                let payload = base.get(i..i + 6)?;

                if let Some((new_row, new_col_field)) =
                    adjust_row_col(row0, col_field, delta_row, delta_col)
                {
                    out.push(ptg);
                    out.extend_from_slice(&ixti.to_le_bytes());
                    out.extend_from_slice(&new_row.to_le_bytes());
                    out.extend_from_slice(&new_col_field.to_le_bytes());
                } else {
                    // Out-of-bounds -> PtgRefErr3d*
                    out.push(ptg.saturating_add(0x02));
                    out.extend_from_slice(payload);
                }

                i += 6;
            }

            // PtgArea3d: [ixti: u16][rwFirst: u16][rwLast: u16][colFirst: u16][colLast: u16]
            0x3B | 0x5B | 0x7B => {
                let ixti = u16::from_le_bytes(base.get(i..i + 2)?.try_into().ok()?);
                let row1 = u16::from_le_bytes(base.get(i + 2..i + 4)?.try_into().ok()?);
                let row2 = u16::from_le_bytes(base.get(i + 4..i + 6)?.try_into().ok()?);
                let col1 = u16::from_le_bytes(base.get(i + 6..i + 8)?.try_into().ok()?);
                let col2 = u16::from_le_bytes(base.get(i + 8..i + 10)?.try_into().ok()?);
                let payload = base.get(i..i + 10)?;

                let adjusted1 = adjust_row_col(row1, col1, delta_row, delta_col);
                let adjusted2 = adjust_row_col(row2, col2, delta_row, delta_col);
                if let (Some((new_row1, new_col1)), Some((new_row2, new_col2))) =
                    (adjusted1, adjusted2)
                {
                    out.push(ptg);
                    out.extend_from_slice(&ixti.to_le_bytes());
                    out.extend_from_slice(&new_row1.to_le_bytes());
                    out.extend_from_slice(&new_row2.to_le_bytes());
                    out.extend_from_slice(&new_col1.to_le_bytes());
                    out.extend_from_slice(&new_col2.to_le_bytes());
                } else {
                    // Out-of-bounds -> PtgAreaErr3d*
                    out.push(ptg.saturating_add(0x02));
                    out.extend_from_slice(payload);
                }

                i += 10;
            }

            // PtgRefErr3d / PtgAreaErr3d: copy verbatim.
            0x3C | 0x5C | 0x7C => {
                out.push(ptg);
                out.extend_from_slice(base.get(i..i + 6)?);
                i += 6;
            }
            0x3D | 0x5D | 0x7D => {
                out.push(ptg);
                out.extend_from_slice(base.get(i..i + 10)?);
                i += 10;
            }

            // PtgRefN3d / PtgAreaN3d: keep verbatim.
            0x3E | 0x5E | 0x7E => {
                out.push(ptg);
                out.extend_from_slice(base.get(i..i + 6)?);
                i += 6;
            }
            0x3F | 0x5F | 0x7F => {
                out.push(ptg);
                out.extend_from_slice(base.get(i..i + 10)?);
                i += 10;
            }

            _ => return None,
        }
    }

    Some(out)
}

fn biff8_short_unicode_string_len(input: &[u8]) -> Option<usize> {
    // ShortXLUnicodeString payload:
    //   [cch: u8][flags: u8]
    //   [cRun: u16]? (if flags & 0x08)
    //   [cbExtRst: u32]? (if flags & 0x04)
    //   [chars: cch bytes or cch*2 bytes]
    //   [rgRun: 4*cRun bytes]?
    //   [ext: cbExtRst bytes]?
    if input.len() < 2 {
        return None;
    }
    let cch = input[0] as usize;
    let flags = input[1];
    let mut offset = 2usize;

    let rich_runs = if flags & 0x08 != 0 {
        let runs = u16::from_le_bytes([*input.get(offset)?, *input.get(offset + 1)?]) as usize;
        offset += 2;
        runs
    } else {
        0usize
    };

    let ext_size = if flags & 0x04 != 0 {
        let size = u32::from_le_bytes([
            *input.get(offset)?,
            *input.get(offset + 1)?,
            *input.get(offset + 2)?,
            *input.get(offset + 3)?,
        ]) as usize;
        offset += 4;
        size
    } else {
        0usize
    };

    let is_unicode = flags & 0x01 != 0;
    let char_bytes = if is_unicode { cch.checked_mul(2)? } else { cch };
    offset = offset.checked_add(char_bytes)?;

    let rich_bytes = rich_runs.checked_mul(4)?;
    offset = offset.checked_add(rich_bytes)?;
    offset = offset.checked_add(ext_size)?;

    (input.len() >= offset).then_some(offset)
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
    fn materializes_ref_col_oob_to_referr_variants() {
        // When a `PtgRef*` token shifts out of bounds during shared-formula materialization, the
        // materializer must emit the 2D error ptg (`PtgRefErr*`), preserving token width.
        for &ptg_ref in &[0x24_u8, 0x44, 0x64] {
            // PtgRef payload: [row:u16][col+flags:u16]
            // Use col=0x3FFF (14-bit max) with the col-relative flag so shifting by +1 col is OOB.
            let base: Vec<u8> = vec![
                ptg_ref, // PtgRef (ref/value/array class)
                0x00,
                0x00, // row=0
                0xFF,
                0xBF, // col=0x3FFF with COL_RELATIVE_BIT set (0xBFFF)
            ];

            let out = materialize_biff8_rgce(&base, 0, 0, 0, 1).expect("materialize");
            assert_eq!(out[0], ptg_ref + 0x06, "ptg={ptg_ref:02X}");
            assert_eq!(&out[1..], &base[1..], "payload should be preserved");
        }
    }

    #[test]
    fn materializes_area_col_oob_to_areaerr_variants() {
        // When a `PtgArea*` token shifts out of bounds during shared-formula materialization, the
        // materializer must emit the 2D error ptg (`PtgAreaErr*`), preserving token width.
        for &ptg_area in &[0x25_u8, 0x45, 0x65] {
            // PtgArea payload: [row1:u16][row2:u16][col1+flags:u16][col2+flags:u16]
            // Use col2=0x3FFF (14-bit max) with the col-relative flag so shifting by +1 col is OOB.
            // Keep col1 in-bounds so we cover the "one endpoint OOB" case.
            let mut base = Vec::new();
            base.push(ptg_area);
            base.extend_from_slice(&0u16.to_le_bytes()); // row1=0
            base.extend_from_slice(&0u16.to_le_bytes()); // row2=0
            base.extend_from_slice(&0xBFFEu16.to_le_bytes()); // col1=0x3FFE + COL_RELATIVE_BIT
            base.extend_from_slice(&0xBFFFu16.to_le_bytes()); // col2=0x3FFF + COL_RELATIVE_BIT

            let out = materialize_biff8_rgce(&base, 0, 0, 0, 1).expect("materialize");
            assert_eq!(out[0], ptg_area + 0x06, "ptg={ptg_area:02X}");
            assert_eq!(&out[1..], &base[1..], "payload should be preserved");
        }
    }

    #[test]
    fn materializes_ref3d_oob_to_referr3d_variants() {
        // When a `PtgRef3d*` token shifts out of BIFF8 bounds during shared-formula materialization,
        // the materializer must emit the *3D* error ptg (`PtgRefErr3d*`), preserving token width.
        for &ptg_ref3d in &[0x3A_u8, 0x5A, 0x7A] {
            // PtgRef3d payload: [ixti:u16][row:u16][col+flags:u16]
            // Use row=65535 (max) with the row-relative flag so shifting by +1 row is OOB.
            let base: Vec<u8> = vec![
                ptg_ref3d, // PtgRef3d (ref/value/array class)
                0x00, 0x00, // ixti=0
                0xFF, 0xFF, // row=65535
                0x00, 0x40, // col=0 with ROW_RELATIVE_BIT set
            ];

            let out = materialize_biff8_rgce(&base, 0, 0, 1, 0).expect("materialize");
            assert_eq!(out[0], ptg_ref3d + 0x02, "ptg={ptg_ref3d:02X}");
            assert_eq!(&out[1..], &base[1..], "payload should be preserved");
        }
    }

    #[test]
    fn materializes_ref3d_col_oob_to_referr3d_variants() {
        // When a `PtgRef3d*` token shifts out of BIFF8 bounds during shared-formula materialization,
        // the materializer must emit the *3D* error ptg (`PtgRefErr3d*`), preserving token width.
        for &ptg_ref3d in &[0x3A_u8, 0x5A, 0x7A] {
            // PtgRef3d payload: [ixti:u16][row:u16][col+flags:u16]
            // Use col=0x3FFF (14-bit max) with the col-relative flag so shifting by +1 col is OOB.
            let base: Vec<u8> = vec![
                ptg_ref3d, // PtgRef3d (ref/value/array class)
                0x00, 0x00, // ixti=0
                0x00, 0x00, // row=0
                0xFF, 0xBF, // col=0x3FFF with COL_RELATIVE_BIT set (0xBFFF)
            ];

            let out = materialize_biff8_rgce(&base, 0, 0, 0, 1).expect("materialize");
            assert_eq!(out[0], ptg_ref3d + 0x02, "ptg={ptg_ref3d:02X}");
            assert_eq!(&out[1..], &base[1..], "payload should be preserved");
        }
    }

    #[test]
    fn materializes_area3d_oob_to_areaerr3d_variants() {
        // When a `PtgArea3d*` token shifts out of BIFF8 bounds during shared-formula materialization,
        // the materializer must emit the *3D* error ptg (`PtgAreaErr3d*`).
        for &ptg_area3d in &[0x3B_u8, 0x5B, 0x7B] {
            // PtgArea3d payload:
            //   [ixti:u16][row1:u16][row2:u16][col1+flags:u16][col2+flags:u16]
            // Use row2=65535 (max) with row-relative flags so shifting by +1 row is OOB.
            let mut base = Vec::new();
            base.push(ptg_area3d);
            base.extend_from_slice(&0u16.to_le_bytes()); // ixti=0
            base.extend_from_slice(&0u16.to_le_bytes()); // row1=0
            base.extend_from_slice(&u16::MAX.to_le_bytes()); // row2=65535 (max)
            base.extend_from_slice(&0x4000u16.to_le_bytes()); // col1=0 + ROW_RELATIVE_BIT
            base.extend_from_slice(&0x4000u16.to_le_bytes()); // col2=0 + ROW_RELATIVE_BIT

            let out = materialize_biff8_rgce(&base, 0, 0, 1, 0).expect("materialize");
            assert_eq!(out[0], ptg_area3d + 0x02, "ptg={ptg_area3d:02X}");
            assert_eq!(&out[1..], &base[1..], "payload should be preserved");
        }
    }

    #[test]
    fn materializes_area3d_col_oob_to_areaerr3d_variants() {
        // When a `PtgArea3d*` token shifts out of BIFF8 bounds during shared-formula materialization,
        // the materializer must emit the *3D* error ptg (`PtgAreaErr3d*`).
        for &ptg_area3d in &[0x3B_u8, 0x5B, 0x7B] {
            // PtgArea3d payload:
            //   [ixti:u16][row1:u16][row2:u16][col1+flags:u16][col2+flags:u16]
            // Use col2=0x3FFF (14-bit max) with the col-relative flag so shifting by +1 col is OOB.
            // Keep col1 in-bounds so we cover the "one endpoint OOB" case.
            let mut base = Vec::new();
            base.push(ptg_area3d);
            base.extend_from_slice(&0u16.to_le_bytes()); // ixti=0
            base.extend_from_slice(&0u16.to_le_bytes()); // row1=0
            base.extend_from_slice(&0u16.to_le_bytes()); // row2=0
            base.extend_from_slice(&0xBFFEu16.to_le_bytes()); // col1=0x3FFE + COL_RELATIVE_BIT
            base.extend_from_slice(&0xBFFFu16.to_le_bytes()); // col2=0x3FFF + COL_RELATIVE_BIT

            let out = materialize_biff8_rgce(&base, 0, 0, 0, 1).expect("materialize");
            assert_eq!(out[0], ptg_area3d + 0x02, "ptg={ptg_area3d:02X}");
            assert_eq!(&out[1..], &base[1..], "payload should be preserved");
        }
    }

    #[test]
    fn shrfmla_ptgexp_selection_uses_deterministic_tiebreak_and_warns() {
        let a1 = CellRef::new(0, 0);
        let b1 = CellRef::new(0, 1);
        let b2 = CellRef::new(1, 1);
        let c2 = CellRef::new(1, 2);

        let mut shrfmla: HashMap<CellRef, worksheet_formulas::Biff8ShrFmlaRecord> = HashMap::new();
        shrfmla.insert(
            a1,
            worksheet_formulas::Biff8ShrFmlaRecord {
                range: (a1, b2),
                rgce: Vec::new(),
                rgcb: Vec::new(),
            },
        );
        shrfmla.insert(
            b1,
            worksheet_formulas::Biff8ShrFmlaRecord {
                range: (b1, c2),
                rgce: Vec::new(),
                rgcb: Vec::new(),
            },
        );

        // Use a master cell that is inside both ranges but does not equal either range start, so
        // the resolver must fall back to containment matching (and warn on ambiguity).
        let masters = vec![b2];

        let mut warnings = Vec::new();
        let selected = select_ptgexp_backing_record(
            &shrfmla,
            b1,
            &masters,
            &mut warnings,
            "SHRFMLA",
            |r| r.range,
        )
        .expect("expected a match");

        assert_eq!(selected.range, (b1, c2));
        assert!(
            warnings.iter().any(|w| w.contains("ambiguous")),
            "expected ambiguity warning, got {warnings:?}"
        );
    }

    #[test]
    fn recovers_shrfmla_ptgarray_constants_using_rgcb() {
        // Ensure `recover_ptgexp_formulas_from_shrfmla_and_array` decodes `PtgArray` tokens stored
        // in SHRFMLA using the trailing `rgcb` data blocks (otherwise the array constant would be
        // rendered as `#UNKNOWN!`).

        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<crate::biff::externsheet::ExternSheetEntry> = Vec::new();
        let supbooks: Vec<crate::biff::supbook::SupBookInfo> = Vec::new();
        let defined_names: Vec<rgce::DefinedNameMeta> = Vec::new();
        let ctx = rgce::RgceDecodeContext {
            codepage: 1252,
            sheet_names: &sheet_names,
            externsheet: &externsheet,
            supbooks: &supbooks,
            defined_names: &defined_names,
        };

        // Follower cell B1 contains PtgExp -> A1.
        let mut formula = Vec::<u8>::new();
        formula.extend_from_slice(&0u16.to_le_bytes()); // row
        formula.extend_from_slice(&1u16.to_le_bytes()); // col
        formula.extend_from_slice(&0u16.to_le_bytes()); // xf
        formula.extend_from_slice(&[0u8; 8]); // cached result
        formula.extend_from_slice(&0x0008u16.to_le_bytes()); // grbit (fShrFmla)
        formula.extend_from_slice(&[0u8; 4]); // calc chain
        formula.extend_from_slice(&5u16.to_le_bytes()); // cce
        formula.push(0x01); // PtgExp
        formula.extend_from_slice(&0u16.to_le_bytes()); // base row (A1)
        formula.extend_from_slice(&0u16.to_le_bytes()); // base col (A1)

        // SHRFMLA record: range A1:B1 with rgce = PtgArray and trailing rgcb = {1,2}.
        let mut rgcb = Vec::<u8>::new();
        rgcb.extend_from_slice(&1u16.to_le_bytes()); // cols_minus1 = 1 => 2 cols
        rgcb.extend_from_slice(&0u16.to_le_bytes()); // rows_minus1 = 0 => 1 row
        for n in [1.0f64, 2.0] {
            rgcb.push(0x01); // number
            rgcb.extend_from_slice(&n.to_le_bytes());
        }

        let mut shrfmla = Vec::<u8>::new();
        // RefU range header: [rwFirst:u16][rwLast:u16][colFirst:u8][colLast:u8]
        shrfmla.extend_from_slice(&0u16.to_le_bytes()); // rwFirst
        shrfmla.extend_from_slice(&0u16.to_le_bytes()); // rwLast
        shrfmla.push(0u8); // colFirst
        shrfmla.push(1u8); // colLast
        shrfmla.extend_from_slice(&0u16.to_le_bytes()); // cUse
        shrfmla.extend_from_slice(&8u16.to_le_bytes()); // cce (PtgArray + 7 bytes)
        shrfmla.push(0x20); // PtgArray
        shrfmla.extend_from_slice(&[0u8; 7]); // opaque PtgArray header
        shrfmla.extend_from_slice(&rgcb); // rgcb trailing bytes

        let mut stream = Vec::<u8>::new();
        stream.extend_from_slice(&record(records::RECORD_BOF_BIFF8, &[]));
        stream.extend_from_slice(&record(worksheet_formulas::RECORD_FORMULA, &formula));
        stream.extend_from_slice(&record(worksheet_formulas::RECORD_SHRFMLA, &shrfmla));
        stream.extend_from_slice(&record(records::RECORD_EOF, &[]));

        let recovered =
            recover_ptgexp_formulas_from_shrfmla_and_array(&stream, 0, &ctx).expect("recover");
        assert_eq!(
            recovered
                .formulas
                .get(&CellRef::new(0, 1))
                .map(|s| s.as_str()),
            Some("{1,2}")
        );
    }
}
