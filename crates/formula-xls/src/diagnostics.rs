//! Diagnostics helpers for inspecting BIFF `.xls` workbooks.
//!
//! This module is intentionally **not** a full BIFF parser. It is designed for corpus triage:
//! quantify how often workbooks use shared/array/table formula constructs and how often those
//! constructs cannot be resolved with our best-effort logic.

use std::collections::HashSet;
use std::path::Path;

use crate::biff;

/// Per-worksheet counts of BIFF worksheet formula constructs.
#[derive(Debug, Default, Clone, PartialEq, Eq)]
pub struct SheetFormulaStats {
    /// Number of `FORMULA` records (worksheet cell formula records).
    pub formula_records: usize,
    /// Number of `FORMULA` records whose `rgce` begins with `PtgExp`.
    pub formula_ptgexp: usize,
    /// Number of `FORMULA` records whose `rgce` begins with `PtgTbl`.
    pub formula_ptgtbl: usize,
    /// Number of `SHRFMLA` records.
    pub shrfmla_records: usize,
    /// Number of `ARRAY` records.
    pub array_records: usize,
    /// Number of `TABLE` records.
    pub table_records: usize,
    /// Number of `PtgExp`-backed formulas that could not be resolved.
    pub unresolved_ptgexp: usize,
    /// Number of `PtgTbl`-backed formulas that could not be resolved.
    pub unresolved_ptgtbl: usize,
    /// Malformed BIFF record boundary errors (truncated headers/lengths).
    pub record_parse_errors: usize,
    /// Malformed/unsupported payloads encountered while extracting stats.
    pub payload_parse_errors: usize,
}

impl SheetFormulaStats {
    fn add_assign(&mut self, other: &SheetFormulaStats) {
        self.formula_records += other.formula_records;
        self.formula_ptgexp += other.formula_ptgexp;
        self.formula_ptgtbl += other.formula_ptgtbl;
        self.shrfmla_records += other.shrfmla_records;
        self.array_records += other.array_records;
        self.table_records += other.table_records;
        self.unresolved_ptgexp += other.unresolved_ptgexp;
        self.unresolved_ptgtbl += other.unresolved_ptgtbl;
        self.record_parse_errors += other.record_parse_errors;
        self.payload_parse_errors += other.payload_parse_errors;
    }
}

/// Diagnostics for one worksheet substream.
#[derive(Debug, Default, Clone, PartialEq, Eq)]
pub struct SheetFormulaDiagnostics {
    pub name: String,
    pub offset: usize,
    pub stats: SheetFormulaStats,
    /// Non-fatal parse errors encountered while scanning the sheet.
    pub errors: Vec<String>,
}

/// Diagnostics for an `.xls` workbook file.
#[derive(Debug, Default, Clone, PartialEq, Eq)]
pub struct WorkbookFormulaDiagnostics {
    pub sheets: Vec<SheetFormulaDiagnostics>,
    /// Non-fatal parse errors encountered while scanning workbook globals.
    pub errors: Vec<String>,
}

impl WorkbookFormulaDiagnostics {
    pub fn totals(&self) -> SheetFormulaStats {
        let mut out = SheetFormulaStats::default();
        for sheet in &self.sheets {
            out.add_assign(&sheet.stats);
        }
        out
    }

    pub fn has_errors(&self) -> bool {
        if !self.errors.is_empty() {
            return true;
        }
        self.sheets.iter().any(|s| !s.errors.is_empty())
    }
}

/// Collect shared/array/table formula statistics for a workbook.
///
/// Returns `Err` only for fatal I/O/container errors (e.g. the workbook stream cannot be read).
/// BIFF parse errors are returned in [`WorkbookFormulaDiagnostics::errors`] / sheet errors and
/// should be treated as a non-zero exit status by callers.
pub fn collect_xls_formula_diagnostics(path: &Path) -> Result<WorkbookFormulaDiagnostics, String> {
    let workbook_stream = biff::read_workbook_stream_from_xls(path)?;

    let biff_version = biff::detect_biff_version(&workbook_stream);
    let codepage = biff::parse_biff_codepage(&workbook_stream);

    let mut out = WorkbookFormulaDiagnostics::default();

    let bound_sheets = match biff::parse_biff_bound_sheets(&workbook_stream, biff_version, codepage)
    {
        Ok(v) => v,
        Err(err) => {
            out.errors
                .push(format!("failed to parse BoundSheet records: {err}"));
            return Ok(out);
        }
    };

    for (idx, sheet) in bound_sheets.into_iter().enumerate() {
        let (stats, errors) = collect_sheet_formula_stats(&workbook_stream, sheet.offset);
        out.sheets.push(SheetFormulaDiagnostics {
            name: if sheet.name.is_empty() {
                format!("Sheet{idx}")
            } else {
                sheet.name
            },
            offset: sheet.offset,
            stats,
            errors,
        });
    }

    Ok(out)
}

// Worksheet record ids we care about (BIFF8).
//
// `FORMULA`, `SHRFMLA`, and `ARRAY` are parsed via `biff::worksheet_formulas` so we can handle
// `CONTINUE` boundaries correctly.
const RECORD_TABLE: u16 = 0x0236;

/// Collect formula stats from a worksheet substream starting at `start`.
///
/// This function is best-effort: malformed records stop scanning but still return partial stats.
fn collect_sheet_formula_stats(
    workbook_stream: &[u8],
    start: usize,
) -> (SheetFormulaStats, Vec<String>) {
    let mut stats = SheetFormulaStats::default();
    let mut errors: Vec<String> = Vec::new();

    let mut shared_bases: HashSet<(u16, u16)> = HashSet::new();
    let mut array_bases: HashSet<(u16, u16)> = HashSet::new();
    let mut table_bases: HashSet<(u16, u16)> = HashSet::new();
    // Cells that contain an explicit (non-PtgExp/PtgTbl) rgce stream.
    let mut explicit_formula_cells: HashSet<(u16, u16)> = HashSet::new();

    // PtgExp/PtgTbl references found in FORMULA records.
    let mut ptgexp_refs: Vec<(u16, u16)> = Vec::new();
    let mut ptgtbl_refs: Vec<(u16, u16)> = Vec::new();

    let allows_continuation = |record_id: u16| {
        record_id == biff::worksheet_formulas::RECORD_FORMULA
            || record_id == biff::worksheet_formulas::RECORD_SHRFMLA
            || record_id == biff::worksheet_formulas::RECORD_ARRAY
            || record_id == RECORD_TABLE
    };
    let mut iter = match biff::records::LogicalBiffRecordIter::from_offset(
        workbook_stream,
        start,
        allows_continuation,
    ) {
        Ok(it) => it,
        Err(err) => {
            stats.record_parse_errors += 1;
            errors.push(err);
            return (stats, errors);
        }
    };

    while let Some(next) = iter.next() {
        let record = match next {
            Ok(r) => r,
            Err(err) => {
                stats.record_parse_errors += 1;
                errors.push(format!("malformed BIFF record: {err}"));
                break;
            }
        };

        // Stop before consuming the next substream.
        if record.offset != start && biff::records::is_bof_record(record.record_id) {
            break;
        }

        match record.record_id {
            biff::worksheet_formulas::RECORD_FORMULA => {
                stats.formula_records += 1;
                match biff::worksheet_formulas::parse_biff8_formula_record(&record) {
                    Ok(parsed) => {
                        if let Some(&ptg) = parsed.rgce.first() {
                            match ptg {
                                0x01 => {
                                    stats.formula_ptgexp += 1;
                                    if let Some(base) = parse_ptg_ref_cell(&parsed.rgce) {
                                        ptgexp_refs.push(base);
                                    } else {
                                        stats.payload_parse_errors += 1;
                                        errors.push(format!(
                                            "truncated PtgExp payload at FORMULA offset {}",
                                            record.offset
                                        ));
                                    }
                                }
                                0x02 => {
                                    stats.formula_ptgtbl += 1;
                                    if let Some(base) = parse_ptg_ref_cell(&parsed.rgce) {
                                        ptgtbl_refs.push(base);
                                    } else {
                                        stats.payload_parse_errors += 1;
                                        errors.push(format!(
                                            "truncated PtgTbl payload at FORMULA offset {}",
                                            record.offset
                                        ));
                                    }
                                }
                                _ => {
                                    explicit_formula_cells.insert((parsed.row, parsed.col));
                                }
                            }
                        }
                    }
                    Err(err) => {
                        stats.payload_parse_errors += 1;
                        errors.push(format!(
                            "failed to parse FORMULA record at offset {}: {err}",
                            record.offset
                        ));
                    }
                }
            }
            biff::worksheet_formulas::RECORD_SHRFMLA => {
                stats.shrfmla_records += 1;
                match biff::worksheet_formulas::parse_biff8_shrfmla_record(&record) {
                    Ok(_) => match parse_record_base_cell(record.data.as_ref()) {
                        Some(base) => {
                            shared_bases.insert(base);
                        }
                        None => {
                            stats.payload_parse_errors += 1;
                            errors.push(format!(
                                "failed to parse SHRFMLA base cell at offset {}",
                                record.offset
                            ));
                        }
                    },
                    Err(err) => {
                        stats.payload_parse_errors += 1;
                        errors.push(format!(
                            "failed to parse SHRFMLA record at offset {}: {err}",
                            record.offset
                        ));
                    }
                }
            }
            biff::worksheet_formulas::RECORD_ARRAY => {
                stats.array_records += 1;
                match biff::worksheet_formulas::parse_biff8_array_record(&record) {
                    Ok(_) => match parse_record_base_cell(record.data.as_ref()) {
                        Some(base) => {
                            array_bases.insert(base);
                        }
                        None => {
                            stats.payload_parse_errors += 1;
                            errors.push(format!(
                                "failed to parse ARRAY base cell at offset {}",
                                record.offset
                            ));
                        }
                    },
                    Err(err) => {
                        stats.payload_parse_errors += 1;
                        errors.push(format!(
                            "failed to parse ARRAY record at offset {}: {err}",
                            record.offset
                        ));
                    }
                }
            }
            RECORD_TABLE => {
                stats.table_records += 1;
                match parse_record_base_cell(record.data.as_ref()) {
                    Some(base) => {
                        table_bases.insert(base);
                    }
                    None => {
                        stats.payload_parse_errors += 1;
                        errors.push(format!(
                            "failed to parse TABLE base cell at offset {}",
                            record.offset
                        ));
                    }
                }
            }
            biff::records::RECORD_EOF => break,
            _ => {}
        }
    }

    for base in ptgexp_refs {
        let resolved = shared_bases.contains(&base)
            || array_bases.contains(&base)
            || explicit_formula_cells.contains(&base);
        if !resolved {
            stats.unresolved_ptgexp += 1;
        }
    }

    for base in ptgtbl_refs {
        let resolved = table_bases.contains(&base) || explicit_formula_cells.contains(&base);
        if !resolved {
            stats.unresolved_ptgtbl += 1;
        }
    }

    (stats, errors)
}

fn parse_ptg_ref_cell(rgce: &[u8]) -> Option<(u16, u16)> {
    if rgce.len() < 5 {
        return None;
    }
    let row = u16::from_le_bytes([rgce[1], rgce[2]]);
    let col = u16::from_le_bytes([rgce[3], rgce[4]]);
    Some((row, col))
}

/// Best-effort extraction of the "base cell" from records that start with a `RefU`/`Ref8` range.
///
/// `SHRFMLA`, `ARRAY`, and `TABLE` records all begin with a range that anchors the construct.
fn parse_record_base_cell(data: &[u8]) -> Option<(u16, u16)> {
    let (rw_first, _rw_last, col_first, _col_last) = parse_ref_any(data)?;
    Some((rw_first, col_first))
}

fn parse_ref_any(data: &[u8]) -> Option<(u16, u16, u16, u16)> {
    // Prefer Ref8 (u16 columns) when it decodes to "classic" `.xls` column bounds.
    if let Some(r) = parse_ref8(data) {
        if r.2 <= 0x00FF && r.3 <= 0x00FF {
            return Some(r);
        }
    }

    parse_refu(data).or_else(|| parse_ref8(data))
}

fn parse_ref8(data: &[u8]) -> Option<(u16, u16, u16, u16)> {
    let chunk = data.get(0..8)?;
    let rw_first = u16::from_le_bytes([chunk[0], chunk[1]]);
    let rw_last = u16::from_le_bytes([chunk[2], chunk[3]]);
    let col_first = u16::from_le_bytes([chunk[4], chunk[5]]);
    let col_last = u16::from_le_bytes([chunk[6], chunk[7]]);
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

#[cfg(test)]
mod tests {
    use super::*;

    fn record(id: u16, payload: &[u8]) -> Vec<u8> {
        let mut out = Vec::new();
        out.extend_from_slice(&id.to_le_bytes());
        out.extend_from_slice(&(payload.len() as u16).to_le_bytes());
        out.extend_from_slice(payload);
        out
    }

    fn formula_record(row: u16, col: u16, rgce: &[u8]) -> Vec<u8> {
        let mut payload = Vec::new();
        payload.extend_from_slice(&row.to_le_bytes());
        payload.extend_from_slice(&col.to_le_bytes());
        payload.extend_from_slice(&0u16.to_le_bytes()); // ixfe
        payload.extend_from_slice(&[0u8; 8]); // result
        payload.extend_from_slice(&0u16.to_le_bytes()); // grbit
        payload.extend_from_slice(&0u32.to_le_bytes()); // chn
        payload.extend_from_slice(&(rgce.len() as u16).to_le_bytes()); // cce
        payload.extend_from_slice(rgce);
        record(biff::worksheet_formulas::RECORD_FORMULA, &payload)
    }

    #[test]
    fn collects_shared_array_table_formula_stats_from_synthetic_sheet_stream() {
        // Synthetic worksheet substream:
        // - SHRFMLA base @ (0,0) (valid payload; used to resolve one PtgExp)
        // - TABLE base @ (2,2)
        // - 4 FORMULA records:
        //   - (0,1) PtgExp -> (0,0) (resolved)
        //   - (1,0) PtgExp -> (9,9) (unresolved)
        //   - (2,2) PtgTbl -> (2,2) (resolved)
        //   - (3,3) PtgTbl -> (6,6) (unresolved)

        // SHRFMLA: Ref8 (rows 0..0, cols 0..1) + cUse + cce + rgce.
        let shrfmla_rgce = vec![0x16]; // PtgMissArg (simple 1-byte token)
        let mut shrfmla_payload = Vec::new();
        shrfmla_payload.extend_from_slice(&0u16.to_le_bytes()); // rwFirst
        shrfmla_payload.extend_from_slice(&0u16.to_le_bytes()); // rwLast
        shrfmla_payload.extend_from_slice(&0u16.to_le_bytes()); // colFirst
        shrfmla_payload.extend_from_slice(&1u16.to_le_bytes()); // colLast
        shrfmla_payload.extend_from_slice(&0u16.to_le_bytes()); // cUse
        shrfmla_payload.extend_from_slice(&(shrfmla_rgce.len() as u16).to_le_bytes()); // cce
        shrfmla_payload.extend_from_slice(&shrfmla_rgce);
        let table_payload = [
            2u16.to_le_bytes(),
            2u16.to_le_bytes(),
            2u16.to_le_bytes(),
            2u16.to_le_bytes(),
        ]
        .concat(); // Ref8: rows 2..2, cols 2..2

        let mut stream: Vec<u8> = Vec::new();
        stream.extend(record(
            biff::worksheet_formulas::RECORD_SHRFMLA,
            &shrfmla_payload,
        ));
        stream.extend(record(RECORD_TABLE, &table_payload));

        // PtgExp: [0x01][rw:u16][col:u16]
        stream.extend(formula_record(0, 1, &[0x01, 0x00, 0x00, 0x00, 0x00]));
        stream.extend(formula_record(1, 0, &[0x01, 0x09, 0x00, 0x09, 0x00]));

        // PtgTbl: [0x02][rw:u16][col:u16]
        stream.extend(formula_record(2, 2, &[0x02, 0x02, 0x00, 0x02, 0x00]));
        stream.extend(formula_record(3, 3, &[0x02, 0x06, 0x00, 0x06, 0x00]));

        stream.extend(record(biff::records::RECORD_EOF, &[]));

        let (stats, errors) = collect_sheet_formula_stats(&stream, 0);
        assert!(errors.is_empty(), "expected no errors, got: {errors:?}");

        assert_eq!(
            stats,
            SheetFormulaStats {
                formula_records: 4,
                formula_ptgexp: 2,
                formula_ptgtbl: 2,
                shrfmla_records: 1,
                array_records: 0,
                table_records: 1,
                unresolved_ptgexp: 1,
                unresolved_ptgtbl: 1,
                record_parse_errors: 0,
                payload_parse_errors: 0,
            }
        );
    }
}
