use std::collections::{BTreeMap, BTreeSet};
use std::fs;
use std::path::PathBuf;
use std::time::Instant;

use anyhow::{Context, Result};
use clap::Parser;
use formula_engine::{Engine, ErrorKind, NameDefinition, NameScope, Value as EngineValue};
use formula_model::{CellRef, CellValue, DefinedNameScope, ErrorValue};
use serde::Serialize;
use sha2::{Digest, Sha256};

#[derive(Parser, Debug)]
#[command(about = "Compatibility triage helper used by tools/corpus/triage.py")]
struct Args {
    /// Input workbook (XLSX/XLSM).
    #[arg(long)]
    input: PathBuf,

    /// Parts to ignore when diffing round-tripped output.
    #[arg(long = "ignore-part")]
    ignore_parts: Vec<String>,

    /// Maximum number of differences to emit (privacy-safe summary only).
    #[arg(long, default_value_t = 25)]
    diff_limit: usize,

    /// Run a best-effort recalculation check against cached workbook values.
    #[arg(long)]
    recalc: bool,

    /// Run a lightweight headless render/pagination smoke test.
    #[arg(long = "render-smoke")]
    render_smoke: bool,
}

#[derive(Debug, Serialize)]
struct StepResult {
    status: String, // ok | failed | skipped
    #[serde(skip_serializing_if = "Option::is_none")]
    duration_ms: Option<u128>,
    #[serde(skip_serializing_if = "Option::is_none")]
    error: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    details: Option<serde_json::Value>,
}

impl StepResult {
    fn ok(start: Instant, details: impl Serialize) -> Self {
        StepResult {
            status: "ok".to_string(),
            duration_ms: Some(start.elapsed().as_millis()),
            error: None,
            details: Some(serde_json::to_value(details).unwrap_or(serde_json::Value::Null)),
        }
    }

    fn failed(start: Instant, err: impl ToString) -> Self {
        // Triage reports are uploaded as artifacts for both public and private corpora. Avoid
        // leaking workbook content (sheet names, defined names, etc.) through error strings by
        // hashing the message and emitting only the digest.
        let sha = sha256_text(&err.to_string());
        StepResult {
            status: "failed".to_string(),
            duration_ms: Some(start.elapsed().as_millis()),
            error: Some(format!("sha256={sha}")),
            details: None,
        }
    }

    fn skipped(reason: impl ToString) -> Self {
        StepResult {
            status: "skipped".to_string(),
            duration_ms: None,
            error: None,
            details: Some(serde_json::json!({ "reason": reason.to_string() })),
        }
    }
}

#[derive(Debug, Serialize, Default)]
#[serde(rename_all = "snake_case")]
struct TriageResult {
    open_ok: bool,
    round_trip_ok: bool,
    diff_critical_count: usize,
    diff_warning_count: usize,
    diff_info_count: usize,
    diff_total_count: usize,
    calculate_ok: Option<bool>,
    render_ok: Option<bool>,
}

#[derive(Debug, Serialize)]
struct TriageOutput {
    steps: BTreeMap<String, StepResult>,
    result: TriageResult,
}

#[derive(Debug, Serialize)]
struct LoadDetails {
    engine: &'static str,
    parts: usize,
    sheets: usize,
}

#[derive(Debug, Serialize)]
struct RoundTripDetails {
    engine: &'static str,
    output_size_bytes: usize,
    output_parts: usize,
}

#[derive(Debug, Serialize)]
struct DiffCounts {
    critical: usize,
    warning: usize,
    info: usize,
    total: usize,
}

#[derive(Debug, Serialize)]
struct DiffEntry {
    severity: String,
    part: String,
    path: String,
    kind: String,
}

#[derive(Debug, Serialize)]
struct DiffDetails {
    ignore: Vec<String>,
    counts: DiffCounts,
    equal: bool,
    top_differences: Vec<DiffEntry>,
}

#[derive(Debug, Serialize)]
struct SheetRecalcSummary {
    sheet_index: usize,
    formula_cell_count: usize,
    mismatch_count: usize,
    baseline_hash: String,
    computed_hash: String,
}

#[derive(Debug, Serialize)]
struct RecalcDetails {
    sheet_count: usize,
    formula_cell_count: usize,
    mismatch_count: usize,
    sheets: Vec<SheetRecalcSummary>,
}

#[derive(Debug, Serialize)]
struct RenderDetails {
    sheet_index: usize,
    pages: usize,
    pdf_size_bytes: usize,
    print_area: RenderPrintArea,
}

#[derive(Debug, Serialize)]
struct RenderPrintArea {
    start_row: u32,
    end_row: u32,
    start_col: u32,
    end_col: u32,
}

fn main() -> Result<()> {
    let args = Args::parse();

    let mut output = TriageOutput {
        steps: BTreeMap::new(),
        result: TriageResult::default(),
    };

    let input_bytes = {
        let start = Instant::now();
        match fs::read(&args.input)
            .with_context(|| format!("read workbook {}", args.input.display()))
        {
            Ok(bytes) => bytes,
            Err(err) => {
                output
                    .steps
                    .insert("load".to_string(), StepResult::failed(start, err));
                print_json(&output)?;
                return Ok(());
            }
        }
    };

    // Step: load (formula-xlsx)
    let start = Instant::now();
    let doc = match formula_xlsx::load_from_bytes(&input_bytes) {
        Ok(doc) => doc,
        Err(err) => {
            output
                .steps
                .insert("load".to_string(), StepResult::failed(start, err));
            output.result.open_ok = false;
            output
                .steps
                .insert("recalc".to_string(), StepResult::skipped("open_failed"));
            output
                .steps
                .insert("render".to_string(), StepResult::skipped("open_failed"));
            output
                .steps
                .insert("round_trip".to_string(), StepResult::skipped("open_failed"));
            output
                .steps
                .insert("diff".to_string(), StepResult::skipped("open_failed"));
            print_json(&output)?;
            return Ok(());
        }
    };

    output.steps.insert(
        "load".to_string(),
        StepResult::ok(
            start,
            LoadDetails {
                engine: "formula-xlsx",
                parts: doc.parts().len(),
                sheets: doc.workbook.sheets.len(),
            },
        ),
    );
    output.result.open_ok = true;

    // Step: round-trip save (formula-xlsx)
    let start = Instant::now();
    let round_tripped = match doc.save_to_vec() {
        Ok(bytes) => bytes,
        Err(err) => {
            output
                .steps
                .insert("round_trip".to_string(), StepResult::failed(start, err));
            output.result.round_trip_ok = false;
            output
                .steps
                .insert("diff".to_string(), StepResult::skipped("round_trip_failed"));
            output.steps.insert(
                "recalc".to_string(),
                StepResult::skipped("round_trip_failed"),
            );
            output.steps.insert(
                "render".to_string(),
                StepResult::skipped("round_trip_failed"),
            );
            print_json(&output)?;
            return Ok(());
        }
    };

    let output_parts = xlsx_diff::WorkbookArchive::from_bytes(&round_tripped)
        .map(|a| a.part_names().len())
        .unwrap_or(0);

    output.steps.insert(
        "round_trip".to_string(),
        StepResult::ok(
            start,
            RoundTripDetails {
                engine: "formula-xlsx",
                output_size_bytes: round_tripped.len(),
                output_parts,
            },
        ),
    );

    // Step: diff (xlsx-diff)
    let start = Instant::now();
    let diff_details = match diff_workbooks(&input_bytes, &round_tripped, &args) {
        Ok(details) => details,
        Err(err) => {
            output
                .steps
                .insert("diff".to_string(), StepResult::failed(start, err));
            output.result.round_trip_ok = false;
            print_json(&output)?;
            return Ok(());
        }
    };

    output.result.diff_critical_count = diff_details.counts.critical;
    output.result.diff_warning_count = diff_details.counts.warning;
    output.result.diff_info_count = diff_details.counts.info;
    output.result.diff_total_count = diff_details.counts.total;
    output.result.round_trip_ok = diff_details.counts.critical == 0;

    output
        .steps
        .insert("diff".to_string(), StepResult::ok(start, &diff_details));

    // Step: recalc (optional)
    if !args.recalc {
        output.steps.insert(
            "recalc".to_string(),
            StepResult::skipped("disabled (pass --recalc)"),
        );
        output.result.calculate_ok = None;
    } else {
        let start = Instant::now();
        match recalc_against_cached(&doc) {
            Ok(Some(recalc)) => {
                output.result.calculate_ok = Some(recalc.mismatch_count == 0);
                output
                    .steps
                    .insert("recalc".to_string(), StepResult::ok(start, &recalc));
            }
            Ok(None) => {
                output.steps.insert(
                    "recalc".to_string(),
                    StepResult::skipped("no_cached_formula_values_or_no_formulas"),
                );
                output.result.calculate_ok = None;
            }
            Err(err) => {
                output.steps.insert(
                    "recalc".to_string(),
                    StepResult::skipped(format!(
                        "engine_error (sha256={})",
                        sha256_text(&err.to_string())
                    )),
                );
                output.result.calculate_ok = None;
            }
        }
    }

    // Step: render smoke (optional)
    if !args.render_smoke {
        output.steps.insert(
            "render".to_string(),
            StepResult::skipped("disabled (pass --render-smoke)"),
        );
        output.result.render_ok = None;
    } else {
        let start = Instant::now();
        match render_smoke(&doc) {
            Ok(details) => {
                output.result.render_ok = Some(details.pdf_size_bytes > 0);
                output
                    .steps
                    .insert("render".to_string(), StepResult::ok(start, details));
            }
            Err(err) => {
                output.result.render_ok = Some(false);
                output
                    .steps
                    .insert("render".to_string(), StepResult::failed(start, err));
            }
        }
    }

    print_json(&output)?;
    Ok(())
}

fn print_json(output: &TriageOutput) -> Result<()> {
    let json = serde_json::to_string(output).context("serialize triage output")?;
    println!("{json}");
    Ok(())
}

fn diff_workbooks(expected: &[u8], actual: &[u8], args: &Args) -> Result<DiffDetails> {
    let ignore: BTreeSet<String> = args
        .ignore_parts
        .iter()
        .map(|s| s.trim())
        .map(|s| s.replace('\\', "/"))
        .map(|s| s.trim_start_matches('/').to_string())
        .filter(|s| !s.is_empty())
        .collect();
    let mut ignore_sorted: Vec<String> = ignore.iter().cloned().collect();
    ignore_sorted.sort();

    let expected_archive = xlsx_diff::WorkbookArchive::from_bytes(expected)?;
    let actual_archive = xlsx_diff::WorkbookArchive::from_bytes(actual)?;

    let report = xlsx_diff::diff_archives_with_options(
        &expected_archive,
        &actual_archive,
        &xlsx_diff::DiffOptions {
            ignore_parts: ignore,
            ignore_globs: Vec::new(),
        },
    );

    let counts = DiffCounts {
        critical: report.count(xlsx_diff::Severity::Critical),
        warning: report.count(xlsx_diff::Severity::Warning),
        info: report.count(xlsx_diff::Severity::Info),
        total: report.differences.len(),
    };

    let equal = report.differences.is_empty();

    let mut entries: Vec<DiffEntry> = report
        .differences
        .into_iter()
        .map(|d| DiffEntry {
            severity: d.severity.as_str().to_string(),
            part: d.part,
            path: d.path,
            kind: d.kind,
        })
        .collect();

    // Stable output order: (severity, part, path, kind).
    entries.sort_by(|a, b| {
        let rank = |s: &str| match s {
            "CRITICAL" => 0u8,
            "WARN" => 1u8,
            "INFO" => 2u8,
            _ => 3u8,
        };
        (
            rank(a.severity.as_str()),
            a.part.as_str(),
            a.path.as_str(),
            a.kind.as_str(),
        )
            .cmp(&(
                rank(b.severity.as_str()),
                b.part.as_str(),
                b.path.as_str(),
                b.kind.as_str(),
            ))
    });

    entries.truncate(args.diff_limit);

    Ok(DiffDetails {
        ignore: ignore_sorted,
        counts,
        equal,
        top_differences: entries,
    })
}

fn recalc_against_cached(doc: &formula_xlsx::XlsxDocument) -> Result<Option<RecalcDetails>> {
    let sheet_count = doc.workbook.sheets.len();
    if sheet_count == 0 {
        return Ok(None);
    }

    // Collect formula cells + cached baseline values.
    let mut formula_cells_by_sheet: Vec<Vec<(CellRef, CellValue)>> = Vec::new();
    let mut total_formula_cells = 0usize;
    let mut any_cached_non_blank = false;

    for sheet in &doc.workbook.sheets {
        let mut cells: Vec<(CellRef, &formula_model::Cell)> = sheet.iter_cells().collect();
        cells.sort_by_key(|(r, _)| (r.row, r.col));

        let mut formula_cells: Vec<(CellRef, CellValue)> = Vec::new();
        for (cell_ref, cell) in cells {
            if cell.formula.is_some() {
                total_formula_cells += 1;
                if !matches!(cell.value, CellValue::Empty) {
                    any_cached_non_blank = true;
                }
                formula_cells.push((cell_ref, cell.value.clone()));
            }
        }
        formula_cells_by_sheet.push(formula_cells);
    }

    if total_formula_cells == 0 || !any_cached_non_blank {
        return Ok(None);
    }

    let mut engine = Engine::new();
    engine.set_date_system(match doc.workbook.date_system {
        formula_model::DateSystem::Excel1900 => formula_engine::date::ExcelDateSystem::EXCEL_1900,
        formula_model::DateSystem::Excel1904 => formula_engine::date::ExcelDateSystem::Excel1904,
    });

    // Ensure sheets exist up-front so cross-sheet references compile correctly.
    for sheet in &doc.workbook.sheets {
        engine.ensure_sheet(&sheet.name);
    }

    // Best-effort defined name support (named ranges/constants) used by many workbooks.
    if !doc.workbook.defined_names.is_empty() {
        let mut sheet_names_by_id = BTreeMap::new();
        for sheet in &doc.workbook.sheets {
            sheet_names_by_id.insert(sheet.id, sheet.name.as_str());
        }

        for name in &doc.workbook.defined_names {
            let scope = match name.scope {
                DefinedNameScope::Workbook => NameScope::Workbook,
                DefinedNameScope::Sheet(sheet_id) => {
                    let sheet_name = sheet_names_by_id
                        .get(&sheet_id)
                        .context("defined name references unknown worksheet id")?;
                    NameScope::Sheet(sheet_name)
                }
            };

            // Defined names are stored without a leading '=' in `formula-model`; the engine parser
            // accepts canonical formula/refs in the same format.
            engine.define_name(
                &name.name,
                scope,
                NameDefinition::Formula(name.refers_to.clone()),
            )?;
        }
    }

    // Feed values + formulas into the engine.
    for sheet in &doc.workbook.sheets {
        if !sheet.tables.is_empty() {
            engine.set_sheet_tables(&sheet.name, sheet.tables.clone());
        }

        let mut cells: Vec<(CellRef, &formula_model::Cell)> = sheet.iter_cells().collect();
        cells.sort_by_key(|(r, _)| (r.row, r.col));

        let mut formulas: Vec<(String, String)> = Vec::new();
        for (cell_ref, cell) in cells {
            let a1 = cell_ref.to_a1();
            if let Some(formula) = &cell.formula {
                formulas.push((a1, formula.clone()));
                continue;
            }
            set_engine_value(&mut engine, &sheet.name, &a1, &cell.value)?;
        }

        for (a1, formula) in formulas {
            engine.set_cell_formula(&sheet.name, &a1, &formula)?;
        }
    }

    engine.recalculate();

    let mut sheets_out: Vec<SheetRecalcSummary> = Vec::new();
    let mut mismatch_total = 0usize;

    for (sheet_index, sheet) in doc.workbook.sheets.iter().enumerate() {
        let formula_cells = &formula_cells_by_sheet[sheet_index];
        if formula_cells.is_empty() {
            continue;
        }

        let mut baseline_hasher = Sha256::new();
        let mut computed_hasher = Sha256::new();
        let mut mismatch_count = 0usize;

        for (cell_ref, baseline_value) in formula_cells {
            let addr = cell_ref.to_a1();
            let computed_value = engine.get_cell_value(&sheet.name, &addr);

            // Hash baseline vs computed in a stable, typed encoding.
            hash_cell_value(
                &mut baseline_hasher,
                &addr,
                &normalize_model_value(baseline_value),
            );
            hash_cell_value(
                &mut computed_hasher,
                &addr,
                &normalize_engine_value(&computed_value),
            );

            if normalize_engine_value(&computed_value) != normalize_model_value(baseline_value) {
                mismatch_count += 1;
            }
        }

        mismatch_total += mismatch_count;

        sheets_out.push(SheetRecalcSummary {
            sheet_index: sheet_index + 1,
            formula_cell_count: formula_cells.len(),
            mismatch_count,
            baseline_hash: format!("{:x}", baseline_hasher.finalize()),
            computed_hash: format!("{:x}", computed_hasher.finalize()),
        });
    }

    Ok(Some(RecalcDetails {
        sheet_count,
        formula_cell_count: total_formula_cells,
        mismatch_count: mismatch_total,
        sheets: sheets_out,
    }))
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum NormalizedValue {
    Blank,
    Number(u64),
    Bool(bool),
    Text(String),
    Error(String),
}

fn normalize_engine_value(value: &EngineValue) -> NormalizedValue {
    match value {
        EngineValue::Blank => NormalizedValue::Blank,
        EngineValue::Number(n) => NormalizedValue::Number(n.to_bits()),
        EngineValue::Bool(b) => NormalizedValue::Bool(*b),
        EngineValue::Text(s) => NormalizedValue::Text(s.clone()),
        EngineValue::Entity(v) => NormalizedValue::Text(v.display.clone()),
        EngineValue::Record(v) => NormalizedValue::Text(v.display.clone()),
        EngineValue::Error(e) => NormalizedValue::Error(e.as_code().to_string()),
        EngineValue::Reference(_)
        | EngineValue::ReferenceUnion(_)
        | EngineValue::Array(_)
        | EngineValue::Spill { .. } => NormalizedValue::Blank,
        EngineValue::Lambda(_) => NormalizedValue::Error(ErrorKind::Calc.as_code().to_string()),
    }
}

fn normalize_model_value(value: &CellValue) -> NormalizedValue {
    match value {
        CellValue::Empty => NormalizedValue::Blank,
        CellValue::Number(n) => NormalizedValue::Number(n.to_bits()),
        CellValue::Boolean(b) => NormalizedValue::Bool(*b),
        CellValue::String(s) => NormalizedValue::Text(s.clone()),
        CellValue::Error(e) => NormalizedValue::Error(map_error_value(*e).as_code().to_string()),
        CellValue::RichText(r) => NormalizedValue::Text(r.text.clone()),
        CellValue::Entity(e) => NormalizedValue::Text(e.display_value.clone()),
        CellValue::Record(r) => NormalizedValue::Text(r.to_string()),
        CellValue::Image(image) => NormalizedValue::Text(
            image
                .alt_text
                .clone()
                .filter(|s| !s.is_empty())
                .unwrap_or_else(|| "[Image]".to_string()),
        ),
        CellValue::Array(_) | CellValue::Spill(_) => NormalizedValue::Blank,
    }
}

fn map_error_value(error: ErrorValue) -> ErrorKind {
    ErrorKind::from_code(error.as_str()).unwrap_or(ErrorKind::Value)
}

fn set_engine_value(engine: &mut Engine, sheet: &str, addr: &str, value: &CellValue) -> Result<()> {
    let v = match value {
        CellValue::Empty => EngineValue::Blank,
        CellValue::Number(n) => EngineValue::Number(*n),
        CellValue::String(s) => EngineValue::Text(s.clone()),
        CellValue::Boolean(b) => EngineValue::Bool(*b),
        CellValue::Error(e) => EngineValue::Error(map_error_value(*e)),
        CellValue::RichText(r) => EngineValue::Text(r.text.clone()),
        CellValue::Entity(e) => EngineValue::Text(e.display_value.clone()),
        CellValue::Record(r) => EngineValue::Text(r.to_string()),
        CellValue::Image(image) => EngineValue::Text(
            image
                .alt_text
                .clone()
                .filter(|s| !s.is_empty())
                .unwrap_or_else(|| "[Image]".to_string()),
        ),
        CellValue::Array(_) | CellValue::Spill(_) => EngineValue::Blank,
    };
    engine
        .set_cell_value(sheet, addr, v)
        .context("set cell value")?;
    Ok(())
}

fn hash_cell_value(hasher: &mut Sha256, addr: &str, value: &NormalizedValue) {
    hasher.update(addr.as_bytes());
    hasher.update([0u8]);
    match value {
        NormalizedValue::Blank => {
            hasher.update([b'Z']);
        }
        NormalizedValue::Number(bits) => {
            hasher.update([b'N']);
            hasher.update(bits.to_le_bytes());
        }
        NormalizedValue::Bool(b) => {
            hasher.update([b'B', if *b { 1 } else { 0 }]);
        }
        NormalizedValue::Text(s) => {
            hasher.update([b'S']);
            let bytes = s.as_bytes();
            hasher.update((bytes.len() as u64).to_le_bytes());
            hasher.update(bytes);
        }
        NormalizedValue::Error(code) => {
            hasher.update([b'E']);
            hasher.update(code.as_bytes());
        }
    }
    hasher.update([0u8]);
}

fn sha256_text(text: &str) -> String {
    let mut hasher = Sha256::new();
    hasher.update(text.as_bytes());
    format!("{:x}", hasher.finalize())
}

fn render_smoke(doc: &formula_xlsx::XlsxDocument) -> Result<RenderDetails> {
    let sheet = doc
        .workbook
        .sheets
        .first()
        .context("workbook has no sheets")?;

    let used = sheet
        .used_range()
        .unwrap_or_else(|| formula_model::Range::new(CellRef::new(0, 0), CellRef::new(0, 0)));

    let max_rows = 20u32;
    let max_cols = 10u32;

    // `used` stores 0-based row/col indexes. Convert to the 1-based coordinates expected by the
    // print subsystem using saturating math so we don't panic on extremely large indexes
    // (e.g. u32::MAX).
    let start_row = used.start.row.saturating_add(1);
    let start_col = used.start.col.saturating_add(1);
    let end_row = used
        .end
        .row
        .saturating_add(1)
        .min(start_row.saturating_add(max_rows.saturating_sub(1)));
    let end_col = used
        .end
        .col
        .saturating_add(1)
        .min(start_col.saturating_add(max_cols.saturating_sub(1)));

    let print_area = formula_xlsx::print::CellRange {
        start_row,
        end_row,
        start_col,
        end_col,
    };

    // Keep the smoke test cheap even for workbooks whose used range starts at very large
    // row/column indexes. The print APIs treat missing widths/heights as 0.0 and we don't
    // render cell text anyway, so a small fixed buffer is enough to validate "no panic +
    // non-empty PDF output".
    let col_widths_points = vec![64.0; max_cols as usize];
    let row_heights_points = vec![15.0; max_rows as usize];

    let page_setup = formula_xlsx::print::PageSetup::default();
    let manual_breaks = formula_xlsx::print::ManualPageBreaks::default();

    let pages = formula_xlsx::print::calculate_pages(
        print_area,
        &col_widths_points,
        &row_heights_points,
        &page_setup,
        &manual_breaks,
    );

    let pdf = formula_xlsx::print::export_range_to_pdf_bytes(
        "Sheet",
        print_area,
        &col_widths_points,
        &row_heights_points,
        &page_setup,
        &manual_breaks,
        |_row, _col| None,
    )?;

    if pdf.is_empty() {
        anyhow::bail!("pdf output was empty");
    }

    Ok(RenderDetails {
        sheet_index: 1,
        pages: pages.len(),
        pdf_size_bytes: pdf.len(),
        print_area: RenderPrintArea {
            start_row: print_area.start_row,
            end_row: print_area.end_row,
            start_col: print_area.start_col,
            end_col: print_area.end_col,
        },
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn preserves_extended_errors_in_normalization_and_engine_seeding() {
        let mut engine = Engine::new();
        engine.ensure_sheet("Sheet1");

        for (addr, value, expected_code) in [
            ("A1", CellValue::Error(ErrorValue::Field), "#FIELD!"),
            ("A2", CellValue::Error(ErrorValue::GettingData), "#GETTING_DATA"),
        ] {
            set_engine_value(&mut engine, "Sheet1", addr, &value).unwrap();

            let seeded = engine.get_cell_value("Sheet1", addr);
            assert_eq!(
                normalize_engine_value(&seeded),
                normalize_model_value(&value),
                "engine vs model mismatch for {expected_code}"
            );

            assert_eq!(
                normalize_engine_value(&seeded),
                NormalizedValue::Error(expected_code.to_string())
            );
        }
    }

    #[test]
    fn normalizes_model_entity_display_value_and_record_display_field() {
        let entity = formula_model::EntityValue::new("AAPL");
        assert_eq!(
            normalize_model_value(&CellValue::Entity(entity.clone())),
            NormalizedValue::Text("AAPL".to_string())
        );

        let mut record = formula_model::RecordValue::default();
        record.display_value = "fallback".to_string();
        record.display_field = Some("Name".to_string());
        record
            .fields
            .insert("Name".to_string(), CellValue::String("Alice".to_string()));

        assert_eq!(
            normalize_model_value(&CellValue::Record(record.clone())),
            NormalizedValue::Text("Alice".to_string())
        );

        let mut engine = Engine::new();
        engine.ensure_sheet("Sheet1");

        set_engine_value(&mut engine, "Sheet1", "A1", &CellValue::Entity(entity)).unwrap();
        set_engine_value(&mut engine, "Sheet1", "A2", &CellValue::Record(record)).unwrap();

        assert_eq!(
            normalize_engine_value(&engine.get_cell_value("Sheet1", "A1")),
            NormalizedValue::Text("AAPL".to_string())
        );
        assert_eq!(
            normalize_engine_value(&engine.get_cell_value("Sheet1", "A2")),
            NormalizedValue::Text("Alice".to_string())
        );
    }
}
