use anyhow::{anyhow, Context, Result};
use clap::Parser;
use formula_engine::{Engine, Value};
use formula_model::cell_to_a1;
use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};
use std::collections::HashSet;
use std::fs;
use std::path::{Path, PathBuf};
use time::format_description::well_known::Rfc3339;

#[derive(Debug, Parser)]
#[command(
    name = "formula-excel-oracle",
    about = "Evaluate the Excel oracle corpus using the formula-engine and emit results JSON"
)]
struct Args {
    /// Path to the canonical oracle case corpus (cases.json).
    #[arg(long)]
    cases: PathBuf,

    /// Path to write results JSON (engine-results.json).
    #[arg(long)]
    out: PathBuf,

    /// Optional cap for debugging (evaluate only the first N cases).
    #[arg(long, default_value_t = 0)]
    max_cases: usize,

    /// Only include cases containing this tag (can be repeated).
    #[arg(long = "include-tag")]
    include_tags: Vec<String>,

    /// Exclude cases containing this tag (can be repeated).
    #[arg(long = "exclude-tag")]
    exclude_tags: Vec<String>,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct CaseCorpus {
    schema_version: u32,
    case_set: String,
    default_sheet: String,
    cases: Vec<Case>,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct Case {
    id: String,
    formula: String,
    output_cell: String,
    inputs: Vec<CellInput>,
    #[serde(default)]
    tags: Vec<String>,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct CellInput {
    cell: String,
    #[serde(default)]
    value: Option<serde_json::Value>,
    #[serde(default)]
    formula: Option<String>,
}

#[derive(Debug, Serialize)]
#[serde(tag = "t")]
#[allow(dead_code)]
enum EncodedValue {
    #[serde(rename = "blank")]
    Blank,
    #[serde(rename = "n")]
    Number { v: f64 },
    #[serde(rename = "s")]
    String { v: String },
    #[serde(rename = "b")]
    Bool { v: bool },
    #[serde(rename = "e")]
    Error {
        v: String,
        #[serde(skip_serializing_if = "Option::is_none")]
        detail: Option<String>,
    },
    #[serde(rename = "arr")]
    Array { rows: Vec<Vec<EncodedValue>> },
}

impl EncodedValue {
    fn engine_error(message: impl Into<String>) -> Self {
        EncodedValue::Error {
            v: "#ENGINE!".to_string(),
            detail: Some(message.into()),
        }
    }
}

impl From<Value> for EncodedValue {
    fn from(value: Value) -> Self {
        match value {
            Value::Blank => EncodedValue::Blank,
            Value::Number(v) if v.is_finite() => EncodedValue::Number { v },
            Value::Number(v) => {
                EncodedValue::engine_error(format!("non-finite numeric result: {v}"))
            }
            Value::Text(v) => EncodedValue::String { v },
            Value::Entity(v) => EncodedValue::String { v: v.display },
            Value::Record(v) => EncodedValue::String { v: v.display },
            Value::Bool(v) => EncodedValue::Bool { v },
            Value::Error(kind) => EncodedValue::Error {
                v: kind.as_code().to_string(),
                detail: None,
            },
            Value::Lambda(_) => EncodedValue::Error {
                v: "#CALC!".to_string(),
                detail: None,
            },
            Value::Array(arr) => {
                let mut rows = Vec::new();
                let _ = rows.try_reserve_exact(arr.rows);
                for r in 0..arr.rows {
                    let mut row = Vec::new();
                    let _ = row.try_reserve_exact(arr.cols);
                    for c in 0..arr.cols {
                        row.push(arr.get(r, c).cloned().unwrap_or(Value::Blank).into());
                    }
                    rows.push(row);
                }
                EncodedValue::Array { rows }
            }
            Value::Reference(_) | Value::ReferenceUnion(_) => {
                EncodedValue::engine_error("unexpected reference value")
            }
            Value::Spill { .. } => EncodedValue::Error {
                v: "#SPILL!".to_string(),
                detail: None,
            },
        }
    }
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
struct ResultsFile {
    schema_version: u32,
    generated_at: String,
    source: SourceInfo,
    case_set: CaseSetInfo,
    results: Vec<ResultEntry>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
struct SourceInfo {
    kind: String,
    version: String,
    os: String,
    arch: String,
    case_set: String,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
struct CaseSetInfo {
    path: String,
    sha256: String,
    count: usize,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
struct ResultEntry {
    case_id: String,
    output_cell: String,
    result: EncodedValue,
    address: String,
    display_text: String,
}

fn sha256_file(path: &Path) -> Result<String> {
    let bytes = fs::read(path).with_context(|| format!("read {}", path.display()))?;
    let mut hasher = Sha256::new();
    hasher.update(bytes);
    Ok(format!("{:x}", hasher.finalize()))
}

fn encode_json_value(value: serde_json::Value) -> Result<Value> {
    match value {
        serde_json::Value::Null => Ok(Value::Blank),
        serde_json::Value::Bool(b) => Ok(Value::Bool(b)),
        serde_json::Value::Number(n) => n
            .as_f64()
            .ok_or_else(|| anyhow!("numeric value out of range for f64"))
            .map(Value::Number),
        serde_json::Value::String(s) => Ok(Value::Text(s)),
        other => Err(anyhow!("unsupported input value type: {other}")),
    }
}

fn coord_to_a1(row: u32, col: u32) -> String {
    // Prefer the shared A1 formatter so very large row/col indices (e.g. u32::MAX) do not
    // overflow when converting from 0-based internal coordinates to 1-based A1 notation.
    cell_to_a1(row, col)
}

fn range_to_a1(
    start: formula_engine::eval::CellAddr,
    end: formula_engine::eval::CellAddr,
) -> String {
    let mut out = String::new();
    formula_model::push_a1_cell_range(
        start.row,
        start.col,
        end.row,
        end.col,
        false,
        false,
        &mut out,
    );
    out
}

fn main() -> Result<()> {
    let args = Args::parse();

    let corpus_bytes = fs::read_to_string(&args.cases)
        .with_context(|| format!("read cases corpus {}", args.cases.display()))?;
    let corpus: CaseCorpus =
        serde_json::from_str(&corpus_bytes).context("parse cases corpus JSON")?;

    if corpus.schema_version != 1 {
        return Err(anyhow!(
            "unsupported cases schemaVersion: {}",
            corpus.schema_version
        ));
    }

    let sha256 = sha256_file(&args.cases)?;
    let CaseCorpus {
        case_set,
        default_sheet,
        cases,
        ..
    } = corpus;

    let include: HashSet<String> = args
        .include_tags
        .iter()
        .map(|s| s.trim())
        .filter(|s| !s.is_empty())
        .map(|s| s.to_string())
        .collect();
    let exclude: HashSet<String> = args
        .exclude_tags
        .iter()
        .map(|s| s.trim())
        .filter(|s| !s.is_empty())
        .map(|s| s.to_string())
        .collect();

    let filtered_cases: Vec<Case> = cases
        .into_iter()
        .filter(|case| {
            if !include.is_empty() && !case.tags.iter().any(|t| include.contains(t)) {
                return false;
            }
            if !exclude.is_empty() && case.tags.iter().any(|t| exclude.contains(t)) {
                return false;
            }
            true
        })
        .collect();

    let max_cases = if args.max_cases > 0 {
        args.max_cases.min(filtered_cases.len())
    } else {
        filtered_cases.len()
    };

    let mut results = Vec::new();
    let _ = results.try_reserve_exact(max_cases);

    for case in filtered_cases.into_iter().take(max_cases) {
        let mut engine = Engine::new();
        let mut case_error: Option<String> = None;

        // Apply inputs first.
        for input in case.inputs.iter() {
            if let Some(formula) = &input.formula {
                if let Err(err) = engine.set_cell_formula(&default_sheet, &input.cell, formula) {
                    case_error = Some(format!("set input formula failed: {err}"));
                    break;
                }
            } else if let Some(value) = input.value.clone() {
                let v = match encode_json_value(value) {
                    Ok(v) => v,
                    Err(err) => {
                        case_error = Some(format!("encode input value failed: {err}"));
                        break;
                    }
                };
                if let Err(err) = engine.set_cell_value(&default_sheet, &input.cell, v) {
                    case_error = Some(format!("set input value failed: {err}"));
                    break;
                }
            } else {
                // No value or formula provided; treat as blank (default).
            }
        }

        if let Some(err) = case_error {
            results.push(ResultEntry {
                case_id: case.id,
                output_cell: case.output_cell.clone(),
                result: EncodedValue::engine_error(err),
                address: case.output_cell,
                display_text: "#ENGINE!".to_string(),
            });
            continue;
        }

        // Apply the formula under test.
        if let Err(err) = engine.set_cell_formula(&default_sheet, &case.output_cell, &case.formula)
        {
            results.push(ResultEntry {
                case_id: case.id,
                output_cell: case.output_cell.clone(),
                result: EncodedValue::engine_error(format!("set case formula failed: {err}")),
                address: case.output_cell,
                display_text: "#ENGINE!".to_string(),
            });
            continue;
        }

        engine.recalculate_single_threaded();
        let value = engine.get_cell_value(&default_sheet, &case.output_cell);

        let display_text = match &value {
            Value::Blank => "".to_string(),
            Value::Number(n) => {
                if n.is_finite() {
                    n.to_string()
                } else {
                    "#ENGINE!".to_string()
                }
            }
            Value::Text(s) => s.clone(),
            Value::Entity(v) => v.display.clone(),
            Value::Record(v) => v.display.clone(),
            Value::Bool(b) => {
                if *b {
                    "TRUE".to_string()
                } else {
                    "FALSE".to_string()
                }
            }
            Value::Error(e) => e.as_code().to_string(),
            Value::Lambda(_) => "#CALC!".to_string(),
            Value::Array(arr) => arr.top_left().to_string(),
            Value::Reference(_) | Value::ReferenceUnion(_) => "#VALUE!".to_string(),
            Value::Spill { .. } => "#SPILL!".to_string(),
        };

        let spill_range = engine.spill_range(&default_sheet, &case.output_cell);
        let (result, address) = match spill_range {
            Some((start, end)) => {
                let rows = (end.row - start.row + 1) as usize;
                let cols = (end.col - start.col + 1) as usize;
                let mut out_rows = Vec::new();
                let _ = out_rows.try_reserve_exact(rows);
                for r in 0..rows {
                    let mut row = Vec::new();
                    let _ = row.try_reserve_exact(cols);
                    for c in 0..cols {
                        let addr = coord_to_a1(start.row + r as u32, start.col + c as u32);
                        row.push(engine.get_cell_value(&default_sheet, &addr).into());
                    }
                    out_rows.push(row);
                }
                (
                    EncodedValue::Array { rows: out_rows },
                    range_to_a1(start, end),
                )
            }
            None => (value.clone().into(), case.output_cell.clone()),
        };

        results.push(ResultEntry {
            case_id: case.id,
            output_cell: case.output_cell.clone(),
            result,
            address,
            display_text,
        });
    }

    let generated_at = time::OffsetDateTime::now_utc()
        .format(&Rfc3339)
        .unwrap_or_else(|_| "unknown".to_string());

    let payload = ResultsFile {
        schema_version: 1,
        generated_at,
        source: SourceInfo {
            kind: "formula-engine".to_string(),
            version: env!("CARGO_PKG_VERSION").to_string(),
            os: std::env::consts::OS.to_string(),
            arch: std::env::consts::ARCH.to_string(),
            case_set,
        },
        case_set: CaseSetInfo {
            path: args.cases.display().to_string(),
            sha256,
            count: results.len(),
        },
        results,
    };

    if let Some(parent) = args.out.parent() {
        if !parent.as_os_str().is_empty() {
            fs::create_dir_all(parent)
                .with_context(|| format!("create output dir {}", parent.display()))?;
        }
    }

    let json = serde_json::to_string_pretty(&payload).context("serialize results JSON")?;
    fs::write(&args.out, format!("{json}\n"))
        .with_context(|| format!("write results {}", args.out.display()))?;

    Ok(())
}
