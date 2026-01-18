use std::borrow::Cow;
use std::collections::{BTreeMap, BTreeSet};
use std::fs;
use std::io::Cursor;
use std::io::Write;
use std::path::PathBuf;
use std::time::Instant;

use anyhow::{Context, Result};
use clap::{Parser, ValueEnum};
use formula_engine::{Engine, ErrorKind, NameDefinition, NameScope, Value as EngineValue};
use formula_model::{CellRef, CellValue, DefinedNameScope, ErrorValue};
use formula_xlsb::XlsbWorkbook;
use globset::{Glob, GlobSet, GlobSetBuilder};
use serde::Serialize;
use sha2::{Digest, Sha256};

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default, ValueEnum, Serialize)]
#[serde(rename_all = "snake_case")]
enum RoundTripFailOn {
    #[default]
    Critical,
    Warning,
    Info,
}

impl RoundTripFailOn {
    fn round_trip_ok(self, counts: &DiffCounts) -> bool {
        match self {
            RoundTripFailOn::Critical => counts.critical == 0,
            RoundTripFailOn::Warning => counts.critical == 0 && counts.warning == 0,
            RoundTripFailOn::Info => counts.total == 0,
        }
    }
}

#[derive(Parser, Debug)]
#[command(about = "Compatibility triage helper used by tools/corpus/triage.py")]
struct Args {
    /// Input workbook (XLSX/XLSM/XLSB).
    #[arg(long)]
    input: PathBuf,

    /// Workbook container format.
    ///
    /// Defaults to `auto` (detect from file extension, with a best-effort fallback based on
    /// package contents).
    #[arg(long, value_enum, default_value_t = WorkbookFormat::Auto)]
    format: WorkbookFormat,

    /// Optional password for Office-encrypted workbooks.
    ///
    /// When unset and `--skip-encrypted` is enabled (default), encrypted XLSX/XLSM/XLSB inputs are
    /// reported as `skipped` instead of `failed`.
    #[arg(long, conflicts_with = "password_file")]
    password: Option<String>,

    /// Read workbook password from a file (first line).
    #[arg(long, value_name = "PATH", conflicts_with = "password")]
    password_file: Option<PathBuf>,

    /// Skip encrypted workbooks when no password is provided.
    #[arg(long, default_value_t = true, action = clap::ArgAction::Set)]
    skip_encrypted: bool,

    /// Parts to ignore when diffing round-tripped output.
    #[arg(long = "ignore-part")]
    ignore_parts: Vec<String>,

    /// Glob patterns to ignore when diffing round-tripped output (repeatable).
    #[arg(long = "ignore-glob")]
    ignore_globs: Vec<String>,

    /// Substring patterns to ignore within XML diff paths (repeatable).
    ///
    /// This is useful for suppressing known-noisy attributes (e.g. `dyDescent`,
    /// `xr:uid`) without ignoring the entire part.
    #[arg(long = "ignore-path")]
    ignore_paths: Vec<String>,

    /// Like `--ignore-path`, but scoped to parts matched by a glob.
    ///
    /// Format: `<part_glob>:<path_substring>`. Repeatable.
    #[arg(long = "ignore-path-in")]
    ignore_paths_in: Vec<String>,

    /// Like `--ignore-path`, but only applies to diffs whose kind matches the provided kind.
    ///
    /// Format: `<kind>:<path_substring>`. Repeatable.
    #[arg(long = "ignore-path-kind")]
    ignore_paths_kind: Vec<String>,

    /// Like `--ignore-path-in`, but only applies to diffs whose kind matches the provided kind.
    ///
    /// Format: `<part_glob>:<kind>:<path_substring>`. Repeatable.
    #[arg(long = "ignore-path-kind-in")]
    ignore_paths_kind_in: Vec<String>,

    /// Built-in ignore presets for diffing round-tripped output (repeatable).
    #[arg(long = "ignore-preset")]
    ignore_presets: Vec<xlsx_diff::IgnorePreset>,

    /// Treat calcChain-related diffs as CRITICAL instead of downgrading them to WARNING.
    #[arg(long = "strict-calc-chain")]
    strict_calc_chain: bool,

    /// Maximum number of differences to emit (privacy-safe summary only).
    #[arg(long, default_value_t = 25)]
    diff_limit: usize,

    /// Round-trip diff severity threshold considered a failure.
    #[arg(long = "fail-on", value_enum, default_value_t = RoundTripFailOn::Critical)]
    fail_on: RoundTripFailOn,

    /// Run a best-effort recalculation check against cached workbook values.
    #[arg(long)]
    recalc: bool,

    /// Run a lightweight headless render/pagination smoke test.
    #[arg(long = "render-smoke")]
    render_smoke: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, ValueEnum)]
enum WorkbookFormat {
    Auto,
    Xlsx,
    Xlsb,
}

impl WorkbookFormat {
    fn resolve(self, input: &PathBuf, input_bytes: &[u8]) -> WorkbookFormat {
        match self {
            WorkbookFormat::Auto => detect_workbook_format(input, input_bytes),
            other => other,
        }
    }
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
    round_trip_fail_on: RoundTripFailOn,
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
struct DiffPartStats {
    parts_total: usize,
    parts_changed: usize,
    parts_changed_critical: usize,
}

#[derive(Debug, Serialize)]
struct DiffEntry {
    fingerprint: String,
    severity: String,
    part: String,
    path: String,
    kind: String,
}

#[derive(Debug, Serialize, Clone, PartialEq, Eq)]
struct PartDiffSummary {
    part: String,
    group: String,
    critical: usize,
    warning: usize,
    info: usize,
    total: usize,
}

#[derive(Debug, Serialize)]
struct DiffDetails {
    ignore: Vec<String>,
    ignore_globs: Vec<String>,
    ignore_paths: Vec<String>,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    ignore_presets: Vec<String>,
    strict_calc_chain: bool,
    counts: DiffCounts,
    part_stats: DiffPartStats,
    equal: bool,
    parts_with_diffs: Vec<PartDiffSummary>,
    critical_parts: Vec<String>,
    #[serde(skip_serializing_if = "BTreeMap::is_empty")]
    part_groups: BTreeMap<String, String>,
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
    validate_ignore_globs(&args.ignore_globs)?;
    let output = run(&args);
    print_json(&output)?;
    Ok(())
}

fn run(args: &Args) -> TriageOutput {
    let mut output = TriageOutput {
        steps: BTreeMap::new(),
        result: TriageResult::default(),
    };
    output.result.round_trip_fail_on = args.fail_on;

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
                return output;
            }
        }
    };

    // Decrypt Office-encrypted workbooks when possible. If no password was provided, treat them
    // as `skipped` to avoid polluting corpus stats with generic open failures.
    let decrypt_start = Instant::now();
    let (workbook_bytes, skipped_reason, encrypted) = match prepare_workbook_bytes(&input_bytes, args)
    {
        Ok(res) => res,
        Err(err) => {
            output.steps.insert(
                "load".to_string(),
                StepResult::failed(decrypt_start, err),
            );
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
            return output;
        }
    };

    if let Some(reason) = skipped_reason {
        output.result.open_ok = false;
        if encrypted {
            // For encrypted workbooks we want a stable, explicit skip reason for every step so
            // corpus triage doesn't record a generic load failure hash.
            for step in ["load", "round_trip", "diff", "recalc", "render"] {
                output
                    .steps
                    .insert(step.to_string(), StepResult::skipped(reason.clone()));
            }
        } else {
            output
                .steps
                .insert("load".to_string(), StepResult::skipped(reason));
            for step in ["recalc", "render", "round_trip", "diff"] {
                output
                    .steps
                    .insert(step.to_string(), StepResult::skipped("open_skipped"));
            }
        }
        return output;
    }

    let mut format = args.format.resolve(&args.input, &input_bytes);
    if encrypted {
        // Encrypted workbooks wrap a ZIP-based OOXML payload in an OLE container. Once decrypted,
        // detect whether the payload is XLSX/XLSM (`xl/workbook.xml`) or XLSB (`xl/workbook.bin`) so
        // we can route to the correct triage implementation.
        format = detect_workbook_format_from_zip_payload(workbook_bytes.as_ref());
    }

    match format {
        WorkbookFormat::Xlsx => {
            // Step: load (formula-xlsx)
            let start = Instant::now();
            let doc = match formula_xlsx::load_from_bytes(workbook_bytes.as_ref()) {
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
                    return output;
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
                    return output;
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
            let diff_details = match diff_workbooks(workbook_bytes.as_ref(), &round_tripped, args) {
                Ok(details) => details,
                Err(err) => {
                    output
                        .steps
                        .insert("diff".to_string(), StepResult::failed(start, err));
                    output.result.round_trip_ok = false;
                    return output;
                }
            };

            output.result.diff_critical_count = diff_details.counts.critical;
            output.result.diff_warning_count = diff_details.counts.warning;
            output.result.diff_info_count = diff_details.counts.info;
            output.result.diff_total_count = diff_details.counts.total;
            output.result.round_trip_ok = args.fail_on.round_trip_ok(&diff_details.counts);

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

            output
        }
        WorkbookFormat::Xlsb => {
            let input_parts = xlsx_diff::WorkbookArchive::from_bytes(workbook_bytes.as_ref())
                .map(|a| a.part_names().len())
                .unwrap_or(0);

            // Step: load (formula-xlsb)
            let start = Instant::now();
            let wb = match XlsbWorkbook::open_from_bytes_with_options(
                workbook_bytes.as_ref(),
                formula_xlsb::OpenOptions::default(),
            ) {
                Ok(wb) => wb,
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
                    return output;
                }
            };

            output.steps.insert(
                "load".to_string(),
                StepResult::ok(
                    start,
                    LoadDetails {
                        engine: "formula-xlsb",
                        parts: input_parts,
                        sheets: wb.sheet_metas().len(),
                    },
                ),
            );
            output.result.open_ok = true;

            // Step: round-trip save (formula-xlsb)
            let start = Instant::now();
            let mut round_tripped_writer = Cursor::new(Vec::new());
            if let Err(err) = wb.save_as_to_writer(&mut round_tripped_writer) {
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
                return output;
            }
            let round_tripped = round_tripped_writer.into_inner();

            let output_parts = xlsx_diff::WorkbookArchive::from_bytes(&round_tripped)
                .map(|a| a.part_names().len())
                .unwrap_or(0);

            output.steps.insert(
                "round_trip".to_string(),
                StepResult::ok(
                    start,
                    RoundTripDetails {
                        engine: "formula-xlsb",
                        output_size_bytes: round_tripped.len(),
                        output_parts,
                    },
                ),
            );

            // Step: diff (xlsx-diff)
            let start = Instant::now();
            let diff_details = match diff_workbooks(workbook_bytes.as_ref(), &round_tripped, args) {
                Ok(details) => details,
                Err(err) => {
                    output
                        .steps
                        .insert("diff".to_string(), StepResult::failed(start, err));
                    output.result.round_trip_ok = false;
                    return output;
                }
            };

            output.result.diff_critical_count = diff_details.counts.critical;
            output.result.diff_warning_count = diff_details.counts.warning;
            output.result.diff_info_count = diff_details.counts.info;
            output.result.diff_total_count = diff_details.counts.total;
            output.result.round_trip_ok = args.fail_on.round_trip_ok(&diff_details.counts);

            output
                .steps
                .insert("diff".to_string(), StepResult::ok(start, &diff_details));

            // Recalc/render are currently implemented against the XLSX data model only.
            if !args.recalc {
                output.steps.insert(
                    "recalc".to_string(),
                    StepResult::skipped("disabled (pass --recalc)"),
                );
            } else {
                output.steps.insert(
                    "recalc".to_string(),
                    StepResult::skipped("unsupported_for_xlsb"),
                );
            }
            output.result.calculate_ok = None;

            if !args.render_smoke {
                output.steps.insert(
                    "render".to_string(),
                    StepResult::skipped("disabled (pass --render-smoke)"),
                );
            } else {
                output.steps.insert(
                    "render".to_string(),
                    StepResult::skipped("unsupported_for_xlsb"),
                );
            }
            output.result.render_ok = None;

            output
        }
        WorkbookFormat::Auto => {
            debug_assert!(false, "WorkbookFormat::Auto should be resolved before triage");
            let start = Instant::now();
            output.steps.insert(
                "load".to_string(),
                StepResult::failed(start, "internal error: workbook format unresolved"),
            );
            output.result.open_ok = false;
            for step in ["recalc", "render", "round_trip", "diff"] {
                output
                    .steps
                    .insert(step.to_string(), StepResult::skipped("open_failed"));
            }
            output
        }
    }
}

fn resolve_password(args: &Args) -> Result<Option<String>> {
    if let Some(pw) = args.password.as_ref() {
        return Ok(Some(pw.clone()));
    }
    let Some(path) = args.password_file.as_ref() else {
        return Ok(None);
    };

    let pw = fs::read_to_string(path).with_context(|| format!("read password file {}", path.display()))?;
    let pw = pw.lines().next().unwrap_or("").trim_end_matches(&['\r', '\n'][..]).to_string();
    Ok(Some(pw))
}

fn prepare_workbook_bytes<'a>(
    input_bytes: &'a [u8],
    args: &Args,
) -> Result<(Cow<'a, [u8]>, Option<String>, bool)> {
    if !formula_office_crypto::is_encrypted_ooxml_ole(input_bytes) {
        return Ok((Cow::Borrowed(input_bytes), None, false));
    }

    let password = resolve_password(args)?;
    let Some(password) = password.as_deref() else {
        if args.skip_encrypted {
            return Ok((Cow::Borrowed(input_bytes), Some("encrypted".to_string()), true));
        }
        return Ok((Cow::Borrowed(input_bytes), None, true));
    };

    let decrypted = formula_office_crypto::decrypt_encrypted_package(input_bytes, password)
        .context("decrypt encrypted workbook")?;
    if !looks_like_ooxml_workbook_zip(&decrypted) {
        return Ok((Cow::Borrowed(input_bytes), Some("encrypted".to_string()), true));
    }

    Ok((Cow::Owned(decrypted), None, true))
}

fn looks_like_ooxml_workbook_zip(bytes: &[u8]) -> bool {
    if bytes.len() < 2 || &bytes[..2] != b"PK" {
        return false;
    }
    let cursor = Cursor::new(bytes);
    let Ok(mut archive) = zip::ZipArchive::new(cursor) else {
        return false;
    };
    for i in 0..archive.len() {
        let Ok(file) = archive.by_index(i) else {
            continue;
        };
        if file.is_dir() {
            continue;
        }
        let name = file.name().trim_start_matches(|c| c == '/' || c == '\\');
        let name = if name.contains('\\') {
            Cow::Owned(name.replace('\\', "/"))
        } else {
            Cow::Borrowed(name)
        };
        if name.as_ref().eq_ignore_ascii_case("xl/workbook.xml")
            || name.as_ref().eq_ignore_ascii_case("xl/workbook.bin")
        {
            return true;
        }
    }
    false
}

fn detect_workbook_format_from_zip_payload(bytes: &[u8]) -> WorkbookFormat {
    if bytes.len() < 2 || &bytes[..2] != b"PK" {
        return WorkbookFormat::Xlsx;
    }
    let Ok(archive) = xlsx_diff::WorkbookArchive::from_bytes(bytes) else {
        return WorkbookFormat::Xlsx;
    };
    for part in archive.part_names() {
        if part.eq_ignore_ascii_case("xl/workbook.bin") {
            return WorkbookFormat::Xlsb;
        }
        if part.eq_ignore_ascii_case("xl/workbook.xml") {
            // Keep scanning in case the package contains both `.xml` and `.bin` (prefer `.bin`).
        }
    }
    WorkbookFormat::Xlsx
}
fn print_json(output: &TriageOutput) -> Result<()> {
    let json = serde_json::to_string(output).context("serialize triage output")?;
    let stdout = std::io::stdout();
    let mut out = std::io::BufWriter::new(stdout.lock());
    if let Err(err) = writeln!(&mut out, "{json}") {
        if err.kind() == std::io::ErrorKind::BrokenPipe {
            // Allow piping output to tools like `head` without panicking.
            return Ok(());
        }
        return Err(err.into());
    }
    Ok(())
}

fn detect_workbook_format(input: &PathBuf, input_bytes: &[u8]) -> WorkbookFormat {
    if let Some(ext) = input.extension().and_then(|s| s.to_str()) {
        if ext.eq_ignore_ascii_case("xlsb") {
            return WorkbookFormat::Xlsb;
        }
        if ext.eq_ignore_ascii_case("xlsx") || ext.eq_ignore_ascii_case("xlsm") {
            return WorkbookFormat::Xlsx;
        }
    }

    // Best-effort fallback: inspect the package for the XLSB workbook part.
    if let Ok(archive) = xlsx_diff::WorkbookArchive::from_bytes(input_bytes) {
        let mut has_workbook_bin = false;
        let mut has_workbook_xml = false;
        for part in archive.part_names() {
            if part.eq_ignore_ascii_case("xl/workbook.bin") {
                has_workbook_bin = true;
            }
            if part.eq_ignore_ascii_case("xl/workbook.xml") {
                has_workbook_xml = true;
            }
        }
        if has_workbook_bin && !has_workbook_xml {
            return WorkbookFormat::Xlsb;
        }
        if has_workbook_bin {
            return WorkbookFormat::Xlsb;
        }
    }

    WorkbookFormat::Xlsx
}

fn build_ignore_path_rules(args: &Args) -> Result<Vec<xlsx_diff::IgnorePathRule>> {
    let mut rules = Vec::new();

    for pattern in &args.ignore_paths {
        let trimmed = pattern.trim();
        if trimmed.is_empty() {
            continue;
        }
        rules.push(xlsx_diff::IgnorePathRule {
            part: None,
            path_substring: if trimmed.contains('\\') {
                trimmed.replace('\\', "/")
            } else {
                trimmed.to_string()
            },
            kind: None,
        });
    }

    for scoped in &args.ignore_paths_in {
        let trimmed = scoped.trim();
        if trimmed.is_empty() {
            continue;
        }
        let Some((part_glob, substring)) = trimmed.split_once(':') else {
            anyhow::bail!(
                "invalid --ignore-path-in '{trimmed}' (expected format: <part_glob>:<path_substring>)"
            );
        };
        let part_glob = normalize_ignore_pattern(part_glob);
        let substring = substring.trim();
        if part_glob.is_empty() || substring.is_empty() {
            anyhow::bail!(
                "invalid --ignore-path-in '{trimmed}' (expected non-empty <part_glob> and <path_substring>)"
            );
        }

        // Validate the (normalized) glob syntax so callers get deterministic errors. Patterns
        // without `*`/`?` are treated as exact matches (so callers can target `[Content_Types].xml`
        // without needing to escape `[`/`]` for glob syntax).
        if part_glob.contains('*') || part_glob.contains('?') {
            Glob::new(&part_glob)?;
        }

        rules.push(xlsx_diff::IgnorePathRule {
            part: Some(part_glob),
            path_substring: if substring.contains('\\') {
                substring.replace('\\', "/")
            } else {
                substring.to_string()
            },
            kind: None,
        });
    }

    for spec in &args.ignore_paths_kind {
        let trimmed = spec.trim();
        if trimmed.is_empty() {
            continue;
        }
        let Some((kind, substring)) = trimmed.split_once(':') else {
            anyhow::bail!(
                "invalid --ignore-path-kind '{trimmed}' (expected format: <kind>:<path_substring>)"
            );
        };
        let kind = kind.trim();
        let substring = substring.trim();
        if kind.is_empty() || substring.is_empty() {
            anyhow::bail!(
                "invalid --ignore-path-kind '{trimmed}' (expected non-empty <kind> and <path_substring>)"
            );
        }
        rules.push(xlsx_diff::IgnorePathRule {
            part: None,
            path_substring: if substring.contains('\\') {
                substring.replace('\\', "/")
            } else {
                substring.to_string()
            },
            kind: Some(kind.to_string()),
        });
    }

    for spec in &args.ignore_paths_kind_in {
        let trimmed = spec.trim();
        if trimmed.is_empty() {
            continue;
        }
        let mut iter = trimmed.splitn(3, ':');
        let part_glob = iter.next().unwrap_or_default();
        let kind = iter.next();
        let substring = iter.next();
        let (Some(kind), Some(substring)) = (kind, substring) else {
            anyhow::bail!(
                "invalid --ignore-path-kind-in '{trimmed}' (expected format: <part_glob>:<kind>:<path_substring>)"
            );
        };
        let part_glob = normalize_ignore_pattern(part_glob);
        let kind = kind.trim();
        let substring = substring.trim();
        if part_glob.is_empty() || kind.is_empty() || substring.is_empty() {
            anyhow::bail!(
                "invalid --ignore-path-kind-in '{trimmed}' (expected non-empty <part_glob>, <kind>, and <path_substring>)"
            );
        }

        if part_glob.contains('*') || part_glob.contains('?') {
            Glob::new(&part_glob)?;
        }

        rules.push(xlsx_diff::IgnorePathRule {
            part: Some(part_glob),
            path_substring: if substring.contains('\\') {
                substring.replace('\\', "/")
            } else {
                substring.to_string()
            },
            kind: Some(kind.to_string()),
        });
    }

    Ok(rules)
}

fn diff_workbooks(expected: &[u8], actual: &[u8], args: &Args) -> Result<DiffDetails> {
    let ignore: BTreeSet<String> = args
        .ignore_parts
        .iter()
        .map(|s| normalize_opc_part_name(s))
        .filter(|s| !s.is_empty())
        .collect();
    let ignore_sorted: Vec<String> = ignore.iter().cloned().collect();

    let ignore_globs: BTreeSet<String> = args
        .ignore_globs
        .iter()
        .map(|s| normalize_ignore_pattern(s))
        .filter(|s| !s.is_empty())
        .collect();
    let ignore_globs_sorted: Vec<String> = ignore_globs.iter().cloned().collect();
    let ignore_globs_matcher: GlobSet = {
        let mut builder = GlobSetBuilder::new();
        for pattern in &ignore_globs_sorted {
            // `xlsx-diff` ignores invalid glob patterns at runtime, but the CLI validates them
            // up-front. Mirror that behavior here so unit tests that call `diff_workbooks` directly
            // don't panic when given invalid patterns.
            if let Ok(glob) = Glob::new(pattern) {
                builder.add(glob);
            }
        }
        builder.build().unwrap_or_else(|_| GlobSet::empty())
    };

    let ignore_paths = build_ignore_path_rules(args)?;
    let mut ignore_paths_sorted: Vec<String> = ignore_paths
        .iter()
        .map(|r| match (&r.part, r.kind.as_deref()) {
            (Some(part), Some(kind)) => format!("{part}:{kind}:{}", r.path_substring),
            (Some(part), None) => format!("{part}:{}", r.path_substring),
            (None, Some(kind)) => format!("{kind}:{}", r.path_substring),
            (None, None) => r.path_substring.clone(),
        })
        .collect();
    ignore_paths_sorted.sort();
    let mut ignore_presets_sorted: Vec<String> =
        args.ignore_presets.iter().map(|p| p.to_string()).collect();
    ignore_presets_sorted.sort();
    ignore_presets_sorted.dedup();

    let expected_archive = xlsx_diff::WorkbookArchive::from_bytes(expected)?;
    let actual_archive = xlsx_diff::WorkbookArchive::from_bytes(actual)?;

    // Part-level stats are computed after applying ignore rules, mirroring the diff itself.
    let expected_parts: BTreeSet<&str> = expected_archive
        .part_names()
        .into_iter()
        .filter(|part| !ignore.contains(*part) && !ignore_globs_matcher.is_match(*part))
        .collect();
    let actual_parts: BTreeSet<&str> = actual_archive
        .part_names()
        .into_iter()
        .filter(|part| !ignore.contains(*part) && !ignore_globs_matcher.is_match(*part))
        .collect();
    let parts_total = expected_parts.union(&actual_parts).count();

    let strict_calc_chain = args.strict_calc_chain;
    let mut options = xlsx_diff::DiffOptions {
        ignore_parts: ignore,
        ignore_globs: ignore_globs_sorted.clone(),
        ignore_paths,
        strict_calc_chain,
    };
    for preset in &args.ignore_presets {
        options.apply_preset(*preset);
    }

    let report = xlsx_diff::diff_archives_with_options(&expected_archive, &actual_archive, &options);

    let (parts_with_diffs, critical_parts, part_groups) = summarize_diffs_by_part(&report);
    let parts_changed = parts_with_diffs.len();
    let parts_changed_critical = critical_parts.len();

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
        .map(|d| {
            let severity = d.severity.as_str().to_string();
            let part = d.part;
            let path = d.path;
            let kind = d.kind;
            let fingerprint = diff_entry_fingerprint(&severity, &part, &path, &kind);
            DiffEntry {
                fingerprint,
                severity,
                part,
                path,
                kind,
            }
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
        ignore_globs: ignore_globs_sorted,
        ignore_paths: ignore_paths_sorted,
        strict_calc_chain,
        ignore_presets: ignore_presets_sorted,
        counts,
        part_stats: DiffPartStats {
            parts_total,
            parts_changed,
            parts_changed_critical,
        },
        equal,
        parts_with_diffs,
        critical_parts,
        part_groups,
        top_differences: entries,
    })
}

fn summarize_diffs_by_part(
    report: &xlsx_diff::DiffReport,
) -> (Vec<PartDiffSummary>, Vec<String>, BTreeMap<String, String>) {
    #[derive(Default)]
    struct Counts {
        critical: usize,
        warning: usize,
        info: usize,
        total: usize,
    }

    let mut counts_by_part: BTreeMap<String, Counts> = BTreeMap::new();
    for diff in &report.differences {
        let counts = counts_by_part.entry(diff.part.clone()).or_default();
        counts.total += 1;
        match diff.severity {
            xlsx_diff::Severity::Critical => counts.critical += 1,
            xlsx_diff::Severity::Warning => counts.warning += 1,
            xlsx_diff::Severity::Info => counts.info += 1,
        }
    }

    let mut parts_with_diffs: Vec<PartDiffSummary> = counts_by_part
        .into_iter()
        .map(|(part, c)| {
            let group = part_group(&part).to_string();
            PartDiffSummary {
                part,
                group,
                critical: c.critical,
                warning: c.warning,
                info: c.info,
                total: c.total,
            }
        })
        .collect();

    // Stable output order: (critical desc, warning desc, info desc, part asc).
    parts_with_diffs.sort_by(|a, b| {
        (
            std::cmp::Reverse(a.critical),
            std::cmp::Reverse(a.warning),
            std::cmp::Reverse(a.info),
            a.part.as_str(),
        )
            .cmp(&(
                std::cmp::Reverse(b.critical),
                std::cmp::Reverse(b.warning),
                std::cmp::Reverse(b.info),
                b.part.as_str(),
            ))
    });

    let critical_parts: Vec<String> = parts_with_diffs
        .iter()
        .filter(|p| p.critical > 0)
        .map(|p| p.part.clone())
        .collect();

    let part_groups: BTreeMap<String, String> = parts_with_diffs
        .iter()
        .map(|p| (p.part.clone(), p.group.clone()))
        .collect();

    (parts_with_diffs, critical_parts, part_groups)
}

fn part_group(part: &str) -> &'static str {
    fn starts_with_ignore_ascii_case(s: &str, prefix: &str) -> bool {
        s.get(..prefix.len())
            .is_some_and(|p| p.eq_ignore_ascii_case(prefix))
    }

    fn ends_with_ignore_ascii_case(s: &str, suffix: &str) -> bool {
        if suffix.len() > s.len() {
            return false;
        }
        s[s.len() - suffix.len()..].eq_ignore_ascii_case(suffix)
    }

    fn contains_ignore_ascii_case(s: &str, needle: &str) -> bool {
        let s = s.as_bytes();
        let needle = needle.as_bytes();
        if needle.is_empty() {
            return true;
        }
        if needle.len() > s.len() {
            return false;
        }
        for start in 0..=s.len() - needle.len() {
            if s[start..start + needle.len()].eq_ignore_ascii_case(needle) {
                return true;
            }
        }
        false
    }

    if part.eq_ignore_ascii_case("[content_types].xml") {
        return "content_types";
    }
    if ends_with_ignore_ascii_case(part, ".rels") || contains_ignore_ascii_case(part, "/_rels/") {
        return "rels";
    }
    if part.eq_ignore_ascii_case("xl/workbook.xml") || part.eq_ignore_ascii_case("xl/workbook.bin")
    {
        return "workbook";
    }
    if part.eq_ignore_ascii_case("xl/styles.xml")
        || part.eq_ignore_ascii_case("xl/styles.bin")
        || starts_with_ignore_ascii_case(part, "xl/styles/")
    {
        return "styles";
    }
    if part.eq_ignore_ascii_case("xl/sharedstrings.xml")
        || part.eq_ignore_ascii_case("xl/sharedstrings.bin")
        || ends_with_ignore_ascii_case(part, "/sharedstrings.xml")
        || ends_with_ignore_ascii_case(part, "/sharedstrings.bin")
    {
        return "shared_strings";
    }
    if starts_with_ignore_ascii_case(part, "xl/worksheets/") {
        if ends_with_ignore_ascii_case(part, ".bin") {
            return "worksheet_bin";
        }
        return "worksheet_xml";
    }
    if part.eq_ignore_ascii_case("xl/calcchain.xml") || part.eq_ignore_ascii_case("xl/calcchain.bin")
    {
        return "calc_chain";
    }
    if starts_with_ignore_ascii_case(part, "xl/theme/") {
        return "theme";
    }
    if starts_with_ignore_ascii_case(part, "xl/pivottables/")
        || starts_with_ignore_ascii_case(part, "xl/pivotcache/")
    {
        return "pivots";
    }
    if starts_with_ignore_ascii_case(part, "xl/charts/") {
        return "charts";
    }
    if starts_with_ignore_ascii_case(part, "xl/drawings/") {
        return "drawings";
    }
    if starts_with_ignore_ascii_case(part, "xl/tables/") {
        return "tables";
    }
    if starts_with_ignore_ascii_case(part, "xl/externallinks/") {
        return "external_links";
    }
    if starts_with_ignore_ascii_case(part, "xl/media/") {
        return "media";
    }
    if starts_with_ignore_ascii_case(part, "docprops/") {
        return "doc_props";
    }
    "other"
}

fn normalize_opc_part_name(part: &str) -> String {
    normalize_opc_path(part.trim().trim_start_matches('/'))
}

fn normalize_opc_path(path: &str) -> String {
    // Keep in sync with `xlsx-diff`'s internal normalization logic so ignore rules and part-level
    // stats remain stable and match the diff report semantics.
    let normalized: Cow<'_, str> = if path.contains('\\') {
        Cow::Owned(path.replace('\\', "/"))
    } else {
        Cow::Borrowed(path)
    };
    let mut out: Vec<&str> = Vec::new();
    for segment in normalized.split('/') {
        match segment {
            "" | "." => {}
            ".." => {
                out.pop();
            }
            _ => out.push(segment),
        }
    }
    out.join("/")
}

fn normalize_ignore_pattern(input: &str) -> String {
    let trimmed = input.trim();
    let trimmed = trimmed.trim_start_matches(|c| c == '/' || c == '\\');
    if trimmed.contains('\\') {
        trimmed.replace('\\', "/")
    } else {
        trimmed.to_string()
    }
}

fn validate_ignore_globs(ignore_globs: &[String]) -> Result<()> {
    for raw in ignore_globs {
        let pattern = normalize_ignore_pattern(raw);
        if pattern.is_empty() {
            continue;
        }
        if let Err(err) = Glob::new(&pattern) {
            // Avoid leaking user-provided patterns into logs/artifacts; follow the same hashing
            // convention as other triage errors.
            let sha = sha256_text(&err.to_string());
            anyhow::bail!("invalid --ignore-glob pattern (sha256={sha})");
        }
    }
    Ok(())
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
            let a1 = formula_model::cell_to_a1(cell_ref.row, cell_ref.col);
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
            let addr = formula_model::cell_to_a1(cell_ref.row, cell_ref.col);
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

fn diff_entry_fingerprint(severity: &str, part: &str, path: &str, kind: &str) -> String {
    // Privacy-safe stable fingerprint: aggregate diffs across workbooks without including any
    // expected/actual values (which may contain workbook data).
    sha256_text(&format!("{severity}|{part}|{path}|{kind}"))
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
    use base64::Engine as _;
    use std::io::{Cursor, Write};
    use zip::write::FileOptions;
    use zip::{CompressionMethod, ZipWriter};

    fn plain_xlsx_bytes() -> Vec<u8> {
        let b64_path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../tools/corpus/public/simple.xlsx.b64");
        let b64 = std::fs::read_to_string(&b64_path)
            .unwrap_or_else(|e| panic!("read {b64_path:?}: {e}"));
        base64::engine::general_purpose::STANDARD
            .decode(b64.trim())
            .expect("decode base64 xlsx fixture")
    }

    fn plain_xlsb_bytes() -> Vec<u8> {
        let b64_path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../tools/corpus/public/simple.xlsb.b64");
        let b64 = std::fs::read_to_string(&b64_path)
            .unwrap_or_else(|e| panic!("read {b64_path:?}: {e}"));
        base64::engine::general_purpose::STANDARD
            .decode(b64.trim())
            .expect("decode base64 xlsb fixture")
    }

    fn encrypt_ooxml_agile(plaintext_zip: &[u8], password: &str) -> Vec<u8> {
        let opts = formula_office_crypto::EncryptOptions {
            // Keep the fixture cheap: 1000 iterations is plenty to exercise the decrypt path.
            spin_count: 1_000,
            ..Default::default()
        };
        formula_office_crypto::encrypt_package_to_ole(plaintext_zip, password, opts)
            .expect("encrypt workbook")
    }

    #[test]
    fn preserves_extended_errors_in_normalization_and_engine_seeding() {
        let mut engine = Engine::new();
        engine.ensure_sheet("Sheet1");

        for (addr, value, expected_code) in [
            ("A1", CellValue::Error(ErrorValue::Field), "#FIELD!"),
            (
                "A2",
                CellValue::Error(ErrorValue::GettingData),
                "#GETTING_DATA",
            ),
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

    fn make_zip(parts: &[(&str, &str)]) -> Vec<u8> {
        let cursor = Cursor::new(Vec::new());
        let mut zip = ZipWriter::new(cursor);
        let opts = FileOptions::<()>::default().compression_method(CompressionMethod::Stored);

        for (name, content) in parts {
            zip.start_file(*name, opts).unwrap();
            zip.write_all(content.as_bytes()).unwrap();
        }

        zip.finish().unwrap().into_inner()
    }

    #[test]
    fn diff_part_stats_counts_parts_after_ignore_and_by_severity() {
        let expected = make_zip(&[
            ("xl/workbook.xml", "<workbook foo=\"1\"/>"),
            ("docProps/app.xml", "<app foo=\"1\"/>"),
            ("xl/theme/theme1.xml", "<theme foo=\"1\"/>"),
            ("xl/styles.xml", "<styles foo=\"1\"/>"),
        ]);

        let actual = make_zip(&[
            ("xl/workbook.xml", "<workbook foo=\"2\"/>"), // critical diff
            ("docProps/app.xml", "<app foo=\"2\"/>"),     // ignored diff
            ("xl/theme/theme1.xml", "<theme foo=\"1\"/>"), // unchanged
            ("docProps/core.xml", "<core foo=\"1\"/>"),   // extra part (info)
        ]);

        let args = Args {
            input: PathBuf::from("dummy.xlsx"),
            format: WorkbookFormat::Xlsx,
            password: None,
            password_file: None,
            skip_encrypted: true,
            ignore_parts: vec!["docProps/app.xml".to_string()],
            ignore_globs: Vec::new(),
            ignore_paths: Vec::new(),
            ignore_paths_in: Vec::new(),
            ignore_paths_kind: Vec::new(),
            ignore_paths_kind_in: Vec::new(),
            ignore_presets: Vec::new(),
            strict_calc_chain: false,
            diff_limit: 100,
            fail_on: RoundTripFailOn::Critical,
            recalc: false,
            render_smoke: false,
        };

        let details = diff_workbooks(&expected, &actual, &args).unwrap();
        assert_eq!(details.part_stats.parts_total, 4);
        assert_eq!(details.part_stats.parts_changed, 3);
        assert_eq!(details.part_stats.parts_changed_critical, 2);
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

    #[test]
    fn diff_part_breakdown_is_populated_and_sorted() {
        use xlsx_diff::{DiffReport, Difference, Severity};

        let report = DiffReport {
            differences: vec![
                Difference::new(
                    Severity::Info,
                    "docProps/app.xml",
                    "",
                    "binary_diff",
                    None,
                    None,
                ),
                Difference::new(
                    Severity::Info,
                    "docProps/core.xml",
                    "",
                    "binary_diff",
                    None,
                    None,
                ),
                Difference::new(
                    Severity::Warning,
                    "xl/theme/theme1.xml",
                    "",
                    "binary_diff",
                    None,
                    None,
                ),
                Difference::new(
                    Severity::Warning,
                    "xl/theme/theme1.xml",
                    "",
                    "binary_diff",
                    None,
                    None,
                ),
                Difference::new(
                    Severity::Warning,
                    "xl/theme/theme1.xml",
                    "",
                    "binary_diff",
                    None,
                    None,
                ),
                Difference::new(
                    Severity::Critical,
                    "xl/workbook.xml",
                    "",
                    "missing_part",
                    None,
                    None,
                ),
                Difference::new(
                    Severity::Critical,
                    "xl/workbook.xml",
                    "",
                    "extra_part",
                    None,
                    None,
                ),
            ],
        };

        let (parts, critical_parts, part_groups) = summarize_diffs_by_part(&report);

        assert_eq!(
            parts.iter().map(|p| p.part.as_str()).collect::<Vec<_>>(),
            vec![
                "xl/workbook.xml",
                "xl/theme/theme1.xml",
                "docProps/app.xml",
                "docProps/core.xml",
            ],
            "parts should be sorted by severity counts (critical > warning > info) then part name"
        );

        assert_eq!(
            parts[0],
            PartDiffSummary {
                part: "xl/workbook.xml".to_string(),
                group: "workbook".to_string(),
                critical: 2,
                warning: 0,
                info: 0,
                total: 2,
            }
        );
        assert_eq!(parts[1].warning, 3);
        assert_eq!(parts[2].info, 1);
        assert_eq!(parts[3].info, 1);

        assert_eq!(critical_parts, vec!["xl/workbook.xml".to_string()]);

        assert_eq!(part_groups.get("xl/workbook.xml").map(String::as_str), Some("workbook"));
        assert_eq!(
            part_groups.get("docProps/app.xml").map(String::as_str),
            Some("doc_props")
        );

        assert_eq!(parts[2].group, "doc_props");
    }

    #[test]
    fn diff_entry_fingerprint_is_stable() {
        let fp = diff_entry_fingerprint(
            "CRITICAL",
            "xl/workbook.xml.rels",
            "/Relationships/Relationship[1]/@Id",
            "attribute_changed",
        );
        assert_eq!(
            fp,
            "37a012601a0da63445b4fbe412c6c753406e776b652013e0ca21a56a36fb634e"
        );
    }

    #[test]
    fn diff_workbooks_emits_fingerprints_for_entries() {
        fn zip_with_part(name: &str, bytes: &[u8]) -> Vec<u8> {
            let cursor = std::io::Cursor::new(Vec::new());
            let mut zip = ZipWriter::new(cursor);
            let options =
                FileOptions::<()>::default().compression_method(CompressionMethod::Stored);
            zip.start_file(name, options).unwrap();
            zip.write_all(bytes).unwrap();
            let cursor = zip.finish().unwrap();
            cursor.into_inner()
        }

        let expected_xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="t" Target="worksheets/sheet1.xml"/>
</Relationships>"#;
        let actual_xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="t" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

        let expected = zip_with_part("xl/_rels/workbook.xml.rels", expected_xml);
        let actual = zip_with_part("xl/_rels/workbook.xml.rels", actual_xml);

        let args = Args {
            input: PathBuf::new(),
            format: WorkbookFormat::Xlsx,
            password: None,
            password_file: None,
            skip_encrypted: true,
            ignore_parts: Vec::new(),
            ignore_globs: Vec::new(),
            ignore_paths: Vec::new(),
            ignore_paths_in: Vec::new(),
            ignore_paths_kind: Vec::new(),
            ignore_paths_kind_in: Vec::new(),
            ignore_presets: Vec::new(),
            strict_calc_chain: false,
            diff_limit: 10,
            fail_on: RoundTripFailOn::Critical,
            recalc: false,
            render_smoke: false,
        };

        let details = diff_workbooks(&expected, &actual, &args).unwrap();
        assert!(
            !details.top_differences.is_empty(),
            "expected at least one diff entry"
        );

        let entry = &details.top_differences[0];
        assert_eq!(
            entry.fingerprint,
            diff_entry_fingerprint(&entry.severity, &entry.part, &entry.path, &entry.kind)
        );
        assert_eq!(entry.fingerprint.len(), 64);
    }

    #[test]
    fn round_trip_fail_on_thresholds_match_expected_semantics() {
        let warning_only = DiffCounts {
            critical: 0,
            warning: 1,
            info: 0,
            total: 1,
        };
        assert!(RoundTripFailOn::Critical.round_trip_ok(&warning_only));
        assert!(!RoundTripFailOn::Warning.round_trip_ok(&warning_only));
        assert!(!RoundTripFailOn::Info.round_trip_ok(&warning_only));

        let info_only = DiffCounts {
            critical: 0,
            warning: 0,
            info: 1,
            total: 1,
        };
        assert!(RoundTripFailOn::Critical.round_trip_ok(&info_only));
        assert!(RoundTripFailOn::Warning.round_trip_ok(&info_only));
        assert!(!RoundTripFailOn::Info.round_trip_ok(&info_only));

        let critical_present = DiffCounts {
            critical: 1,
            warning: 0,
            info: 0,
            total: 1,
        };
        assert!(!RoundTripFailOn::Critical.round_trip_ok(&critical_present));
        assert!(!RoundTripFailOn::Warning.round_trip_ok(&critical_present));
        assert!(!RoundTripFailOn::Info.round_trip_ok(&critical_present));

        let no_diffs = DiffCounts {
            critical: 0,
            warning: 0,
            info: 0,
            total: 0,
        };
        assert!(RoundTripFailOn::Critical.round_trip_ok(&no_diffs));
        assert!(RoundTripFailOn::Warning.round_trip_ok(&no_diffs));
        assert!(RoundTripFailOn::Info.round_trip_ok(&no_diffs));
    }

    #[test]
    fn diff_workbooks_respects_ignore_path_rules() {
        fn zip_with_part(name: &str, bytes: &[u8]) -> Vec<u8> {
            let cursor = std::io::Cursor::new(Vec::new());
            let mut zip = ZipWriter::new(cursor);
            let options =
                FileOptions::<()>::default().compression_method(CompressionMethod::Stored);
            zip.start_file(name, options).unwrap();
            zip.write_all(bytes).unwrap();
            let cursor = zip.finish().unwrap();
            cursor.into_inner()
        }

        let expected_xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" dyDescent="1"/>"#;
        let actual_xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" dyDescent="2"/>"#;

        let expected = zip_with_part("xl/workbook.xml", expected_xml);
        let actual = zip_with_part("xl/workbook.xml", actual_xml);

        let args = Args {
            input: PathBuf::new(),
            format: WorkbookFormat::Xlsx,
            password: None,
            password_file: None,
            skip_encrypted: true,
            ignore_parts: Vec::new(),
            ignore_globs: Vec::new(),
            ignore_paths: Vec::new(),
            ignore_paths_in: Vec::new(),
            ignore_paths_kind: Vec::new(),
            ignore_paths_kind_in: Vec::new(),
            ignore_presets: Vec::new(),
            strict_calc_chain: false,
            diff_limit: 10,
            fail_on: RoundTripFailOn::Critical,
            recalc: false,
            render_smoke: false,
        };

        let details = diff_workbooks(&expected, &actual, &args).unwrap();
        assert!(
            details.counts.total > 0,
            "expected at least one diff without ignore-path"
        );

        let args = Args {
            ignore_paths: vec!["dyDescent".to_string()],
            ..args
        };

        let details = diff_workbooks(&expected, &actual, &args).unwrap();
        assert_eq!(
            details.counts.total, 0,
            "diffs should be suppressed by ignore-path"
        );
        assert!(details.equal);
        assert_eq!(details.ignore_paths, vec!["dyDescent".to_string()]);
    }

    #[test]
    fn diff_workbooks_respects_ignore_path_in_rules() {
        let expected = make_zip(&[
            (
                "xl/workbook.xml",
                r#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" dyDescent="1"/>"#,
            ),
            (
                "xl/worksheets/sheet1.xml",
                r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" dyDescent="1"/>"#,
            ),
        ]);
        let actual = make_zip(&[
            (
                "xl/workbook.xml",
                r#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" dyDescent="2"/>"#,
            ),
            (
                "xl/worksheets/sheet1.xml",
                r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" dyDescent="2"/>"#,
            ),
        ]);

        let args = Args {
            input: PathBuf::new(),
            format: WorkbookFormat::Xlsx,
            password: None,
            password_file: None,
            skip_encrypted: true,
            ignore_parts: Vec::new(),
            ignore_globs: Vec::new(),
            ignore_paths: Vec::new(),
            ignore_paths_in: Vec::new(),
            ignore_paths_kind: Vec::new(),
            ignore_paths_kind_in: Vec::new(),
            ignore_presets: Vec::new(),
            strict_calc_chain: false,
            diff_limit: 10,
            fail_on: RoundTripFailOn::Critical,
            recalc: false,
            render_smoke: false,
        };

        let details = diff_workbooks(&expected, &actual, &args).unwrap();
        assert!(
            details.counts.total > 0,
            "expected at least one diff without ignore-path-in"
        );

        let args = Args {
            ignore_paths_in: vec!["xl/workbook.xml:dyDescent".to_string()],
            ..args
        };

        let details = diff_workbooks(&expected, &actual, &args).unwrap();
        assert!(
            details
                .top_differences
                .iter()
                .all(|d| d.part != "xl/workbook.xml"),
            "expected workbook.xml diffs to be suppressed, got {:#?}",
            details.top_differences
        );
        assert!(
            details
                .top_differences
                .iter()
                .any(|d| d.part == "xl/worksheets/sheet1.xml"),
            "expected sheet1.xml diffs to remain, got {:#?}",
            details.top_differences
        );
        assert_eq!(
            details.ignore_paths,
            vec!["xl/workbook.xml:dyDescent".to_string()]
        );
    }

    #[test]
    fn diff_workbooks_respects_ignore_path_kind_rules() {
        let expected = make_zip(&[(
            "xl/worksheets/sheet1.xml",
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
  <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
</worksheet>"#,
        )]);
        let actual = make_zip(&[(
            "xl/worksheets/sheet1.xml",
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
  <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.30" foo="bar"/>
</worksheet>"#,
        )]);

        let args = Args {
            input: PathBuf::new(),
            format: WorkbookFormat::Xlsx,
            password: None,
            password_file: None,
            skip_encrypted: true,
            ignore_parts: Vec::new(),
            ignore_globs: Vec::new(),
            ignore_paths: Vec::new(),
            ignore_paths_in: Vec::new(),
            ignore_paths_kind: Vec::new(),
            ignore_paths_kind_in: Vec::new(),
            ignore_presets: Vec::new(),
            strict_calc_chain: false,
            diff_limit: 10,
            fail_on: RoundTripFailOn::Critical,
            recalc: false,
            render_smoke: false,
        };

        let details = diff_workbooks(&expected, &actual, &args).unwrap();
        assert!(
            details
                .top_differences
                .iter()
                .any(|d| d.kind == "attribute_changed"),
            "expected a dyDescent attribute_changed diff, got {:#?}",
            details.top_differences
        );
        assert!(
            details
                .top_differences
                .iter()
                .any(|d| d.kind == "attribute_added"),
            "expected a foo attribute_added diff, got {:#?}",
            details.top_differences
        );

        let args = Args {
            ignore_paths_kind: vec!["attribute_changed:dyDescent".to_string()],
            ..args
        };
        let details = diff_workbooks(&expected, &actual, &args).unwrap();
        assert_eq!(details.counts.total, 1);
        assert_eq!(details.top_differences.len(), 1);
        let diff = &details.top_differences[0];
        assert_eq!(diff.part, "xl/worksheets/sheet1.xml");
        assert_eq!(diff.kind, "attribute_added");
        assert!(
            diff.path.contains("@foo"),
            "expected diff path to include '@foo', got {}",
            diff.path
        );
        assert_eq!(
            details.ignore_paths,
            vec!["attribute_changed:dyDescent".to_string()]
        );
    }

    #[test]
    fn diff_workbooks_respects_ignore_path_kind_in_rules() {
        let expected = make_zip(&[
            (
                "xl/worksheets/sheet1.xml",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
  <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
</worksheet>"#,
            ),
            (
                "xl/worksheets/sheet2.xml",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
  <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
</worksheet>"#,
            ),
        ]);
        let actual = make_zip(&[
            (
                "xl/worksheets/sheet1.xml",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
  <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.30" foo="bar"/>
</worksheet>"#,
            ),
            (
                "xl/worksheets/sheet2.xml",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
  <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.30" foo="bar"/>
</worksheet>"#,
            ),
        ]);

        let args = Args {
            input: PathBuf::new(),
            format: WorkbookFormat::Xlsx,
            password: None,
            password_file: None,
            skip_encrypted: true,
            ignore_parts: Vec::new(),
            ignore_globs: Vec::new(),
            ignore_paths: Vec::new(),
            ignore_paths_in: Vec::new(),
            ignore_paths_kind: Vec::new(),
            ignore_paths_kind_in: Vec::new(),
            ignore_presets: Vec::new(),
            strict_calc_chain: false,
            diff_limit: 10,
            fail_on: RoundTripFailOn::Critical,
            recalc: false,
            render_smoke: false,
        };

        let details = diff_workbooks(&expected, &actual, &args).unwrap();
        assert_eq!(details.counts.total, 4);

        let args = Args {
            ignore_paths_kind_in: vec![
                "xl/worksheets/sheet1.xml:attribute_changed:dyDescent".to_string(),
            ],
            ..args
        };
        let details = diff_workbooks(&expected, &actual, &args).unwrap();

        assert_eq!(details.counts.total, 3);
        assert!(
            !details
                .top_differences
                .iter()
                .any(|d| d.part == "xl/worksheets/sheet1.xml" && d.kind == "attribute_changed"),
            "expected sheet1 attribute_changed diffs to be suppressed, got {:#?}",
            details.top_differences
        );
        assert!(
            details
                .top_differences
                .iter()
                .any(|d| d.part == "xl/worksheets/sheet2.xml" && d.kind == "attribute_changed"),
            "expected sheet2 attribute_changed diffs to remain, got {:#?}",
            details.top_differences
        );
        assert_eq!(
            details.ignore_paths,
            vec!["xl/worksheets/sheet1.xml:attribute_changed:dyDescent".to_string()]
        );
    }

    #[test]
    fn diff_workbooks_respects_ignore_glob_rules() {
        let expected = make_zip(&[
            ("xl/workbook.xml", "<workbook/>"),
            ("xl/media/image1.png", "a"),
        ]);
        let actual = make_zip(&[
            ("xl/workbook.xml", "<workbook/>"),
            ("xl/media/image1.png", "b"),
        ]);

        let args = Args {
            input: PathBuf::new(),
            format: WorkbookFormat::Xlsx,
            password: None,
            password_file: None,
            skip_encrypted: true,
            ignore_parts: Vec::new(),
            ignore_globs: Vec::new(),
            ignore_paths: Vec::new(),
            ignore_paths_in: Vec::new(),
            ignore_paths_kind: Vec::new(),
            ignore_paths_kind_in: Vec::new(),
            ignore_presets: Vec::new(),
            strict_calc_chain: false,
            diff_limit: 10,
            fail_on: RoundTripFailOn::Critical,
            recalc: false,
            render_smoke: false,
        };

        let details = diff_workbooks(&expected, &actual, &args).unwrap();
        assert_eq!(details.counts.total, 1);
        assert_eq!(details.part_stats.parts_total, 2);
        assert_eq!(details.top_differences.len(), 1);
        assert_eq!(details.top_differences[0].part, "xl/media/image1.png");

        let args = Args {
            ignore_globs: vec!["xl/media/*".to_string()],
            ..args
        };
        let details = diff_workbooks(&expected, &actual, &args).unwrap();
        assert_eq!(details.counts.total, 0, "expected diffs suppressed by ignore-glob");
        assert!(details.equal);
        assert!(details.top_differences.is_empty());
        assert_eq!(details.part_stats.parts_total, 1);
        assert_eq!(details.ignore_globs, vec!["xl/media/*".to_string()]);
    }

    #[test]
    fn strict_calc_chain_promotes_calcchain_diffs_to_critical() {
        let expected = make_zip(&[("xl/workbook.xml", "<workbook/>")]);
        let actual = make_zip(&[
            ("xl/workbook.xml", "<workbook/>"),
            ("xl/calcChain.xml", "<calcChain/>"),
        ]);

        let args = Args {
            input: PathBuf::new(),
            format: WorkbookFormat::Xlsx,
            password: None,
            password_file: None,
            skip_encrypted: true,
            ignore_parts: Vec::new(),
            ignore_globs: Vec::new(),
            ignore_paths: Vec::new(),
            ignore_paths_in: Vec::new(),
            ignore_paths_kind: Vec::new(),
            ignore_paths_kind_in: Vec::new(),
            ignore_presets: Vec::new(),
            strict_calc_chain: false,
            diff_limit: 10,
            fail_on: RoundTripFailOn::Critical,
            recalc: false,
            render_smoke: false,
        };

        let details = diff_workbooks(&expected, &actual, &args).unwrap();
        assert_eq!(details.counts.total, 1);
        assert_eq!(details.counts.critical, 0);
        assert_eq!(details.counts.warning, 1);
        assert_eq!(details.top_differences.len(), 1);
        assert_eq!(details.top_differences[0].part, "xl/calcChain.xml");
        assert_eq!(details.top_differences[0].kind, "extra_part");
        assert_eq!(details.top_differences[0].severity, "WARN");

        let args = Args {
            strict_calc_chain: true,
            ..args
        };

        let details = diff_workbooks(&expected, &actual, &args).unwrap();
        assert_eq!(details.counts.total, 1);
        assert_eq!(details.counts.critical, 1);
        assert_eq!(details.counts.warning, 0);
        assert_eq!(details.top_differences.len(), 1);
        assert_eq!(details.top_differences[0].part, "xl/calcChain.xml");
        assert_eq!(details.top_differences[0].kind, "extra_part");
        assert_eq!(details.top_differences[0].severity, "CRITICAL");
    }

    #[test]
    fn diff_workbooks_respects_ignore_presets() {
        let expected = make_zip(&[(
            "xl/worksheets/sheet1.xml",
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
  <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
</worksheet>"#,
        )]);
        let actual = make_zip(&[(
            "xl/worksheets/sheet1.xml",
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
  <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.30"/>
</worksheet>"#,
        )]);

        let args = Args {
            input: PathBuf::new(),
            format: WorkbookFormat::Xlsx,
            password: None,
            password_file: None,
            skip_encrypted: true,
            ignore_parts: Vec::new(),
            ignore_globs: Vec::new(),
            ignore_paths: Vec::new(),
            ignore_paths_in: Vec::new(),
            ignore_paths_kind: Vec::new(),
            ignore_paths_kind_in: Vec::new(),
            ignore_presets: Vec::new(),
            strict_calc_chain: false,
            diff_limit: 10,
            fail_on: RoundTripFailOn::Critical,
            recalc: false,
            render_smoke: false,
        };

        let details = diff_workbooks(&expected, &actual, &args).unwrap();
        assert!(
            details.counts.total > 0,
            "expected at least one diff without ignore presets"
        );

        let args = Args {
            ignore_presets: vec![xlsx_diff::IgnorePreset::ExcelVolatileIds],
            ..args
        };
        let details = diff_workbooks(&expected, &actual, &args).unwrap();
        assert_eq!(details.counts.total, 0, "expected diffs suppressed by preset");
        assert!(details.equal);
        assert!(details.ignore_paths.is_empty());
        assert_eq!(details.ignore_presets, vec!["excel-volatile-ids".to_string()]);
    }

    #[test]
    fn skips_encrypted_workbooks_without_password_and_decrypts_with_password() {
        let password = "secret";
        let encrypted_bytes = encrypt_ooxml_agile(&plain_xlsx_bytes(), password);

        let dir = tempfile::tempdir().expect("temp dir");
        let path = dir.path().join("encrypted.xlsx");
        std::fs::write(&path, encrypted_bytes).expect("write encrypted workbook");

        // Without password -> skipped.
        let args = Args::try_parse_from(["triage", "--input", path.to_str().unwrap()])
            .expect("parse args");
        let out = run(&args);
        for step in ["load", "round_trip", "diff", "recalc", "render"] {
            let res = out.steps.get(step).expect("step missing");
            assert_eq!(res.status, "skipped", "expected {step} to be skipped");
            assert_eq!(
                res.details
                    .as_ref()
                    .and_then(|v| v.get("reason"))
                    .and_then(|v| v.as_str()),
                Some("encrypted"),
                "expected {step} to have encrypted reason"
            );
            assert!(
                res.error.is_none(),
                "skipped {step} should not include error digest"
            );
        }

        // With password -> load OK (full triage may still find diffs).
        let args = Args::try_parse_from([
            "triage",
            "--input",
            path.to_str().unwrap(),
            "--password",
            password,
        ])
        .expect("parse args");
        let out = run(&args);
        let load = out.steps.get("load").expect("load step missing");
        assert_eq!(load.status, "ok");
        assert!(out.result.open_ok, "expected open_ok with correct password");
    }

    #[test]
    fn decrypts_and_triages_encrypted_xlsb_payload_when_password_is_provided() {
        let password = "secret";
        let encrypted_bytes = encrypt_ooxml_agile(&plain_xlsb_bytes(), password);

        let dir = tempfile::tempdir().expect("temp dir");
        let path = dir.path().join("encrypted.xlsb");
        std::fs::write(&path, encrypted_bytes).expect("write encrypted workbook");

        // Without password -> skipped (avoid polluting open failure stats).
        let args = Args::try_parse_from([
            "triage",
            "--input",
            path.to_str().unwrap(),
            "--format",
            "xlsb",
        ])
        .expect("parse args");
        let out = run(&args);
        for step in ["load", "round_trip", "diff", "recalc", "render"] {
            let res = out.steps.get(step).expect("step missing");
            assert_eq!(res.status, "skipped", "expected {step} to be skipped");
            assert_eq!(
                res.details
                    .as_ref()
                    .and_then(|v| v.get("reason"))
                    .and_then(|v| v.as_str()),
                Some("encrypted"),
                "expected {step} to have encrypted reason"
            );
            assert!(
                res.error.is_none(),
                "skipped {step} should not include error digest"
            );
        }
        assert!(!out.result.open_ok);

        // With password -> decrypt + triage (load/round-trip/diff). Recalc/render remain opt-in and
        // are skipped by default.
        let args = Args::try_parse_from([
            "triage",
            "--input",
            path.to_str().unwrap(),
            "--format",
            "xlsb",
            "--password",
            password,
        ])
        .expect("parse args");
        let out = run(&args);
        let load = out.steps.get("load").expect("load step missing");
        assert_eq!(load.status, "ok");
        assert!(out.result.open_ok, "expected open_ok with correct password");
    }
}
