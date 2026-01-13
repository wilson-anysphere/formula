use std::io::Write;
use std::path::PathBuf;

use anyhow::Result;
use clap::{Parser, ValueEnum};
use globset::Glob;
use serde::Serialize;

#[derive(Clone, Debug, ValueEnum)]
enum OutputFormat {
    Text,
    Json,
}

#[derive(Parser)]
#[command(about = "Diff two XLSX/XLSM/XLSB workbooks at the Open Packaging Convention part level.")]
struct Args {
    /// Original workbook.
    original: PathBuf,

    /// Modified workbook (e.g. round-tripped output).
    modified: PathBuf,

    /// Exact part names to ignore (repeatable).
    #[arg(long = "ignore-part")]
    ignore_parts: Vec<String>,

    /// Glob patterns to ignore (repeatable).
    #[arg(long = "ignore-glob")]
    ignore_globs: Vec<String>,

    /// Output format.
    #[arg(long, value_enum, default_value_t = OutputFormat::Text)]
    format: OutputFormat,

    /// Maximum number of diffs to include in JSON output (default: unlimited).
    #[arg(long)]
    max_diffs: Option<usize>,

    /// Minimum severity that will cause a non-zero exit code.
    #[arg(long, default_value = "critical")]
    fail_on: String,
}

#[derive(Debug, Serialize)]
struct JsonCounts {
    critical: usize,
    warning: usize,
    info: usize,
}

#[derive(Debug, Serialize)]
struct JsonDiff<'a> {
    severity: &'static str,
    part: &'a str,
    path: &'a str,
    kind: &'a str,
    expected: Option<&'a str>,
    actual: Option<&'a str>,
}

#[derive(Debug, Serialize)]
struct JsonReport<'a> {
    original: &'a str,
    modified: &'a str,
    ignore_parts: Vec<&'a str>,
    ignore_globs: Vec<&'a str>,
    counts: JsonCounts,
    diffs: Vec<JsonDiff<'a>>,
}

fn main() -> Result<()> {
    let args = Args::parse();
    let threshold = parse_severity(&args.fail_on)?;

    let options = xlsx_diff::DiffOptions {
        ignore_parts: args
            .ignore_parts
            .iter()
            .map(|s| normalize_ignore_pattern(s))
            .filter(|s| !s.is_empty())
            .collect(),
        ignore_globs: args
            .ignore_globs
            .iter()
            .map(|s| normalize_ignore_pattern(s))
            .filter(|s| !s.is_empty())
            .collect(),
    };

    for pattern in &options.ignore_globs {
        Glob::new(pattern)?;
    }

    let mut report =
        xlsx_diff::diff_workbooks_with_options(&args.original, &args.modified, &options)?;
    report.differences.sort_by(|a, b| {
        let rank = |s: xlsx_diff::Severity| match s {
            xlsx_diff::Severity::Critical => 0u8,
            xlsx_diff::Severity::Warning => 1u8,
            xlsx_diff::Severity::Info => 2u8,
        };
        (
            rank(a.severity),
            a.part.as_str(),
            a.path.as_str(),
            a.kind.as_str(),
        )
            .cmp(&(
                rank(b.severity),
                b.part.as_str(),
                b.path.as_str(),
                b.kind.as_str(),
            ))
    });

    match args.format {
        OutputFormat::Text => {
            println!("Workbook diff report (OPC parts)");
            println!("  original: {}", args.original.display());
            println!("  modified: {}", args.modified.display());
            let mut parts: Vec<&str> = options.ignore_parts.iter().map(|s| s.as_str()).collect();
            parts.sort();
            println!(
                "  ignore-part: {}",
                if parts.is_empty() {
                    "(none)".to_string()
                } else {
                    parts.join(", ")
                }
            );
            let mut globs: Vec<&str> = options.ignore_globs.iter().map(|s| s.as_str()).collect();
            globs.sort();
            println!(
                "  ignore-glob: {}",
                if globs.is_empty() {
                    "(none)".to_string()
                } else {
                    globs.join(", ")
                }
            );
            println!();

            if report.is_empty() {
                println!("No differences.");
                return Ok(());
            }

            println!(
                "Summary: critical={} warn={} info={}",
                report.count(xlsx_diff::Severity::Critical),
                report.count(xlsx_diff::Severity::Warning),
                report.count(xlsx_diff::Severity::Info)
            );
            println!();

            for diff in &report.differences {
                print!("{diff}");
            }

            if report.has_at_least(threshold) {
                std::process::exit(1);
            }

            Ok(())
        }
        OutputFormat::Json => {
            let original = args.original.to_string_lossy().into_owned();
            let modified = args.modified.to_string_lossy().into_owned();

            let mut ignore_parts: Vec<&str> =
                options.ignore_parts.iter().map(|s| s.as_str()).collect();
            ignore_parts.sort();
            let mut ignore_globs: Vec<&str> =
                options.ignore_globs.iter().map(|s| s.as_str()).collect();
            ignore_globs.sort();

            let counts = JsonCounts {
                critical: report.count(xlsx_diff::Severity::Critical),
                warning: report.count(xlsx_diff::Severity::Warning),
                info: report.count(xlsx_diff::Severity::Info),
            };

            let limit = args.max_diffs.unwrap_or(report.differences.len());
            let diffs = report
                .differences
                .iter()
                .take(limit)
                .map(|diff| JsonDiff {
                    severity: severity_label(diff.severity),
                    part: diff.part.as_str(),
                    path: diff.path.as_str(),
                    kind: diff.kind.as_str(),
                    expected: diff.expected.as_deref(),
                    actual: diff.actual.as_deref(),
                })
                .collect();

            let json_report = JsonReport {
                original: &original,
                modified: &modified,
                ignore_parts,
                ignore_globs,
                counts,
                diffs,
            };

            let stdout = std::io::stdout();
            let mut handle = stdout.lock();
            serde_json::to_writer(&mut handle, &json_report)?;
            handle.write_all(b"\n")?;

            if report.has_at_least(threshold) {
                std::process::exit(1);
            }

            Ok(())
        }
    }
}

fn parse_severity(input: &str) -> Result<xlsx_diff::Severity> {
    match input.to_ascii_lowercase().as_str() {
        "critical" | "crit" => Ok(xlsx_diff::Severity::Critical),
        "warning" | "warn" => Ok(xlsx_diff::Severity::Warning),
        "info" => Ok(xlsx_diff::Severity::Info),
        _ => anyhow::bail!("unknown severity '{input}' (expected: critical|warning|info)"),
    }
}

fn normalize_ignore_pattern(input: &str) -> String {
    let trimmed = input.trim();
    let normalized = trimmed.replace('\\', "/");
    normalized.trim_start_matches('/').to_string()
}

fn severity_label(severity: xlsx_diff::Severity) -> &'static str {
    match severity {
        xlsx_diff::Severity::Critical => "critical",
        xlsx_diff::Severity::Warning => "warning",
        xlsx_diff::Severity::Info => "info",
    }
}
