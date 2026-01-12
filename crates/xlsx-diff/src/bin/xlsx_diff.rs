use std::path::PathBuf;

use anyhow::Result;
use clap::Parser;
use globset::Glob;

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

    /// Minimum severity that will cause a non-zero exit code.
    #[arg(long, default_value = "critical")]
    fail_on: String,
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

    let expected = xlsx_diff::WorkbookArchive::open(&args.original)?;
    let actual = xlsx_diff::WorkbookArchive::open(&args.modified)?;
    let report = xlsx_diff::diff_archives_with_options(&expected, &actual, &options);

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
