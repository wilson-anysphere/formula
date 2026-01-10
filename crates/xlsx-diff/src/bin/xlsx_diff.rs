use std::path::PathBuf;

use anyhow::Result;
use clap::Parser;

#[derive(Parser)]
#[command(about = "Diff two XLSX/XLSM workbooks at the OpenXML part level.")]
struct Args {
    /// Original workbook.
    original: PathBuf,

    /// Modified workbook (e.g. round-tripped output).
    modified: PathBuf,

    /// Minimum severity that will cause a non-zero exit code.
    #[arg(long, default_value = "critical")]
    fail_on: String,
}

fn main() -> Result<()> {
    let args = Args::parse();
    let threshold = parse_severity(&args.fail_on)?;

    let report = xlsx_diff::diff_workbooks(&args.original, &args.modified)?;

    println!("XLSX diff report");
    println!("  original: {}", args.original.display());
    println!("  modified: {}", args.modified.display());
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
