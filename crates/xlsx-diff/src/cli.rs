use std::io::Write;
use std::path::PathBuf;

use anyhow::Result;
use clap::{Parser, ValueEnum};
use globset::Glob;
use serde::Serialize;

use crate::{DiffInput, DiffOptions, IgnorePathRule, Severity};

#[derive(Clone, Debug, ValueEnum)]
enum OutputFormat {
    Text,
    Json,
}

/// CLI arguments shared by both the `xlsx-diff` and deprecated `xlsb-diff` binaries.
///
/// This lives in the library crate so `xlsb-diff` can be a true thin wrapper around
/// the single diff implementation and command-line surface area.
#[derive(Parser)]
#[command(about = "Diff two XLSX/XLSM/XLSB workbooks at the Open Packaging Convention part level.")]
pub struct Args {
    /// Original workbook.
    original: PathBuf,

    /// Modified workbook (e.g. round-tripped output).
    modified: PathBuf,

    /// Password for both workbooks (if they are encrypted).
    #[arg(long)]
    password: Option<String>,

    /// Password for the original workbook (if encrypted). Overrides `--password`.
    #[arg(long = "original-password")]
    original_password: Option<String>,

    /// Password for the modified workbook (if encrypted). Overrides `--password`.
    #[arg(long = "modified-password")]
    modified_password: Option<String>,

    /// Exact part names to ignore (repeatable).
    #[arg(long = "ignore-part")]
    ignore_parts: Vec<String>,

    /// Glob patterns to ignore (repeatable).
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

    /// Treat calcChain-related diffs as CRITICAL instead of downgrading them to WARNING.
    #[arg(long = "strict-calc-chain")]
    strict_calc_chain: bool,

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
    ignore_paths: Vec<JsonIgnorePathRule<'a>>,
    strict_calc_chain: bool,
    counts: JsonCounts,
    diffs: Vec<JsonDiff<'a>>,
}

#[derive(Debug, Serialize)]
struct JsonIgnorePathRule<'a> {
    part: Option<&'a str>,
    path_substring: &'a str,
    kind: Option<&'a str>,
}

pub fn run() -> Result<()> {
    let args = Args::parse();
    run_with_args(args)
}

/// Parse CLI arguments.
///
/// This exists so wrapper binaries (e.g. the deprecated `xlsb-diff`) can parse
/// the shared CLI args without taking a direct dependency on `clap`.
pub fn parse_args() -> Args {
    Args::parse()
}

pub fn run_with_args(args: Args) -> Result<()> {
    let threshold = parse_severity(&args.fail_on)?;

    let options = DiffOptions {
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
        ignore_paths: build_ignore_path_rules(&args)?,
        strict_calc_chain: args.strict_calc_chain,
    };

    for pattern in &options.ignore_globs {
        Glob::new(pattern)?;
    }
    for rule in &options.ignore_paths {
        if let Some(part) = &rule.part {
            if part.contains('*') || part.contains('?') {
                Glob::new(part)?;
            }
        }
    }

    let original_pw = args
        .original_password
        .as_deref()
        .or(args.password.as_deref());
    let modified_pw = args
        .modified_password
        .as_deref()
        .or(args.password.as_deref());

    let mut report = crate::diff_workbooks_with_inputs_and_options(
        DiffInput {
            path: &args.original,
            password: original_pw,
        },
        DiffInput {
            path: &args.modified,
            password: modified_pw,
        },
        &options,
    )?;
    report.differences.sort_by(|a, b| {
        let rank = |s: Severity| match s {
            Severity::Critical => 0u8,
            Severity::Warning => 1u8,
            Severity::Info => 2u8,
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
            let mut ignore_paths: Vec<String> = options
                .ignore_paths
                .iter()
                .map(|r| match (&r.part, r.kind.as_deref()) {
                    (Some(part), Some(kind)) => format!("{part}:{kind}:{}", r.path_substring),
                    (Some(part), None) => format!("{part}:{}", r.path_substring),
                    (None, Some(kind)) => format!("{kind}:{}", r.path_substring),
                    (None, None) => r.path_substring.clone(),
                })
                .collect();
            ignore_paths.sort();
            println!(
                "  ignore-path: {}",
                if ignore_paths.is_empty() {
                    "(none)".to_string()
                } else {
                    ignore_paths.join(", ")
                }
            );
            println!();

            if report.is_empty() {
                println!("No differences.");
                return Ok(());
            }

            println!(
                "Summary: critical={} warn={} info={}",
                report.count(Severity::Critical),
                report.count(Severity::Warning),
                report.count(Severity::Info)
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
            let mut ignore_paths: Vec<JsonIgnorePathRule<'_>> = options
                .ignore_paths
                .iter()
                .map(|rule| JsonIgnorePathRule {
                    part: rule.part.as_deref(),
                    path_substring: rule.path_substring.as_str(),
                    kind: rule.kind.as_deref(),
                })
                .collect();
            ignore_paths.sort_by(|a, b| {
                (a.part, a.kind, a.path_substring).cmp(&(b.part, b.kind, b.path_substring))
            });

            let counts = JsonCounts {
                critical: report.count(Severity::Critical),
                warning: report.count(Severity::Warning),
                info: report.count(Severity::Info),
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
                ignore_paths,
                strict_calc_chain: options.strict_calc_chain,
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

fn parse_severity(input: &str) -> Result<Severity> {
    match input.to_ascii_lowercase().as_str() {
        "critical" | "crit" => Ok(Severity::Critical),
        "warning" | "warn" => Ok(Severity::Warning),
        "info" => Ok(Severity::Info),
        _ => anyhow::bail!("unknown severity '{input}' (expected: critical|warning|info)"),
    }
}

fn normalize_ignore_pattern(input: &str) -> String {
    let trimmed = input.trim();
    let normalized = trimmed.replace('\\', "/");
    normalized.trim_start_matches('/').to_string()
}

fn severity_label(severity: Severity) -> &'static str {
    match severity {
        Severity::Critical => "critical",
        Severity::Warning => "warning",
        Severity::Info => "info",
    }
}

fn build_ignore_path_rules(args: &Args) -> Result<Vec<IgnorePathRule>> {
    let mut rules = Vec::new();

    for pattern in &args.ignore_paths {
        let trimmed = pattern.trim();
        if trimmed.is_empty() {
            continue;
        }
        rules.push(IgnorePathRule {
            part: None,
            path_substring: trimmed.replace('\\', "/"),
            kind: None,
        });
    }

    for scoped in &args.ignore_paths_in {
        let trimmed = scoped.trim();
        if trimmed.is_empty() {
            continue;
        }
        let Some((part, pattern)) = trimmed.split_once(':') else {
            anyhow::bail!(
                "invalid --ignore-path-in '{trimmed}' (expected format: <part_glob>:<path_substring>)"
            );
        };
        let part = normalize_ignore_pattern(part);
        let pattern = pattern.trim();
        if part.is_empty() || pattern.is_empty() {
            continue;
        }
        rules.push(IgnorePathRule {
            part: Some(part),
            path_substring: pattern.replace('\\', "/"),
            kind: None,
        });
    }

    Ok(rules)
}
