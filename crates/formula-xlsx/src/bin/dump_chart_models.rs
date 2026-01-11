use std::error::Error;
use std::fs;
use std::path::{Path, PathBuf};

use formula_model::charts::Chart;
use formula_xlsx::XlsxPackage;
use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct ChartFixtureModel {
    chart_index: usize,
    chart: Chart,
}

fn usage() -> &'static str {
    "dump_chart_models <path.xlsx> [--out-dir <dir>] [--print-parts]\n\
\n\
Writes one JSON file per extracted chart under:\n\
  <out-dir>/<workbook-stem>/chart<N>.json\n\
\n\
Defaults:\n\
  --out-dir fixtures/charts/models\n\
"
}

fn main() -> Result<(), Box<dyn Error>> {
    let mut xlsx_path: Option<PathBuf> = None;
    let mut out_dir: PathBuf = PathBuf::from("fixtures/charts/models");
    let mut print_parts = false;

    let mut args = std::env::args().skip(1).peekable();
    while let Some(arg) = args.next() {
        match arg.as_str() {
            "--help" | "-h" => {
                eprintln!("{}", usage());
                return Ok(());
            }
            "--out-dir" => {
                let value = args
                    .next()
                    .ok_or("--out-dir expects a value (directory path)")?;
                out_dir = PathBuf::from(value);
            }
            "--print-parts" => {
                print_parts = true;
            }
            flag if flag.starts_with('-') => {
                return Err(format!("unknown flag: {flag}\n\n{}", usage()).into());
            }
            value => {
                if xlsx_path.is_some() {
                    return Err(format!("unexpected extra argument: {value}\n\n{}", usage()).into());
                }
                xlsx_path = Some(PathBuf::from(value));
            }
        }
    }

    let xlsx_path = xlsx_path.ok_or_else(|| format!("missing <path.xlsx>\n\n{}", usage()))?;
    let workbook_stem = xlsx_path
        .file_stem()
        .and_then(|s| s.to_str())
        .ok_or("invalid xlsx path (missing file stem)")?
        .to_string();

    let bytes = fs::read(&xlsx_path)?;
    let pkg = XlsxPackage::from_bytes(&bytes)?;
    let charts = pkg.extract_charts()?;

    if print_parts {
        for (idx, chart) in charts.iter().enumerate() {
            eprintln!(
                "chart[{idx}]: sheet={:?} drawing_part={} chart_part={:?} rel_id={}",
                chart.sheet_name, chart.drawing_part, chart.chart_part, chart.rel_id
            );
        }
    }

    let workbook_out_dir = out_dir.join(&workbook_stem);
    recreate_dir(&workbook_out_dir)?;

    for (idx, chart) in charts.into_iter().enumerate() {
        let out_path = workbook_out_dir.join(format!("chart{idx}.json"));
        let model = ChartFixtureModel {
            chart_index: idx,
            chart,
        };
        let mut json = serde_json::to_string_pretty(&model)?;
        json.push('\n');
        fs::write(&out_path, json)?;
    }

    Ok(())
}

fn recreate_dir(path: &Path) -> std::io::Result<()> {
    match fs::remove_dir_all(path) {
        Ok(()) => {}
        Err(err) if err.kind() == std::io::ErrorKind::NotFound => {}
        Err(err) => return Err(err),
    }
    fs::create_dir_all(path)
}

