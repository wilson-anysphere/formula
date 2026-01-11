use std::error::Error;
use std::fs;
use std::path::{Path, PathBuf};

use formula_model::charts::ChartModel;
use formula_model::drawings::Anchor;
use formula_xlsx::drawingml::charts::parse_chart_space;
use formula_xlsx::XlsxPackage;
use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct ChartParts {
    drawing_part: String,
    chart_part: String,
    chart_ex_part: Option<String>,
    style_part: Option<String>,
    colors_part: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct ChartFixtureModel {
    chart_index: usize,
    sheet_name: Option<String>,
    anchor: Anchor,
    parts: ChartParts,
    model: ChartModel,
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
    let charts = pkg.extract_chart_objects()?;

    if print_parts {
        for (idx, chart) in charts.iter().enumerate() {
            eprintln!(
                "chart[{idx}]: sheet={:?} drawing_part={} chart_part={} chart_ex_part={:?} style_part={:?} colors_part={:?}",
                chart.sheet_name,
                chart.drawing_part,
                chart.parts.chart.path,
                chart.parts.chart_ex.as_ref().map(|p| p.path.as_str()),
                chart.parts.style.as_ref().map(|p| p.path.as_str()),
                chart.parts.colors.as_ref().map(|p| p.path.as_str()),
            );
        }
    }

    let workbook_out_dir = out_dir.join(&workbook_stem);
    recreate_dir(&workbook_out_dir)?;

    for (idx, chart) in charts.into_iter().enumerate() {
        let out_path = workbook_out_dir.join(format!("chart{idx}.json"));
        let model = parse_chart_space(&chart.parts.chart.bytes, &chart.parts.chart.path)?;
        let fixture_model = ChartFixtureModel {
            chart_index: idx,
            sheet_name: chart.sheet_name,
            anchor: chart.anchor,
            parts: ChartParts {
                drawing_part: chart.drawing_part,
                chart_part: chart.parts.chart.path,
                chart_ex_part: chart.parts.chart_ex.map(|p| p.path),
                style_part: chart.parts.style.map(|p| p.path),
                colors_part: chart.parts.colors.map(|p| p.path),
            },
            model,
        };
        let mut json = serde_json::to_string_pretty(&fixture_model)?;
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
