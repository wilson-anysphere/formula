#[cfg(not(target_arch = "wasm32"))]
use std::error::Error;
#[cfg(not(target_arch = "wasm32"))]
use std::fs;
#[cfg(not(target_arch = "wasm32"))]
use std::path::{Path, PathBuf};

#[cfg(not(target_arch = "wasm32"))]
use formula_model::charts::ChartModel;
#[cfg(not(target_arch = "wasm32"))]
use formula_model::drawings::Anchor;
#[cfg(not(target_arch = "wasm32"))]
use formula_xlsx::drawingml::charts::{parse_chart_ex, parse_chart_space};
#[cfg(not(target_arch = "wasm32"))]
use formula_xlsx::XlsxPackage;
#[cfg(not(target_arch = "wasm32"))]
use serde::{Deserialize, Serialize};

#[cfg(not(target_arch = "wasm32"))]
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct ChartParts {
    drawing_part: String,
    chart_part: String,
    chart_ex_part: Option<String>,
    style_part: Option<String>,
    colors_part: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    user_shapes_part: Option<String>,
}

#[cfg(not(target_arch = "wasm32"))]
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct ChartFixtureModel {
    chart_index: usize,
    sheet_name: Option<String>,
    anchor: Anchor,
    drawing_rel_id: String,
    drawing_object_id: Option<u32>,
    drawing_object_name: Option<String>,
    parts: ChartParts,
    model: ChartModel,
}

#[cfg(not(target_arch = "wasm32"))]
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct ChartFixtureModels {
    chart_index: usize,
    sheet_name: Option<String>,
    anchor: Anchor,
    drawing_rel_id: String,
    drawing_object_id: Option<u32>,
    drawing_object_name: Option<String>,
    parts: ChartParts,
    model_chart_space: ChartModel,
    model_chart_ex: Option<ChartModel>,
}

#[cfg(not(target_arch = "wasm32"))]
fn usage() -> &'static str {
    "dump_chart_models <path.xlsx|dir>\n\
  [--out-dir <dir>]\n\
  [--print-parts]\n\
  [--use-chart-object-model]\n\
  [--emit-both-models]\n\
\n\
Writes one JSON file per extracted chart under:\n\
  <out-dir>/<workbook-stem>/chart<N>.json\n\
\n\
Defaults:\n\
  --out-dir fixtures/charts/models\n\
"
}

#[cfg(target_arch = "wasm32")]
fn main() {}

#[cfg(not(target_arch = "wasm32"))]
fn main() -> Result<(), Box<dyn Error>> {
    let mut xlsx_path: Option<PathBuf> = None;
    let mut out_dir: PathBuf = PathBuf::from("fixtures/charts/models");
    let mut print_parts = false;
    let mut use_chart_object_model = false;
    let mut emit_both_models = false;

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
            "--use-chart-object-model" => {
                use_chart_object_model = true;
            }
            "--emit-both-models" => {
                emit_both_models = true;
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

    let xlsx_path = xlsx_path.ok_or_else(|| format!("missing <path.xlsx|dir>\n\n{}", usage()))?;
    let inputs = if xlsx_path.is_dir() {
        collect_xlsx_files(&xlsx_path)?
    } else {
        vec![xlsx_path]
    };

    for xlsx_path in inputs {
        dump_one(
            &xlsx_path,
            &out_dir,
            print_parts,
            use_chart_object_model,
            emit_both_models,
        )?;
    }

    Ok(())
}

#[cfg(not(target_arch = "wasm32"))]
fn dump_one(
    xlsx_path: &Path,
    out_dir: &Path,
    print_parts: bool,
    use_chart_object_model: bool,
    emit_both_models: bool,
) -> Result<(), Box<dyn Error>> {
    let workbook_stem = xlsx_path
        .file_stem()
        .and_then(|s| s.to_str())
        .ok_or("invalid xlsx path (missing file stem)")?
        .to_string();

    let bytes = fs::read(xlsx_path)?;
    let pkg = XlsxPackage::from_bytes(&bytes)?;
    let charts = pkg.extract_chart_objects()?;

    if print_parts {
        eprintln!("workbook: {}", xlsx_path.display());
        for (idx, chart) in charts.iter().enumerate() {
            eprintln!(
                "  chart[{idx}]: sheet={:?} drawing_part={} chart_part={} chart_ex_part={:?} style_part={:?} colors_part={:?} user_shapes_part={:?}",
                chart.sheet_name,
                chart.drawing_part,
                chart.parts.chart.path,
                chart.parts.chart_ex.as_ref().map(|p| p.path.as_str()),
                chart.parts.style.as_ref().map(|p| p.path.as_str()),
                chart.parts.colors.as_ref().map(|p| p.path.as_str()),
                chart.parts.user_shapes.as_ref().map(|p| p.path.as_str()),
            );
        }
    }

    let workbook_out_dir = out_dir.join(&workbook_stem);
    recreate_dir(&workbook_out_dir)?;

    for (idx, chart) in charts.into_iter().enumerate() {
        let out_path = workbook_out_dir.join(format!("chart{idx}.json"));
        let drawing_rel_id = chart.drawing_rel_id;
        let drawing_object_id = chart.drawing_object_id;
        let drawing_object_name = chart.drawing_object_name;

        let parts = ChartParts {
            drawing_part: chart.drawing_part,
            chart_part: chart.parts.chart.path.clone(),
            chart_ex_part: chart.parts.chart_ex.as_ref().map(|p| p.path.clone()),
            style_part: chart.parts.style.as_ref().map(|p| p.path.clone()),
            colors_part: chart.parts.colors.as_ref().map(|p| p.path.clone()),
            user_shapes_part: chart.parts.user_shapes.as_ref().map(|p| p.path.clone()),
        };

        let mut json = if emit_both_models {
            let model_chart_space =
                parse_chart_space(&chart.parts.chart.bytes, &chart.parts.chart.path)?;
            let model_chart_ex = match chart.parts.chart_ex.as_ref() {
                Some(part) => Some(parse_chart_ex(&part.bytes, &part.path)?),
                None => None,
            };

            let fixture_model = ChartFixtureModels {
                chart_index: idx,
                sheet_name: chart.sheet_name,
                anchor: chart.anchor,
                drawing_rel_id,
                drawing_object_id,
                drawing_object_name,
                parts,
                model_chart_space,
                model_chart_ex,
            };
            serde_json::to_string_pretty(&fixture_model)?
        } else {
            let model = if use_chart_object_model {
                chart.model.ok_or_else(|| {
                    format!(
                        "chart[{idx}] did not include a parsed model (use --print-parts for debugging)"
                    )
                })?
            } else {
                parse_chart_space(&chart.parts.chart.bytes, &chart.parts.chart.path)?
            };

            let fixture_model = ChartFixtureModel {
                chart_index: idx,
                sheet_name: chart.sheet_name,
                anchor: chart.anchor,
                drawing_rel_id,
                drawing_object_id,
                drawing_object_name,
                parts,
                model,
            };
            serde_json::to_string_pretty(&fixture_model)?
        };
        json.push('\n');
        fs::write(&out_path, json)?;
    }

    Ok(())
}

#[cfg(not(target_arch = "wasm32"))]
fn recreate_dir(path: &Path) -> std::io::Result<()> {
    match fs::remove_dir_all(path) {
        Ok(()) => {}
        Err(err) if err.kind() == std::io::ErrorKind::NotFound => {}
        Err(err) => return Err(err),
    }
    fs::create_dir_all(path)
}

#[cfg(not(target_arch = "wasm32"))]
fn collect_xlsx_files(dir: &Path) -> Result<Vec<PathBuf>, Box<dyn Error>> {
    let mut out: Vec<PathBuf> = fs::read_dir(dir)?
        .filter_map(|entry| entry.ok())
        .map(|entry| entry.path())
        .filter(|path| path.extension().and_then(|s| s.to_str()) == Some("xlsx"))
        .collect();
    out.sort();
    Ok(out)
}
