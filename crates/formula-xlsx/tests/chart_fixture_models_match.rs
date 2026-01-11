use std::fs;
use std::path::{Path, PathBuf};

use formula_model::charts::Chart;
use pretty_assertions::assert_eq;
use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct ChartFixtureModel {
    chart_index: usize,
    chart: Chart,
}

#[test]
fn chart_fixture_models_match() -> Result<(), Box<dyn std::error::Error>> {
    let repo_root: PathBuf = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../..");
    let xlsx_root = repo_root.join("fixtures/charts/xlsx");
    let models_root = repo_root.join("fixtures/charts/models");

    let mut fixtures = read_dir_sorted(&xlsx_root)?;
    fixtures.retain(|path| path.extension().and_then(|s| s.to_str()) == Some("xlsx"));
    assert!(!fixtures.is_empty(), "no chart fixtures found under {xlsx_root:?}");

    for fixture in fixtures {
        let stem = fixture
            .file_stem()
            .and_then(|s| s.to_str())
            .ok_or("invalid fixture filename")?;
        let expected_dir = models_root.join(stem);

        let bytes = fs::read(&fixture)?;
        let pkg = formula_xlsx::XlsxPackage::from_bytes(&bytes)?;
        let charts = pkg.extract_charts()?;

        // Validate that we have an expected JSON file for every parsed chart and
        // no extra files lingering in the directory.
        let expected_paths = expected_chart_paths(&expected_dir)?;
        assert_eq!(
            expected_paths.len(),
            charts.len(),
            "fixture {stem}: expected {} chart model json files, found {} charts",
            expected_paths.len(),
            charts.len(),
        );

        for (idx, chart) in charts.into_iter().enumerate() {
            let expected_path = expected_dir.join(format!("chart{idx}.json"));
            let expected_bytes = fs::read(&expected_path).map_err(|err| {
                format!(
                    "fixture {stem}: missing expected model {path}: {err}",
                    path = expected_path.display()
                )
            })?;
            let expected: ChartFixtureModel = serde_json::from_slice(&expected_bytes).map_err(|err| {
                format!(
                    "fixture {stem}: failed to parse {path}: {err}",
                    path = expected_path.display()
                )
            })?;

            let actual = ChartFixtureModel {
                chart_index: idx,
                chart,
            };
            assert_eq!(expected, actual, "fixture {stem}: chart model mismatch");
        }
    }

    Ok(())
}

fn read_dir_sorted(dir: &Path) -> Result<Vec<PathBuf>, std::io::Error> {
    let mut entries: Vec<_> = fs::read_dir(dir)?
        .filter_map(|entry| entry.ok())
        .map(|entry| entry.path())
        .collect();
    entries.sort();
    Ok(entries)
}

fn expected_chart_paths(dir: &Path) -> Result<Vec<PathBuf>, Box<dyn std::error::Error>> {
    if !dir.exists() {
        return Err(format!("missing expected model directory: {}", dir.display()).into());
    }
    let mut paths = read_dir_sorted(dir)?;
    paths.retain(|path| {
        path.extension().and_then(|s| s.to_str()) == Some("json")
            && path
                .file_name()
                .and_then(|s| s.to_str())
                .is_some_and(|name| name.starts_with("chart") && name.ends_with(".json"))
    });
    Ok(paths)
}

