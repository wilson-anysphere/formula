use std::collections::BTreeSet;
use std::fs;
use std::path::{Path, PathBuf};

use formula_xlsx::XlsxPackage;

const GOLDEN_WIDTH_PX: u32 = 800;
const GOLDEN_HEIGHT_PX: u32 = 600;

#[test]
fn chart_fixture_corpus_complete() -> Result<(), Box<dyn std::error::Error>> {
    let repo_root: PathBuf = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../..");
    let xlsx_root = repo_root.join("fixtures/charts/xlsx");
    let models_root = repo_root.join("fixtures/charts/models");
    let golden_root = repo_root.join("fixtures/charts/golden/excel");

    let mut fixtures = read_dir_sorted(&xlsx_root)?;
    fixtures.retain(|path| path.extension().and_then(|s| s.to_str()) == Some("xlsx"));
    assert!(!fixtures.is_empty(), "no chart fixtures found under {xlsx_root:?}");

    let chartex_stems: BTreeSet<&'static str> = [
        "waterfall",
        "histogram",
        "pareto",
        "box-whisker",
        "treemap",
        "sunburst",
        "funnel",
    ]
    .into_iter()
    .collect();

    for fixture in fixtures {
        let stem = fixture
            .file_stem()
            .and_then(|s| s.to_str())
            .ok_or("invalid fixture filename")?;

        // Ensure golden image exists and is the documented fixed size.
        let golden_path = golden_root.join(format!("{stem}.png"));
        let golden_bytes = fs::read(&golden_path).map_err(|err| {
            format!(
                "missing golden image for {stem}: expected {path}: {err}",
                path = golden_path.display()
            )
        })?;
        let (w, h) = png_dimensions(&golden_bytes).map_err(|err| {
            format!(
                "invalid PNG for {stem}: {path}: {err}",
                path = golden_path.display()
            )
        })?;
        assert_eq!(
            (w, h),
            (GOLDEN_WIDTH_PX, GOLDEN_HEIGHT_PX),
            "golden image {stem} must be {GOLDEN_WIDTH_PX}x{GOLDEN_HEIGHT_PX}px",
        );

        // Ensure models exist (the stricter semantic equality check lives in
        // `chart_fixture_models_match.rs`).
        let model_dir = models_root.join(stem);
        let model_paths = read_dir_sorted(&model_dir).map_err(|err| {
            format!(
                "missing model directory for {stem}: expected {dir}: {err}",
                dir = model_dir.display()
            )
        })?;
        let json_count = model_paths
            .iter()
            .filter(|p| p.extension().and_then(|s| s.to_str()) == Some("json"))
            .count();
        assert!(
            json_count > 0,
            "fixture {stem}: expected at least one model json under {}",
            model_dir.display()
        );

        if chartex_stems.contains(stem) {
            let bytes = fs::read(&fixture)?;
            let pkg = XlsxPackage::from_bytes(&bytes)?;

            // Ensure the chartEx part exists.
            assert!(
                pkg.part("xl/charts/chartEx1.xml").is_some(),
                "fixture {stem}: expected xl/charts/chartEx1.xml",
            );

            // Ensure a non-default theme (Excel \"Office Theme\" is the default).
            let theme = pkg
                .part("xl/theme/theme1.xml")
                .ok_or("missing xl/theme/theme1.xml")?;
            let theme_str = std::str::from_utf8(theme)?;
            assert!(
                !theme_str.contains("name=\"Office Theme\""),
                "fixture {stem}: expected non-default theme (theme name should not be \"Office Theme\")",
            );

            // Ensure axis number formats + legend position + per-point override exist in chart XML.
            let chart = pkg
                .part("xl/charts/chart1.xml")
                .ok_or("missing xl/charts/chart1.xml")?;
            let chart_str = std::str::from_utf8(chart)?;
            assert!(
                chart_str.contains("<c:numFmt"),
                "fixture {stem}: expected axis number format (<c:numFmt .../>)"
            );
            assert!(
                chart_str.contains("<c:legendPos"),
                "fixture {stem}: expected legend position (<c:legendPos .../>)"
            );
            assert!(
                chart_str.contains("<c:dPt>"),
                "fixture {stem}: expected per-point override (<c:dPt> ... </c:dPt>)"
            );
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

fn png_dimensions(bytes: &[u8]) -> Result<(u32, u32), String> {
    const SIG: &[u8; 8] = b"\x89PNG\r\n\x1a\n";
    if bytes.len() < SIG.len() + 8 + 13 {
        return Err("file too small to be a valid PNG".to_string());
    }
    if &bytes[..8] != SIG {
        return Err("invalid PNG signature".to_string());
    }

    // First chunk should be IHDR.
    let len = u32::from_be_bytes(bytes[8..12].try_into().unwrap()) as usize;
    let chunk_type = &bytes[12..16];
    if chunk_type != b"IHDR" {
        return Err("first PNG chunk is not IHDR".to_string());
    }
    if len < 8 {
        return Err("IHDR chunk too small".to_string());
    }

    let data_start = 16;
    let data_end = data_start + len;
    if bytes.len() < data_end {
        return Err("truncated IHDR chunk".to_string());
    }

    let width = u32::from_be_bytes(bytes[data_start..data_start + 4].try_into().unwrap());
    let height = u32::from_be_bytes(bytes[data_start + 4..data_start + 8].try_into().unwrap());
    Ok((width, height))
}

