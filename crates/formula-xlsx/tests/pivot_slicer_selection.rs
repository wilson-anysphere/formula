use formula_xlsx::{SlicerSelectionState, XlsxPackage};
use std::collections::HashSet;
use std::path::PathBuf;

#[test]
fn parses_slicer_selection_state() -> Result<(), Box<dyn std::error::Error>> {
    let fixture_path =
        PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/slicer-selection.xlsx");
    let bytes = std::fs::read(&fixture_path)?;
    let package = XlsxPackage::from_bytes(&bytes)?;

    let parts = package.pivot_slicer_parts()?;
    assert_eq!(parts.slicers.len(), 1);
    let slicer = &parts.slicers[0];

    assert_eq!(
        slicer.selection.available_items,
        vec!["East".to_string(), "West".to_string(), "North".to_string()]
    );

    let expected = HashSet::from(["East".to_string()]);
    assert_eq!(slicer.selection.selected_items, Some(expected));
    Ok(())
}

#[test]
fn writes_slicer_selection_state() -> Result<(), Box<dyn std::error::Error>> {
    let fixture_path =
        PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/slicer-selection.xlsx");
    let bytes = std::fs::read(&fixture_path)?;
    let mut package = XlsxPackage::from_bytes(&bytes)?;

    let parts = package.pivot_slicer_parts()?;
    let slicer = parts.slicers.first().expect("fixture slicer");
    let cache_part = slicer.cache_part.as_deref().expect("cache part");

    let selection = SlicerSelectionState {
        available_items: slicer.selection.available_items.clone(),
        selected_items: Some(HashSet::from(["West".to_string()])),
    };
    package.set_slicer_cache_selection(cache_part, &selection)?;

    let written = package.write_to_bytes()?;
    let roundtrip = XlsxPackage::from_bytes(&written)?;
    let parts = roundtrip.pivot_slicer_parts()?;
    let slicer = parts.slicers.first().expect("fixture slicer roundtrip");

    assert_eq!(
        slicer.selection.selected_items,
        Some(HashSet::from(["West".to_string()]))
    );

    Ok(())
}

#[test]
fn slicer_selection_defaults_to_all_when_missing() -> Result<(), Box<dyn std::error::Error>> {
    let fixture_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures/pivot_slicers_and_chart.xlsx");
    let bytes = std::fs::read(&fixture_path)?;
    let package = XlsxPackage::from_bytes(&bytes)?;

    let parts = package.pivot_slicer_parts()?;
    assert_eq!(parts.slicers.len(), 1);
    assert_eq!(parts.slicers[0].selection.selected_items, None);
    Ok(())
}

#[test]
fn parses_timeline_selection_state() -> Result<(), Box<dyn std::error::Error>> {
    let fixture_path =
        PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/timeline-selection.xlsx");
    let bytes = std::fs::read(&fixture_path)?;
    let package = XlsxPackage::from_bytes(&bytes)?;

    let parts = package.pivot_slicer_parts()?;
    assert_eq!(parts.timelines.len(), 1);
    let timeline = &parts.timelines[0];

    assert_eq!(timeline.selection.start.as_deref(), Some("2024-01-01"));
    assert_eq!(timeline.selection.end.as_deref(), Some("2024-02-29"));
    Ok(())
}

#[test]
fn timeline_selection_defaults_to_empty_when_missing() -> Result<(), Box<dyn std::error::Error>> {
    let fixture_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures/pivot_slicers_and_chart.xlsx");
    let bytes = std::fs::read(&fixture_path)?;
    let package = XlsxPackage::from_bytes(&bytes)?;

    let parts = package.pivot_slicer_parts()?;
    assert_eq!(parts.timelines.len(), 1);
    assert_eq!(parts.timelines[0].selection.start, None);
    assert_eq!(parts.timelines[0].selection.end, None);
    Ok(())
}
