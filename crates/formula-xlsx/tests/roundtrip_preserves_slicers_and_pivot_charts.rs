use formula_xlsx::XlsxPackage;
use std::path::PathBuf;

#[test]
fn roundtrip_preserves_slicers_timelines_and_pivot_charts() -> Result<(), Box<dyn std::error::Error>>
{
    let fixture_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures/pivot_slicers_and_chart.xlsx");
    let bytes = std::fs::read(&fixture_path)?;
    let package = XlsxPackage::from_bytes(&bytes)?;

    let slicers = package.pivot_slicer_parts()?;
    assert_eq!(slicers.slicers.len(), 1);
    assert_eq!(slicers.timelines.len(), 1);

    let slicer = &slicers.slicers[0];
    assert_eq!(slicer.name.as_deref(), Some("RegionSlicer"));
    assert_eq!(
        slicer.cache_part.as_deref(),
        Some("xl/slicerCaches/slicerCache1.xml")
    );
    assert_eq!(slicer.cache_name.as_deref(), Some("RegionSlicerCache"));
    assert_eq!(slicer.source_name.as_deref(), Some("PivotTable1"));
    assert_eq!(
        slicer.connected_pivot_tables,
        vec!["xl/pivotTables/pivotTable1.xml".to_string()]
    );
    assert_eq!(slicer.placed_on_drawings, vec!["xl/drawings/drawing1.xml"]);
    assert_eq!(
        slicer.placed_on_sheets,
        vec!["xl/worksheets/sheet1.xml".to_string()]
    );
    assert_eq!(slicer.placed_on_sheet_names, vec!["Sheet1".to_string()]);

    let timeline = &slicers.timelines[0];
    assert_eq!(timeline.name.as_deref(), Some("DateTimeline"));
    assert_eq!(
        timeline.cache_part.as_deref(),
        Some("xl/timelineCaches/timelineCacheDefinition1.xml")
    );
    assert_eq!(
        timeline.connected_pivot_tables,
        vec!["xl/pivotTables/pivotTable1.xml".to_string()]
    );
    assert_eq!(
        timeline.placed_on_drawings,
        vec!["xl/drawings/drawing1.xml".to_string()]
    );
    assert_eq!(
        timeline.placed_on_sheets,
        vec!["xl/worksheets/sheet1.xml".to_string()]
    );
    assert_eq!(timeline.placed_on_sheet_names, vec!["Sheet1".to_string()]);

    let charts = package.pivot_chart_parts()?;
    assert_eq!(charts.len(), 1);
    assert_eq!(charts[0].pivot_source_name.as_deref(), Some("PivotTable1"));
    assert_eq!(
        charts[0].pivot_source_part.as_deref(),
        Some("xl/pivotTables/pivotTable1.xml")
    );

    let charts_with_placement = package.pivot_chart_parts_with_placement()?;
    assert_eq!(charts_with_placement.len(), charts.len());
    assert_eq!(charts_with_placement[0].chart, charts[0]);
    assert_eq!(
        charts_with_placement[0].placed_on_drawings,
        vec!["xl/drawings/drawing1.xml"]
    );
    assert!(
        charts_with_placement[0]
            .placed_on_sheets
            .iter()
            .any(|sheet| sheet == "xl/worksheets/sheet1.xml"),
        "expected pivot chart to be placed on sheet1.xml, got {:?}",
        charts_with_placement[0].placed_on_sheets
    );

    let roundtrip_bytes = package.write_to_bytes()?;
    let roundtrip = XlsxPackage::from_bytes(&roundtrip_bytes)?;
    for entry in [
        "xl/slicers/slicer1.xml",
        "xl/slicers/_rels/slicer1.xml.rels",
        "xl/slicerCaches/slicerCache1.xml",
        "xl/slicerCaches/_rels/slicerCache1.xml.rels",
        "xl/timelines/timeline1.xml",
        "xl/timelines/_rels/timeline1.xml.rels",
        "xl/timelineCaches/timelineCacheDefinition1.xml",
        "xl/timelineCaches/_rels/timelineCacheDefinition1.xml.rels",
        "xl/charts/chart1.xml",
        "xl/charts/_rels/chart1.xml.rels",
        "xl/drawings/_rels/drawing1.xml.rels",
    ] {
        assert_eq!(
            package.part(entry).unwrap(),
            roundtrip.part(entry).unwrap(),
            "entry {entry} changed during round-trip"
        );
    }

    Ok(())
}

#[test]
fn slicer_sheet_names_are_best_effort_when_workbook_xml_is_malformed(
) -> Result<(), Box<dyn std::error::Error>> {
    let fixture_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures/pivot_slicers_and_chart.xlsx");
    let bytes = std::fs::read(&fixture_path)?;
    let mut package = XlsxPackage::from_bytes(&bytes)?;

    // Worksheet placement should still be discoverable via worksheet `.rels` parts even when
    // `xl/workbook.xml` cannot be parsed (sheet names become best-effort/empty).
    package.set_part("xl/workbook.xml", b"not xml".to_vec());

    let slicers = package.pivot_slicer_parts()?;
    assert_eq!(slicers.slicers.len(), 1);
    assert_eq!(
        slicers.slicers[0].placed_on_sheets,
        vec!["xl/worksheets/sheet1.xml".to_string()]
    );
    assert!(
        slicers.slicers[0].placed_on_sheet_names.is_empty(),
        "expected sheet names to be empty when workbook.xml is malformed"
    );

    Ok(())
}

#[test]
fn slicer_placement_is_best_effort_when_worksheet_rels_is_malformed(
) -> Result<(), Box<dyn std::error::Error>> {
    let fixture_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures/pivot_slicers_and_chart.xlsx");
    let bytes = std::fs::read(&fixture_path)?;
    let mut package = XlsxPackage::from_bytes(&bytes)?;

    // When worksheet `.rels` cannot be parsed, we should still parse slicers/timelines, but the
    // drawing -> sheet mapping becomes empty.
    package.set_part(
        "xl/worksheets/_rels/sheet1.xml.rels",
        b"not xml".to_vec(),
    );

    let slicers = package.pivot_slicer_parts()?;
    assert_eq!(slicers.slicers.len(), 1);
    assert_eq!(slicers.timelines.len(), 1);

    assert_eq!(
        slicers.slicers[0].placed_on_drawings,
        vec!["xl/drawings/drawing1.xml".to_string()]
    );
    assert!(slicers.slicers[0].placed_on_sheets.is_empty());
    assert!(slicers.slicers[0].placed_on_sheet_names.is_empty());

    assert_eq!(
        slicers.timelines[0].placed_on_drawings,
        vec!["xl/drawings/drawing1.xml".to_string()]
    );
    assert!(slicers.timelines[0].placed_on_sheets.is_empty());
    assert!(slicers.timelines[0].placed_on_sheet_names.is_empty());

    Ok(())
}

#[test]
fn slicer_placement_is_best_effort_when_drawing_rels_is_malformed(
) -> Result<(), Box<dyn std::error::Error>> {
    let fixture_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures/pivot_slicers_and_chart.xlsx");
    let bytes = std::fs::read(&fixture_path)?;
    let mut package = XlsxPackage::from_bytes(&bytes)?;

    // When drawing `.rels` cannot be parsed, we should still parse slicers/timelines, but
    // slicer/timeline -> drawing mapping becomes empty.
    package.set_part("xl/drawings/_rels/drawing1.xml.rels", b"not xml".to_vec());

    let slicers = package.pivot_slicer_parts()?;
    assert_eq!(slicers.slicers.len(), 1);
    assert_eq!(slicers.timelines.len(), 1);

    assert!(slicers.slicers[0].placed_on_drawings.is_empty());
    assert!(slicers.slicers[0].placed_on_sheets.is_empty());
    assert!(slicers.slicers[0].placed_on_sheet_names.is_empty());

    assert!(slicers.timelines[0].placed_on_drawings.is_empty());
    assert!(slicers.timelines[0].placed_on_sheets.is_empty());
    assert!(slicers.timelines[0].placed_on_sheet_names.is_empty());

    Ok(())
}

#[test]
fn slicer_cache_relationships_are_best_effort_when_slicer_rels_is_malformed(
) -> Result<(), Box<dyn std::error::Error>> {
    let fixture_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures/pivot_slicers_and_chart.xlsx");
    let bytes = std::fs::read(&fixture_path)?;
    let mut package = XlsxPackage::from_bytes(&bytes)?;

    // If the slicer part's `.rels` file is malformed, slicer cache relationships should simply be
    // treated as absent (rather than failing slicer parsing entirely).
    package.set_part("xl/slicers/_rels/slicer1.xml.rels", b"not xml".to_vec());

    let slicers = package.pivot_slicer_parts()?;
    assert_eq!(slicers.slicers.len(), 1);
    assert_eq!(slicers.timelines.len(), 1);

    // Cache metadata should be empty / missing.
    assert!(slicers.slicers[0].cache_part.is_none());
    assert!(slicers.slicers[0].cache_name.is_none());
    assert!(slicers.slicers[0].source_name.is_none());
    assert!(slicers.slicers[0].connected_pivot_tables.is_empty());

    // Placement should still be discoverable via drawings + worksheet rels.
    assert_eq!(
        slicers.slicers[0].placed_on_sheets,
        vec!["xl/worksheets/sheet1.xml".to_string()]
    );
    assert_eq!(
        slicers.slicers[0].placed_on_sheet_names,
        vec!["Sheet1".to_string()]
    );

    // Timeline parsing should be unaffected by slicer rels corruption.
    assert_eq!(
        slicers.timelines[0].cache_part.as_deref(),
        Some("xl/timelineCaches/timelineCacheDefinition1.xml")
    );

    Ok(())
}
