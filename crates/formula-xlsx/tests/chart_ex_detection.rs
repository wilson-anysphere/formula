use formula_model::charts::{ChartKind, TextModel};
use formula_xlsx::XlsxPackage;

fn load_fixture(name: &str) -> Vec<u8> {
    std::fs::read(format!(
        "{}/../../fixtures/xlsx/charts-ex/{name}.xlsx",
        env!("CARGO_MANIFEST_DIR")
    ))
    .expect("fixture exists")
}

#[test]
fn detects_chart_ex_parts_and_parses_kind() {
    for (fixture_name, expected_kind, expected_title) in [
        ("waterfall", "waterfall", "Waterfall"),
        ("histogram", "histogram", "Histogram"),
        ("treemap", "treemap", "Treemap"),
        ("sunburst", "sunburst", "Sunburst"),
        ("funnel", "funnel", "Funnel"),
        ("box-whisker", "boxWhisker", "Box & Whisker"),
        ("pareto", "pareto", "Pareto"),
        ("map", "regionMap", "Map"),
    ] {
        let bytes = load_fixture(fixture_name);
        let pkg = XlsxPackage::from_bytes(&bytes).expect("parse package");

        let charts = pkg.extract_chart_objects().expect("extract chart objects");
        assert!(
            !charts.is_empty(),
            "expected at least one chart object in {fixture_name}.xlsx"
        );

        let chart = &charts[0];
        let chart_part = &chart.parts.chart;
        let chart_rels_path = chart_part
            .rels_path
            .as_deref()
            .expect("fixture charts should include chart rels");
        assert_eq!(
            chart_part.rels_bytes.as_deref(),
            pkg.part(chart_rels_path),
            "chart rels bytes should be embedded on the extracted OpcPart for {fixture_name}.xlsx"
        );
        let chart_rels_xml = std::str::from_utf8(
            chart_part
                .rels_bytes
                .as_deref()
                .expect("chart rels bytes should be present"),
        )
        .expect("chart rels should be valid utf-8");
        assert!(
            chart_rels_xml.contains("schemas.microsoft.com/office/2014/relationships/chartEx"),
            "expected chart1.xml.rels to reference a ChartEx relationship for {fixture_name}.xlsx"
        );
        assert!(
            chart_rels_xml.contains("Target=\"chartEx1.xml\""),
            "expected chart1.xml.rels to target chartEx1.xml for {fixture_name}.xlsx"
        );

        let chart_ex = chart
            .parts
            .chart_ex
            .as_ref()
            .expect("chart should have a chartEx part");
        assert!(!chart_ex.bytes.is_empty());
        assert!(
            chart_ex.rels_path.is_some(),
            "fixture should include chartEx rels"
        );
        let chart_ex_rels_path = chart_ex.rels_path.as_deref().unwrap();
        assert_eq!(
            chart_ex.rels_bytes.as_deref(),
            pkg.part(chart_ex_rels_path),
            "chartEx rels bytes should be embedded on the extracted OpcPart for {fixture_name}.xlsx"
        );

        let model = chart.model.as_ref().expect("chart model present");
        match &model.chart_kind {
            ChartKind::Unknown { name } => {
                let kind = name
                    .strip_prefix("ChartEx:")
                    .expect("chart kind should be prefixed with ChartEx:");
                assert!(
                    !kind.trim().is_empty(),
                    "ChartEx kind should be non-empty for {fixture_name}.xlsx"
                );
                assert_eq!(
                    kind, expected_kind,
                    "ChartEx kind mismatch for {fixture_name}.xlsx"
                );
            }
            other => panic!("expected ChartKind::Unknown for ChartEx, got {other:?}"),
        }

        assert!(
            !model.series.is_empty(),
            "expected ChartEx model to include at least one series for {fixture_name}.xlsx"
        );
        let series = &model.series[0];
        assert_eq!(
            series.idx,
            Some(0),
            "expected series idx to be parsed for {fixture_name}.xlsx"
        );
        assert_eq!(
            series.order,
            Some(0),
            "expected series order to be parsed for {fixture_name}.xlsx"
        );
        assert_eq!(
            series
                .name
                .as_ref()
                .map(|name| name.rich_text.plain_text()),
            Some("Value"),
            "expected series name cache to be present for {fixture_name}.xlsx"
        );

        let categories = series
            .categories
            .as_ref()
            .expect("expected series categories parsed from ChartEx caches");
        assert!(
            categories
                .formula
                .as_deref()
                .is_some_and(|f| !f.trim().is_empty()),
            "expected categories formula to be present for {fixture_name}.xlsx"
        );
        assert!(
            categories.cache.as_ref().is_some_and(|c| !c.is_empty()),
            "expected categories cache to be present for {fixture_name}.xlsx"
        );
        assert_eq!(
            categories.formula.as_deref(),
            Some("Sheet1!$A$2:$A$4"),
            "expected categories formula to match fixture for {fixture_name}.xlsx"
        );
        assert_eq!(
            categories.cache.as_deref(),
            Some(&["A".to_string(), "B".to_string(), "C".to_string()][..]),
            "expected categories cache to match fixture for {fixture_name}.xlsx"
        );

        let values = series
            .values
            .as_ref()
            .expect("expected series values parsed from ChartEx caches");
        assert!(
            values
                .formula
                .as_deref()
                .is_some_and(|f| !f.trim().is_empty()),
            "expected values formula to be present for {fixture_name}.xlsx"
        );
        assert!(
            values.cache.as_ref().is_some_and(|c| !c.is_empty()),
            "expected values cache to be present for {fixture_name}.xlsx"
        );
        assert_eq!(
            values.formula.as_deref(),
            Some("Sheet1!$B$2:$B$4"),
            "expected values formula to match fixture for {fixture_name}.xlsx"
        );
        assert_eq!(
            values.format_code.as_deref(),
            Some("General"),
            "expected values formatCode to match fixture for {fixture_name}.xlsx"
        );
        assert_eq!(
            values.cache.as_deref(),
            Some(&[10.0, 20.0, 30.0][..]),
            "expected values cache to match fixture for {fixture_name}.xlsx"
        );

        // Ensure round-tripping preserves ChartEx parts byte-for-byte.
        let roundtrip = pkg.write_to_bytes().expect("round-trip write");
        let pkg2 = XlsxPackage::from_bytes(&roundtrip).expect("parse round-trip package");
        assert_eq!(
            pkg.part(&chart_ex.path),
            pkg2.part(&chart_ex.path),
            "ChartEx xml should round-trip losslessly for {fixture_name}.xlsx"
        );
        assert_eq!(
            pkg.part(chart_ex.rels_path.as_deref().unwrap()),
            pkg2.part(chart_ex.rels_path.as_deref().unwrap()),
            "ChartEx rels should round-trip losslessly for {fixture_name}.xlsx"
        );

        assert_eq!(
            model.title,
            Some(TextModel::plain(expected_title)),
            "expected {fixture_name}.xlsx ChartEx title to be parsed"
        );
    }
}
