use formula_model::charts::ChartKind;
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
    for fixture_name in [
        "waterfall",
        "histogram",
        "treemap",
        "sunburst",
        "funnel",
        "box-whisker",
        "pareto",
        "map",
    ] {
        let bytes = load_fixture(fixture_name);
        let pkg = XlsxPackage::from_bytes(&bytes).expect("parse package");

        let charts = pkg.extract_chart_objects().expect("extract chart objects");
        assert!(
            !charts.is_empty(),
            "expected at least one chart object in {fixture_name}.xlsx"
        );

        let chart = &charts[0];
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
        let chart_ex_rels_bytes = pkg
            .part(chart_ex_rels_path)
            .expect("fixture should include chartEx rels bytes");
        assert!(
            !chart_ex_rels_bytes.is_empty(),
            "chartEx rels bytes should be non-empty for {fixture_name}.xlsx"
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
            }
            other => panic!("expected ChartKind::Unknown for ChartEx, got {other:?}"),
        }

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
    }
}
