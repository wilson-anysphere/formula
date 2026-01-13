use formula_xlsx::drawingml::charts::{parse_chart_ex, parse_chart_space};
use formula_xlsx::XlsxPackage;

const FIXTURE_WATERFALL_CHARTS: &[u8] =
    include_bytes!("../../../fixtures/charts/xlsx/waterfall.xlsx");
const FIXTURE_WATERFALL_CHARTS_EX: &[u8] =
    include_bytes!("../../../fixtures/xlsx/charts-ex/waterfall.xlsx");

#[test]
fn merges_chart_space_series_when_chart_ex_is_minimal() {
    let pkg = XlsxPackage::from_bytes(FIXTURE_WATERFALL_CHARTS).expect("parse package");
    let charts = pkg.extract_chart_objects().expect("extract chart objects");
    assert!(
        !charts.is_empty(),
        "expected at least one chart object in waterfall.xlsx"
    );

    let chart = &charts[0];
    let chart_ex = chart
        .parts
        .chart_ex
        .as_ref()
        .expect("fixture should include a chartEx part");

    let chart_space_model =
        parse_chart_space(&chart.parts.chart.bytes, &chart.parts.chart.path).expect("chartSpace");
    let chart_ex_model = parse_chart_ex(&chart_ex.bytes, &chart_ex.path).expect("chartEx");

    assert!(
        chart_ex_model.series.is_empty(),
        "fixture chartEx should be minimal and omit series"
    );
    assert!(
        !chart_space_model.series.is_empty(),
        "fixture chartSpace should include series"
    );

    let model = chart.model.as_ref().expect("merged chart model present");
    assert_eq!(
        model.series, chart_space_model.series,
        "series should fall back to chartSpace when ChartEx yields none"
    );
}

#[test]
fn prefers_chart_ex_series_when_available() {
    let pkg = XlsxPackage::from_bytes(FIXTURE_WATERFALL_CHARTS_EX).expect("parse package");
    let charts = pkg.extract_chart_objects().expect("extract chart objects");
    assert!(
        !charts.is_empty(),
        "expected at least one chart object in charts-ex/waterfall.xlsx"
    );

    let chart = &charts[0];
    let chart_ex = chart
        .parts
        .chart_ex
        .as_ref()
        .expect("fixture should include a chartEx part");

    let chart_space_model =
        parse_chart_space(&chart.parts.chart.bytes, &chart.parts.chart.path).expect("chartSpace");
    let chart_ex_model = parse_chart_ex(&chart_ex.bytes, &chart_ex.path).expect("chartEx");

    assert!(
        chart_space_model.series.is_empty(),
        "fixture chartSpace should be minimal and omit series"
    );
    assert!(
        !chart_ex_model.series.is_empty(),
        "fixture chartEx should include series"
    );

    let model = chart.model.as_ref().expect("merged chart model present");
    assert_eq!(
        model.series, chart_ex_model.series,
        "series should come from ChartEx when it is more complete"
    );
}
