use formula_model::charts::{
    AxisKind, AxisPosition, ChartKind, LegendPosition, PlotAreaModel, SeriesData,
};
use formula_xlsx::drawingml::charts::parse_chart_space;
use formula_xlsx::XlsxPackage;

const FIXTURE_BAR: &[u8] = include_bytes!("../../../fixtures/charts/xlsx/bar.xlsx");
const FIXTURE_LINE: &[u8] = include_bytes!("../../../fixtures/charts/xlsx/line.xlsx");
const FIXTURE_PIE: &[u8] = include_bytes!("../../../fixtures/charts/xlsx/pie.xlsx");
const FIXTURE_SCATTER: &[u8] = include_bytes!("../../../fixtures/charts/xlsx/scatter.xlsx");
const FIXTURE_BASIC_CHART: &[u8] = include_bytes!("../../../fixtures/charts/xlsx/basic-chart.xlsx");
const FIXTURE_COMBO_BAR_LINE: &[u8] =
    include_bytes!("../../../fixtures/charts/xlsx/combo-bar-line.xlsx");

fn parse_fixture(bytes: &[u8]) -> formula_model::charts::ChartModel {
    let pkg = XlsxPackage::from_bytes(bytes).expect("open xlsx fixture");
    let chart_xml = pkg
        .part("xl/charts/chart1.xml")
        .expect("fixture contains xl/charts/chart1.xml");
    parse_chart_space(chart_xml, "xl/charts/chart1.xml").expect("parse chartSpace")
}

#[test]
fn parses_generated_chart_fixtures() {
    for (name, fixture, expected_kind) in [
        ("bar.xlsx", FIXTURE_BAR, ChartKind::Bar),
        ("line.xlsx", FIXTURE_LINE, ChartKind::Line),
        ("pie.xlsx", FIXTURE_PIE, ChartKind::Pie),
        ("scatter.xlsx", FIXTURE_SCATTER, ChartKind::Scatter),
    ] {
        let model = parse_fixture(fixture);
        assert_eq!(model.chart_kind, expected_kind, "fixture {name}");

        let title = model
            .title
            .as_ref()
            .map(|t| t.rich_text.text.as_str())
            .unwrap_or("");
        assert_eq!(title, "Example Chart", "fixture {name}");

        let legend = model.legend.as_ref().expect("legend present");
        assert_eq!(legend.position, LegendPosition::Right, "fixture {name}");
        assert!(!legend.overlay, "fixture {name}");

        assert_eq!(model.series.len(), 1, "fixture {name}");

        match expected_kind {
            ChartKind::Scatter => {
                assert_eq!(model.axes.len(), 2, "fixture {name}");
                assert!(model.axes.iter().all(|a| a.kind == AxisKind::Value));
                assert!(model
                    .axes
                    .iter()
                    .any(|a| a.position == Some(AxisPosition::Bottom)));
                assert!(model
                    .axes
                    .iter()
                    .any(|a| a.position == Some(AxisPosition::Left)));

                let ser = &model.series[0];
                match ser.x_values.as_ref().expect("x_values present") {
                    SeriesData::Text(data) => {
                        assert_eq!(data.formula.as_deref(), Some("Sheet1!$A$2:$A$5"));
                        assert_eq!(data.cache.as_ref().map(Vec::len), Some(4));
                    }
                    other => panic!("expected scatter x_values to be text, got {other:?}"),
                }
                match ser.y_values.as_ref().expect("y_values present") {
                    SeriesData::Number(data) => {
                        assert_eq!(data.formula.as_deref(), Some("Sheet1!$B$2:$B$5"));
                        assert_eq!(data.cache.as_ref().map(Vec::len), Some(4));
                        assert_eq!(data.format_code.as_deref(), Some("General"));
                    }
                    other => panic!("expected scatter y_values to be number, got {other:?}"),
                }
            }
            ChartKind::Bar | ChartKind::Line => {
                assert_eq!(model.axes.len(), 2, "fixture {name}");
                let cat = model
                    .axes
                    .iter()
                    .find(|a| a.kind == AxisKind::Category)
                    .expect("category axis present");
                assert_eq!(cat.position, Some(AxisPosition::Bottom));

                let val = model
                    .axes
                    .iter()
                    .find(|a| a.kind == AxisKind::Value)
                    .expect("value axis present");
                assert_eq!(val.position, Some(AxisPosition::Left));
                assert!(val.major_gridlines);
                assert_eq!(
                    val.num_fmt.as_ref().map(|f| f.format_code.as_str()),
                    Some("General")
                );

                let ser = &model.series[0];
                let cats = ser.categories.as_ref().expect("categories present");
                assert_eq!(cats.formula.as_deref(), Some("Sheet1!$A$2:$A$5"));
                assert_eq!(cats.cache.as_ref().map(Vec::len), Some(4));
                let vals = ser.values.as_ref().expect("values present");
                assert_eq!(vals.formula.as_deref(), Some("Sheet1!$B$2:$B$5"));
                assert_eq!(vals.cache.as_ref().map(Vec::len), Some(4));
                assert_eq!(vals.format_code.as_deref(), Some("General"));
            }
            ChartKind::Pie => {
                assert!(model.axes.is_empty(), "pie charts should not have axes");
                let ser = &model.series[0];
                let cats = ser.categories.as_ref().expect("categories present");
                assert_eq!(cats.formula.as_deref(), Some("Sheet1!$A$2:$A$5"));
                assert_eq!(cats.cache.as_ref().map(Vec::len), Some(4));
                let vals = ser.values.as_ref().expect("values present");
                assert_eq!(vals.formula.as_deref(), Some("Sheet1!$B$2:$B$5"));
                assert_eq!(vals.cache.as_ref().map(Vec::len), Some(4));
            }
            ChartKind::Unknown { .. } => unreachable!("fixture should not be unknown"),
        }
    }
}

#[test]
fn parses_basic_chart_fixture_without_caches() {
    let model = parse_fixture(FIXTURE_BASIC_CHART);
    assert_eq!(model.chart_kind, ChartKind::Pie);
    assert!(model.title.is_none());
    assert!(model.legend.is_none());
    assert_eq!(model.series.len(), 1);

    let ser = &model.series[0];
    assert_eq!(
        ser.name.as_ref().and_then(|t| t.formula.as_deref()),
        Some("Sheet1!$B$1")
    );
    assert!(ser
        .categories
        .as_ref()
        .and_then(|c| c.cache.as_ref())
        .is_none());
    assert!(ser.values.as_ref().and_then(|c| c.cache.as_ref()).is_none());
}

#[test]
fn parses_combo_bar_line_fixture() {
    let model = parse_fixture(FIXTURE_COMBO_BAR_LINE);
    assert_eq!(model.chart_kind, ChartKind::Bar);
    assert_eq!(model.series.len(), 2);

    let PlotAreaModel::Combo(combo) = &model.plot_area else {
        panic!("expected combo plot area, got {:?}", model.plot_area);
    };
    assert_eq!(combo.charts.len(), 2);

    // Series should be tagged with their owning subplot index.
    assert_eq!(model.series[0].plot_index, Some(0));
    assert_eq!(model.series[1].plot_index, Some(1));
}
