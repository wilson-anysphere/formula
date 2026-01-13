use formula_model::charts::LineDash;
use formula_model::charts::FillStyle;
use formula_model::Color;
use formula_xlsx::drawingml::charts::parse_chart_space;
use formula_xlsx::XlsxPackage;

const FIXTURE: &[u8] = include_bytes!("fixtures/chart_point_overrides.xlsx");

#[test]
fn extracts_series_and_point_formatting_overrides() {
    let package = XlsxPackage::from_bytes(FIXTURE).unwrap();
    let chart_xml = package.part("xl/charts/chart1.xml").expect("chart1.xml exists");
    let model = parse_chart_space(chart_xml, "xl/charts/chart1.xml").unwrap();
    assert_eq!(model.series.len(), 1);

    let series = &model.series[0];

    let series_style = series.style.as_ref().expect("series has spPr");
    let series_fill = series_style.fill.as_ref().expect("series has fill");
    let FillStyle::Solid(series_fill) = series_fill else {
        panic!("expected series fill to be solidFill, got {series_fill:?}");
    };
    assert_eq!(
        series_fill.color,
        Color::Theme {
            theme: 4,
            tint: None
        }
    );

    let series_line = series_style.line.as_ref().expect("series has ln");
    assert_eq!(series_line.width_100pt, Some(100));
    assert_eq!(series_line.dash, Some(LineDash::Dash));

    let pt = series
        .points
        .iter()
        .find(|p| p.idx == 1)
        .expect("point idx=1 override exists");
    let pt_style = pt.style.as_ref().expect("point has spPr");
    let pt_fill = pt_style.fill.as_ref().expect("point has fill");
    let FillStyle::Solid(pt_fill) = pt_fill else {
        panic!("expected point fill to be solidFill, got {pt_fill:?}");
    };
    assert_eq!(pt_fill.color, Color::Argb(0xFFFF0000));
}
