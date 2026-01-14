use formula_model::charts::{ChartKind, ComboChartEntry, PlotAreaModel};
use formula_xlsx::drawingml::charts::parse_chart_space;

#[test]
fn parse_chart_includes_series_from_all_chart_types_in_combo_plot_area() {
    let xml = r#"
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <c:chart>
            <c:title>
              <c:tx>
                <c:rich>
                  <a:t>Combo</a:t>
                </c:rich>
              </c:tx>
            </c:title>
            <c:plotArea>
              <c:barChart>
                <c:ser>
                  <c:tx><c:v>Bar Series</c:v></c:tx>
                  <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$3</c:f></c:strRef></c:cat>
                  <c:val><c:numRef><c:f>Sheet1!$B$2:$B$3</c:f></c:numRef></c:val>
                </c:ser>
              </c:barChart>
              <c:lineChart>
                <c:ser>
                  <c:tx><c:v>Line Series</c:v></c:tx>
                  <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$3</c:f></c:strRef></c:cat>
                  <c:val><c:numRef><c:f>Sheet1!$C$2:$C$3</c:f></c:numRef></c:val>
                </c:ser>
              </c:lineChart>
            </c:plotArea>
          </c:chart>
        </c:chartSpace>
    "#;

    let model = parse_chart_space(xml.as_bytes(), "combo.xml").expect("parse chart space");

    // Combo charts keep the first chart kind as the "primary" kind.
    assert_eq!(model.chart_kind, ChartKind::Bar);
    let PlotAreaModel::Combo(combo) = &model.plot_area else {
        panic!("expected combo plot area, got {:?}", model.plot_area);
    };
    assert_eq!(combo.charts.len(), 2);
    match &combo.charts[0] {
        ComboChartEntry::Bar { series, .. } => {
            assert_eq!(series.start, 0);
            assert_eq!(series.end, 1);
        }
        other => panic!("expected first combo chart entry to be Bar, got {other:?}"),
    }
    match &combo.charts[1] {
        ComboChartEntry::Line { series, .. } => {
            assert_eq!(series.start, 1);
            assert_eq!(series.end, 2);
        }
        other => panic!("expected second combo chart entry to be Line, got {other:?}"),
    }
    assert_eq!(
        model.title.as_ref().map(|t| t.rich_text.text.as_str()),
        Some("Combo")
    );

    assert_eq!(model.series.len(), 2);

    assert_eq!(
        model.series[0]
            .name
            .as_ref()
            .map(|t| t.rich_text.text.as_str()),
        Some("Bar Series")
    );
    assert_eq!(
        model.series[0]
            .categories
            .as_ref()
            .and_then(|c| c.formula.as_deref()),
        Some("Sheet1!$A$2:$A$3")
    );
    assert_eq!(
        model.series[0]
            .values
            .as_ref()
            .and_then(|v| v.formula.as_deref()),
        Some("Sheet1!$B$2:$B$3")
    );
    assert_eq!(model.series[0].plot_index, Some(0));

    assert_eq!(
        model.series[1]
            .name
            .as_ref()
            .map(|t| t.rich_text.text.as_str()),
        Some("Line Series")
    );
    assert_eq!(
        model.series[1]
            .categories
            .as_ref()
            .and_then(|c| c.formula.as_deref()),
        Some("Sheet1!$A$2:$A$3")
    );
    assert_eq!(
        model.series[1]
            .values
            .as_ref()
            .and_then(|v| v.formula.as_deref()),
        Some("Sheet1!$C$2:$C$3")
    );
    assert_eq!(model.series[1].plot_index, Some(1));
}
