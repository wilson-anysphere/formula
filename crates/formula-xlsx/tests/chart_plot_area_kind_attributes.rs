use formula_model::charts::PlotAreaModel;
use formula_xlsx::drawingml::charts::parse_chart_space;

#[test]
fn parses_bar_chart_gap_width_overlap_and_vary_colors() {
    let xml = r#"
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart>
            <c:plotArea>
              <c:barChart>
                <c:varyColors val="1"/>
                <c:gapWidth val="200"/>
                <c:overlap val="-50"/>
              </c:barChart>
            </c:plotArea>
          </c:chart>
        </c:chartSpace>
    "#;

    let model = parse_chart_space(xml.as_bytes(), "chart1.xml").expect("parse chartSpace");
    match model.plot_area {
        PlotAreaModel::Bar(bar) => {
            assert_eq!(bar.vary_colors, Some(true));
            assert_eq!(bar.gap_width, Some(200));
            assert_eq!(bar.overlap, Some(-50));
        }
        other => panic!("expected bar plot area, got {other:?}"),
    }
}

#[test]
fn parses_line_chart_series_smooth() {
    let xml = r#"
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart>
            <c:plotArea>
              <c:lineChart>
                <c:ser>
                  <c:smooth val="1"/>
                </c:ser>
              </c:lineChart>
            </c:plotArea>
          </c:chart>
        </c:chartSpace>
    "#;

    let model = parse_chart_space(xml.as_bytes(), "chart1.xml").expect("parse chartSpace");
    assert_eq!(model.series.len(), 1);
    assert_eq!(model.series[0].smooth, Some(true));
}

#[test]
fn parses_bar_chart_series_invert_if_negative() {
    let xml = r#"
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart>
            <c:plotArea>
              <c:barChart>
                <c:ser>
                  <c:invertIfNegative val="1"/>
                </c:ser>
              </c:barChart>
            </c:plotArea>
          </c:chart>
        </c:chartSpace>
    "#;

    let model = parse_chart_space(xml.as_bytes(), "chart1.xml").expect("parse chartSpace");
    assert_eq!(model.series.len(), 1);
    assert_eq!(model.series[0].invert_if_negative, Some(true));
}
