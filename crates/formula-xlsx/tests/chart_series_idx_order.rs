use formula_xlsx::drawingml::charts::{parse_chart_ex, parse_chart_space};

#[test]
fn parses_series_idx_and_order_from_chart_space() {
    let xml = r#"
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart>
            <c:plotArea>
              <c:lineChart>
                <c:ser>
                  <c:idx val="2"/>
                  <c:order val="3"/>
                </c:ser>
              </c:lineChart>
            </c:plotArea>
          </c:chart>
        </c:chartSpace>
    "#;

    let model = parse_chart_space(xml.as_bytes(), "chart1.xml").expect("parse chartSpace");
    assert_eq!(model.series.len(), 1);
    assert_eq!(model.series[0].idx, Some(2));
    assert_eq!(model.series[0].order, Some(3));
}

#[test]
fn parses_series_idx_and_order_from_chart_ex() {
    let xml = r#"
        <cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">
          <cx:plotArea>
            <cx:histogramChart>
              <cx:ser>
                <cx:idx val="4"/>
                <cx:order val="5"/>
              </cx:ser>
            </cx:histogramChart>
          </cx:plotArea>
        </cx:chartSpace>
    "#;

    let model = parse_chart_ex(xml.as_bytes(), "chartEx1.xml").expect("parse chartEx");
    assert_eq!(model.series.len(), 1);
    assert_eq!(model.series[0].idx, Some(4));
    assert_eq!(model.series[0].order, Some(5));
}

