use formula_xlsx::drawingml::charts::{parse_chart_ex, parse_chart_space};

#[test]
fn parses_series_idx_and_order_from_chart_space() {
    let xml = br#"
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart>
            <c:plotArea>
              <c:barChart>
                <c:ser>
                  <c:idx val="2"/>
                  <c:order val="3"/>
                </c:ser>
                <c:ser>
                  <c:idx val="4"/>
                  <c:order val="5"/>
                </c:ser>
              </c:barChart>
            </c:plotArea>
          </c:chart>
        </c:chartSpace>
    "#;

    let model = parse_chart_space(xml, "chart1.xml").expect("parse chartSpace");
    assert_eq!(model.series.len(), 2);
    assert_eq!(model.series[0].idx, Some(2));
    assert_eq!(model.series[0].order, Some(3));
    assert_eq!(model.series[1].idx, Some(4));
    assert_eq!(model.series[1].order, Some(5));
}

#[test]
fn warns_on_unparsable_series_idx() {
    let xml = br#"
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart>
            <c:plotArea>
              <c:barChart>
                <c:ser>
                  <c:idx val="nope"/>
                  <c:order val="1"/>
                </c:ser>
              </c:barChart>
            </c:plotArea>
          </c:chart>
        </c:chartSpace>
    "#;

    let model = parse_chart_space(xml, "chart1.xml").expect("parse chartSpace");
    assert_eq!(model.series.len(), 1);
    assert_eq!(model.series[0].idx, None);
    assert_eq!(model.series[0].order, Some(1));
    assert!(
        model.diagnostics.iter().any(|d| d.message.contains("series idx")),
        "expected warning about series idx parse failure, got {:?}",
        model.diagnostics
    );
}

#[test]
fn warns_on_unparsable_series_order() {
    let xml = br#"
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart>
            <c:plotArea>
              <c:barChart>
                <c:ser>
                  <c:idx val="1"/>
                  <c:order val="nope"/>
                </c:ser>
              </c:barChart>
            </c:plotArea>
          </c:chart>
        </c:chartSpace>
    "#;

    let model = parse_chart_space(xml, "chart1.xml").expect("parse chartSpace");
    assert_eq!(model.series.len(), 1);
    assert_eq!(model.series[0].idx, Some(1));
    assert_eq!(model.series[0].order, None);
    assert!(
        model.diagnostics.iter().any(|d| d.message.contains("series order")),
        "expected warning about series order parse failure, got {:?}",
        model.diagnostics
    );
}

#[test]
fn parses_series_idx_and_order_from_chart_ex() {
    let xml = br#"
        <cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">
          <cx:chart>
            <cx:plotArea>
              <cx:histogramChart>
                <cx:ser>
                  <cx:idx val="4"/>
                  <cx:order val="5"/>
                </cx:ser>
              </cx:histogramChart>
            </cx:plotArea>
          </cx:chart>
        </cx:chartSpace>
    "#;

    let model = parse_chart_ex(xml, "chartEx1.xml").expect("parse chartEx");
    assert_eq!(model.series.len(), 1);
    assert_eq!(model.series[0].idx, Some(4));
    assert_eq!(model.series[0].order, Some(5));
}

#[test]
fn chart_ex_missing_idx_order_defaults_to_position() {
    let xml = br#"
        <cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">
          <cx:chart>
            <cx:plotArea>
              <cx:regionMapChart>
                <cx:series>
                  <cx:dataId val="0"/>
                </cx:series>
              </cx:regionMapChart>
            </cx:plotArea>
          </cx:chart>
          <cx:chartData>
            <cx:data id="0">
              <cx:strDim type="cat">
                <cx:f>Sheet1!$A$2:$A$5</cx:f>
              </cx:strDim>
              <cx:numDim type="val">
                <cx:f>Sheet1!$B$2:$B$5</cx:f>
              </cx:numDim>
            </cx:data>
          </cx:chartData>
        </cx:chartSpace>
    "#;

    let model = parse_chart_ex(xml, "chartEx1.xml").expect("parse chartEx");
    assert_eq!(model.series.len(), 1);
    assert_eq!(model.series[0].idx, Some(0));
    assert_eq!(model.series[0].order, Some(0));
}

#[test]
fn parses_chart_ex_idx_order_from_series_attributes() {
    let xml = br#"
        <cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">
          <cx:chart>
            <cx:plotArea>
              <cx:histogramChart>
                <cx:ser idx="7" ORDER="8"/>
              </cx:histogramChart>
            </cx:plotArea>
          </cx:chart>
        </cx:chartSpace>
    "#;

    let model = parse_chart_ex(xml, "chartEx1.xml").expect("parse chartEx");
    assert_eq!(model.series.len(), 1);
    assert_eq!(model.series[0].idx, Some(7));
    assert_eq!(model.series[0].order, Some(8));
}

#[test]
fn parses_chart_ex_idx_order_from_text_content() {
    let xml = br#"
        <cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">
          <cx:chart>
            <cx:plotArea>
              <cx:histogramChart>
                <cx:ser>
                  <cx:idx>9</cx:idx>
                  <cx:order>10</cx:order>
                </cx:ser>
              </cx:histogramChart>
            </cx:plotArea>
          </cx:chart>
        </cx:chartSpace>
    "#;

    let model = parse_chart_ex(xml, "chartEx1.xml").expect("parse chartEx");
    assert_eq!(model.series.len(), 1);
    assert_eq!(model.series[0].idx, Some(9));
    assert_eq!(model.series[0].order, Some(10));
}

#[test]
fn warns_on_unparsable_chart_ex_idx() {
    let xml = br#"
        <cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">
          <cx:chart>
            <cx:plotArea>
              <cx:histogramChart>
                <cx:ser>
                  <cx:idx val="nope"/>
                  <cx:order val="5"/>
                </cx:ser>
              </cx:histogramChart>
            </cx:plotArea>
          </cx:chart>
        </cx:chartSpace>
    "#;

    let model = parse_chart_ex(xml, "chartEx1.xml").expect("parse chartEx");
    assert_eq!(model.series.len(), 1);
    // Unparsable idx should trigger a warning; missing idx is then defaulted to the series position.
    assert_eq!(model.series[0].idx, Some(0));
    assert_eq!(model.series[0].order, Some(5));
    assert!(
        model.diagnostics.iter().any(|d| d.message.contains("ChartEx series idx")),
        "expected warning about ChartEx series idx parse failure, got {:?}",
        model.diagnostics
    );
}
