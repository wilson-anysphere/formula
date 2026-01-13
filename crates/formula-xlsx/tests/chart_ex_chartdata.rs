use formula_model::charts::SeriesData;
use formula_xlsx::drawingml::charts::parse_chart_ex;

#[test]
fn chart_ex_parses_formulas_from_chartdata_by_data_id() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">
  <cx:chartData>
    <cx:data id="0">
      <cx:strDim type="cat">
        <cx:f>Sheet1!$A$2:$A$4</cx:f>
      </cx:strDim>
      <cx:numDim type="val">
        <cx:f>Sheet1!$B$2:$B$4</cx:f>
      </cx:numDim>
      <cx:numDim type="x">
        <cx:f>Sheet1!$C$2:$C$4</cx:f>
      </cx:numDim>
      <cx:numDim type="y">
        <cx:f>Sheet1!$D$2:$D$4</cx:f>
      </cx:numDim>
    </cx:data>
  </cx:chartData>

  <cx:chart>
    <cx:plotArea>
      <cx:histogramChart>
        <cx:ser dataId="0">
          <cx:tx>
            <cx:strRef>
              <cx:f>Sheet1!$B$1</cx:f>
            </cx:strRef>
          </cx:tx>
        </cx:ser>
      </cx:histogramChart>
    </cx:plotArea>
  </cx:chart>
</cx:chartSpace>
"#;

    let model = parse_chart_ex(xml, "chartEx1.xml").expect("parse chartEx");
    assert_eq!(model.series.len(), 1);
    let series = &model.series[0];

    assert_eq!(
        series
            .categories
            .as_ref()
            .and_then(|d| d.formula.as_deref()),
        Some("Sheet1!$A$2:$A$4")
    );
    assert_eq!(
        series.values.as_ref().and_then(|d| d.formula.as_deref()),
        Some("Sheet1!$B$2:$B$4")
    );

    match series.x_values.as_ref() {
        Some(SeriesData::Number(d)) => assert_eq!(d.formula.as_deref(), Some("Sheet1!$C$2:$C$4")),
        other => panic!("expected x_values numeric series data, got {other:?}"),
    }
    match series.y_values.as_ref() {
        Some(SeriesData::Number(d)) => assert_eq!(d.formula.as_deref(), Some("Sheet1!$D$2:$D$4")),
        other => panic!("expected y_values numeric series data, got {other:?}"),
    }
}

#[test]
fn chart_ex_uses_size_dim_as_values_when_no_val_dim_is_present() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">
  <cx:chartData>
    <cx:data id="1">
      <cx:numDim type="x">
        <cx:f>Sheet1!$A$2:$A$4</cx:f>
      </cx:numDim>
      <cx:numDim type="y">
        <cx:f>Sheet1!$B$2:$B$4</cx:f>
      </cx:numDim>
      <cx:numDim type="size">
        <cx:f>Sheet1!$C$2:$C$4</cx:f>
      </cx:numDim>
    </cx:data>
  </cx:chartData>

  <cx:chart>
    <cx:plotArea>
      <cx:bubbleChart>
        <cx:ser dataId="1" />
      </cx:bubbleChart>
    </cx:plotArea>
  </cx:chart>
</cx:chartSpace>
"#;

    let model = parse_chart_ex(xml, "chartEx1.xml").expect("parse chartEx");
    assert_eq!(model.series.len(), 1);
    let series = &model.series[0];

    // `SeriesModel` doesn't currently have a dedicated bubble size slot; we
    // preserve the formula on `values` so callers can at least discover it.
    assert_eq!(
        series.values.as_ref().and_then(|d| d.formula.as_deref()),
        Some("Sheet1!$C$2:$C$4")
    );

    match series.x_values.as_ref() {
        Some(SeriesData::Number(d)) => assert_eq!(d.formula.as_deref(), Some("Sheet1!$A$2:$A$4")),
        other => panic!("expected x_values numeric series data, got {other:?}"),
    }
    match series.y_values.as_ref() {
        Some(SeriesData::Number(d)) => assert_eq!(d.formula.as_deref(), Some("Sheet1!$B$2:$B$4")),
        other => panic!("expected y_values numeric series data, got {other:?}"),
    }
}
