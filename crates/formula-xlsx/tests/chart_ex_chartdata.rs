use formula_model::charts::{ChartKind, SeriesData};
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

#[test]
fn chart_ex_fills_missing_formulas_from_chartdata_and_supports_nested_nodes() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">
  <cx:chartData>
    <cx:dataSet>
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
    </cx:dataSet>
  </cx:chartData>

  <cx:chart>
    <cx:plotArea>
      <cx:scatterChart>
        <cx:ser>
          <cx:layoutPr>
            <cx:dataId val="0"/>
          </cx:layoutPr>
          <cx:cat>
            <cx:strRef>
              <cx:strCache>
                <cx:ptCount val="3"/>
                <cx:pt idx="0"><cx:v>A</cx:v></cx:pt>
                <cx:pt idx="1"><cx:v>B</cx:v></cx:pt>
                <cx:pt idx="2"><cx:v>C</cx:v></cx:pt>
              </cx:strCache>
            </cx:strRef>
          </cx:cat>
          <cx:val>
            <cx:numRef>
              <cx:numCache>
                <cx:ptCount val="3"/>
                <cx:pt idx="0"><cx:v>10</cx:v></cx:pt>
                <cx:pt idx="1"><cx:v>20</cx:v></cx:pt>
                <cx:pt idx="2"><cx:v>30</cx:v></cx:pt>
              </cx:numCache>
            </cx:numRef>
          </cx:val>
          <cx:xVal>
            <cx:numRef>
              <cx:numCache>
                <cx:ptCount val="3"/>
                <cx:pt idx="0"><cx:v>1</cx:v></cx:pt>
                <cx:pt idx="1"><cx:v>2</cx:v></cx:pt>
                <cx:pt idx="2"><cx:v>3</cx:v></cx:pt>
              </cx:numCache>
            </cx:numRef>
          </cx:xVal>
          <cx:yVal>
            <cx:numRef>
              <cx:numCache>
                <cx:ptCount val="3"/>
                <cx:pt idx="0"><cx:v>4</cx:v></cx:pt>
                <cx:pt idx="1"><cx:v>5</cx:v></cx:pt>
                <cx:pt idx="2"><cx:v>6</cx:v></cx:pt>
              </cx:numCache>
            </cx:numRef>
          </cx:yVal>
        </cx:ser>
      </cx:scatterChart>
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
        series
            .categories
            .as_ref()
            .and_then(|d| d.cache.as_ref())
            .and_then(|v| v.first())
            .map(String::as_str),
        Some("A")
    );

    assert_eq!(
        series.values.as_ref().and_then(|d| d.formula.as_deref()),
        Some("Sheet1!$B$2:$B$4")
    );
    assert_eq!(
        series
            .values
            .as_ref()
            .and_then(|d| d.cache.as_ref())
            .and_then(|v| v.first())
            .copied(),
        Some(10.0)
    );

    match series.x_values.as_ref() {
        Some(SeriesData::Number(d)) => {
            assert_eq!(d.formula.as_deref(), Some("Sheet1!$C$2:$C$4"));
            assert_eq!(d.cache.as_ref().and_then(|v| v.first()).copied(), Some(1.0));
        }
        other => panic!("expected x_values numeric series data, got {other:?}"),
    }
    match series.y_values.as_ref() {
        Some(SeriesData::Number(d)) => {
            assert_eq!(d.formula.as_deref(), Some("Sheet1!$D$2:$D$4"));
            assert_eq!(d.cache.as_ref().and_then(|v| v.first()).copied(), Some(4.0));
        }
        other => panic!("expected y_values numeric series data, got {other:?}"),
    }
}

#[test]
fn chart_ex_collects_series_from_multiple_chart_type_nodes() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">
  <cx:chartData>
    <cx:data id="0">
      <cx:numDim type="val">
        <cx:f>Sheet1!$A$2:$A$4</cx:f>
      </cx:numDim>
    </cx:data>
    <cx:data id="1">
      <cx:numDim type="val">
        <cx:f>Sheet1!$B$2:$B$4</cx:f>
      </cx:numDim>
    </cx:data>
  </cx:chartData>

  <cx:chart>
    <cx:plotArea>
      <cx:histogramChart>
        <cx:ser dataId="0" />
      </cx:histogramChart>
      <cx:waterfallChart>
        <cx:ser dataId="1" />
      </cx:waterfallChart>
    </cx:plotArea>
  </cx:chart>
</cx:chartSpace>
"#;

    let model = parse_chart_ex(xml, "chartEx1.xml").expect("parse chartEx");
    assert_eq!(model.series.len(), 2);
    assert_eq!(
        model.series[0]
            .values
            .as_ref()
            .and_then(|d| d.formula.as_deref()),
        Some("Sheet1!$A$2:$A$4")
    );
    assert_eq!(
        model.series[1]
            .values
            .as_ref()
            .and_then(|d| d.formula.as_deref()),
        Some("Sheet1!$B$2:$B$4")
    );
}

#[test]
fn chart_ex_parses_series_without_explicit_chart_type_node() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">
  <cx:chartData>
    <cx:data id="0">
      <cx:numDim type="val">
        <cx:f>Sheet1!$A$2:$A$4</cx:f>
      </cx:numDim>
    </cx:data>
  </cx:chartData>

  <cx:chart>
    <cx:plotArea>
      <cx:series layoutId="treemap" dataId="0" />
    </cx:plotArea>
  </cx:chart>
</cx:chartSpace>
"#;

    let model = parse_chart_ex(xml, "chartEx1.xml").expect("parse chartEx");

    match &model.chart_kind {
        ChartKind::Unknown { name } => assert_eq!(name, "ChartEx:treemap"),
        other => panic!("expected ChartKind::Unknown, got {other:?}"),
    }

    assert_eq!(model.series.len(), 1);
    assert_eq!(
        model.series[0]
            .values
            .as_ref()
            .and_then(|d| d.formula.as_deref()),
        Some("Sheet1!$A$2:$A$4")
    );
}

#[test]
fn chart_ex_parses_chartdata_caches_when_series_omits_inline_dims() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">
  <cx:chartData>
    <cx:data id="0">
      <cx:strDim type="cat">
        <cx:f>Sheet1!$A$2:$A$4</cx:f>
        <cx:strCache>
          <cx:ptCount val="3"/>
          <cx:pt idx="0"><cx:v>A</cx:v></cx:pt>
          <cx:pt idx="1"><cx:v>B</cx:v></cx:pt>
          <cx:pt idx="2"><cx:v>C</cx:v></cx:pt>
        </cx:strCache>
      </cx:strDim>
      <cx:numDim type="val">
        <cx:f>Sheet1!$B$2:$B$4</cx:f>
        <cx:numCache>
          <cx:ptCount val="3"/>
          <cx:pt idx="0"><cx:v>10</cx:v></cx:pt>
          <cx:pt idx="1"><cx:v>20</cx:v></cx:pt>
          <cx:pt idx="2"><cx:v>30</cx:v></cx:pt>
        </cx:numCache>
      </cx:numDim>
    </cx:data>
  </cx:chartData>

  <cx:chart>
    <cx:plotArea>
      <cx:histogramChart>
        <cx:ser dataId="0" />
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
        series
            .categories
            .as_ref()
            .and_then(|d| d.cache.as_ref())
            .cloned(),
        Some(vec!["A".to_string(), "B".to_string(), "C".to_string()])
    );

    assert_eq!(
        series.values.as_ref().and_then(|d| d.formula.as_deref()),
        Some("Sheet1!$B$2:$B$4")
    );
    assert_eq!(
        series
            .values
            .as_ref()
            .and_then(|d| d.cache.as_ref())
            .cloned(),
        Some(vec![10.0, 20.0, 30.0])
    );
}
