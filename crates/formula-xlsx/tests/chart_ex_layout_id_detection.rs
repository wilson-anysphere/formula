use formula_model::charts::ChartKind;
use formula_xlsx::drawingml::charts::parse_chart_ex;

#[test]
fn detects_chart_ex_kind_from_series_layout_id_attribute() {
    // Some real-world ChartEx documents omit a `<*Chart>` node (e.g.
    // `<cx:treemapChart>`). Excel still includes enough hints to classify the
    // chart kind via `layoutId` on `<cx:series>`.
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">
  <cx:chart>
    <cx:plotArea>
      <cx:chartData>
        <cx:series layoutId="treemap"/>
      </cx:chartData>
    </cx:plotArea>
  </cx:chart>
</cx:chartSpace>
"#;

    let model = parse_chart_ex(xml.as_bytes(), "xl/charts/chartEx1.xml").expect("parse chartEx");

    match &model.chart_kind {
        ChartKind::Unknown { name } => assert_eq!(name, "ChartEx:treemap"),
        other => panic!("expected ChartKind::Unknown for ChartEx, got {other:?}"),
    }
}

#[test]
fn normalizes_layout_id_values_and_strips_optional_chart_suffix() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">
  <cx:chart>
    <cx:plotArea>
      <cx:chartData>
        <cx:series layoutId="TreemapChart"/>
      </cx:chartData>
    </cx:plotArea>
  </cx:chart>
</cx:chartSpace>
"#;

    let model = parse_chart_ex(xml.as_bytes(), "xl/charts/chartEx1.xml").expect("parse chartEx");
    match &model.chart_kind {
        ChartKind::Unknown { name } => assert_eq!(name, "ChartEx:treemap"),
        other => panic!("expected ChartKind::Unknown for ChartEx, got {other:?}"),
    }
}

#[test]
fn detects_chart_ex_kind_from_chart_type_attribute_when_no_chart_node_exists() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">
  <cx:chart chartType="sunburstChart">
    <cx:plotArea/>
  </cx:chart>
</cx:chartSpace>
"#;

    let model = parse_chart_ex(xml.as_bytes(), "xl/charts/chartEx1.xml").expect("parse chartEx");
    match &model.chart_kind {
        ChartKind::Unknown { name } => assert_eq!(name, "ChartEx:sunburst"),
        other => panic!("expected ChartKind::Unknown for ChartEx, got {other:?}"),
    }
}

#[test]
fn prefers_known_chart_type_nodes_even_if_other_chart_suffixed_nodes_appear_first() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">
  <cx:chart>
    <cx:plotArea>
      <cx:someStyleChart/>
      <cx:histogramChart/>
    </cx:plotArea>
  </cx:chart>
</cx:chartSpace>
"#;

    let model = parse_chart_ex(xml.as_bytes(), "xl/charts/chartEx1.xml").expect("parse chartEx");
    match &model.chart_kind {
        ChartKind::Unknown { name } => assert_eq!(name, "ChartEx:histogram"),
        other => panic!("expected ChartKind::Unknown for ChartEx, got {other:?}"),
    }
}
