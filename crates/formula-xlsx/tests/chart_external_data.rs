use formula_model::charts::ChartKind;
use formula_xlsx::drawingml::charts::parse_chart_space;

#[test]
fn parses_chart_space_external_data_link() {
    let xml = r#"<c:chartSpace
        xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <c:externalData r:id="rId42">
        <c:autoUpdate val="1"/>
      </c:externalData>
      <c:chart>
        <c:plotArea>
          <c:barChart/>
        </c:plotArea>
      </c:chart>
    </c:chartSpace>"#;

    let model = parse_chart_space(xml.as_bytes(), "in-memory.xml").expect("parse chartSpace");
    assert_eq!(model.chart_kind, ChartKind::Bar);
    assert_eq!(model.external_data_rel_id.as_deref(), Some("rId42"));
    assert_eq!(model.external_data_auto_update, Some(true));
}
