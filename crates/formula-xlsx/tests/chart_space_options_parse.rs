use formula_xlsx::drawingml::charts::parse_chart_space;

#[test]
fn parses_chart_space_level_options_into_model() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:style val="7"/>
  <c:roundedCorners val="1"/>
  <c:chart>
    <c:plotVisOnly val="0"/>
    <c:dispBlanksAs val="span"/>
    <c:plotArea>
      <c:barChart/>
    </c:plotArea>
  </c:chart>
</c:chartSpace>
"#;

    let model = parse_chart_space(xml.as_bytes(), "chart.xml").expect("parse chartSpace XML");

    assert_eq!(model.style_id, Some(7));
    assert_eq!(model.rounded_corners, Some(true));
    assert_eq!(model.disp_blanks_as.as_deref(), Some("span"));
    assert_eq!(model.plot_vis_only, Some(false));
}

