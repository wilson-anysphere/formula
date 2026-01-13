use formula_xlsx::drawingml::charts::parse_chart_space;

#[test]
fn chart_space_captures_plot_area_ext_lst_xml() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:extLst>
        <c:ext uri="{12345678-1234-1234-1234-1234567890AB}">
          <c15:dummy xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart">
            <c15:payload>hello</c15:payload>
          </c15:dummy>
        </c:ext>
      </c:extLst>
      <c:barChart/>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

    let model = parse_chart_space(xml, "in-memory-chart.xml").expect("parse chartSpace");
    let ext_lst_xml = model
        .plot_area_ext_lst_xml
        .as_ref()
        .expect("plotArea extLst should be captured");
    assert!(!ext_lst_xml.trim().is_empty());
    assert!(ext_lst_xml.contains("<c:extLst"));
    assert!(ext_lst_xml.contains("uri=\"{12345678-1234-1234-1234-1234567890AB}\""));
}
