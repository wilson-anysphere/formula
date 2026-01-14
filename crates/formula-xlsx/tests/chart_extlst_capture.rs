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

#[test]
fn chart_space_captures_ext_lst_xml_for_chart_space_chart_axis_and_series() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart">
  <c:extLst>
    <c:ext uri="{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}">
      <c15:dummy><c15:payload>space</c15:payload></c15:dummy>
    </c:ext>
  </c:extLst>
  <c:chart>
    <c:extLst>
      <c:ext uri="{BBBBBBBB-BBBB-BBBB-BBBB-BBBBBBBBBBBB}">
        <c15:dummy><c15:payload>chart</c15:payload></c15:dummy>
      </c:ext>
    </c:extLst>
    <c:plotArea>
      <c:extLst>
        <c:ext uri="{CCCCCCCC-CCCC-CCCC-CCCC-CCCCCCCCCCCC}">
          <c15:dummy><c15:payload>plot</c15:payload></c15:dummy>
        </c:ext>
      </c:extLst>
      <c:barChart>
        <c:ser>
          <c:extLst>
            <c:ext uri="{DDDDDDDD-DDDD-DDDD-DDDD-DDDDDDDDDDDD}">
              <c15:dummy><c15:payload>ser</c15:payload></c15:dummy>
            </c:ext>
          </c:extLst>
        </c:ser>
      </c:barChart>
      <c:catAx>
        <c:axId val="1"/>
        <c:extLst>
          <c:ext uri="{EEEEEEEE-EEEE-EEEE-EEEE-EEEEEEEEEEEE}">
            <c15:dummy><c15:payload>catAx</c15:payload></c15:dummy>
          </c:ext>
        </c:extLst>
      </c:catAx>
      <c:valAx>
        <c:axId val="2"/>
        <c:extLst>
          <c:ext uri="{FFFFFFFF-FFFF-FFFF-FFFF-FFFFFFFFFFFF}">
            <c15:dummy><c15:payload>valAx</c15:payload></c15:dummy>
          </c:ext>
        </c:extLst>
      </c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

    let model = parse_chart_space(xml, "in-memory-chart.xml").expect("parse chartSpace");

    assert!(model
        .chart_space_ext_lst_xml
        .as_ref()
        .expect("chartSpace extLst should be captured")
        .contains("{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}"));
    assert!(model
        .chart_ext_lst_xml
        .as_ref()
        .expect("chart extLst should be captured")
        .contains("{BBBBBBBB-BBBB-BBBB-BBBB-BBBBBBBBBBBB}"));
    assert!(model
        .plot_area_ext_lst_xml
        .as_ref()
        .expect("plotArea extLst should be captured")
        .contains("{CCCCCCCC-CCCC-CCCC-CCCC-CCCCCCCCCCCC}"));

    assert_eq!(model.series.len(), 1);
    assert!(model.series[0]
        .ext_lst_xml
        .as_ref()
        .expect("series extLst should be captured")
        .contains("{DDDDDDDD-DDDD-DDDD-DDDD-DDDDDDDDDDDD}"));

    let cat_ax = model.axes.iter().find(|ax| ax.id == 1).expect("catAx");
    assert!(cat_ax
        .ext_lst_xml
        .as_ref()
        .expect("axis extLst should be captured")
        .contains("{EEEEEEEE-EEEE-EEEE-EEEE-EEEEEEEEEEEE}"));

    let val_ax = model.axes.iter().find(|ax| ax.id == 2).expect("valAx");
    assert!(val_ax
        .ext_lst_xml
        .as_ref()
        .expect("axis extLst should be captured")
        .contains("{FFFFFFFF-FFFF-FFFF-FFFF-FFFFFFFFFFFF}"));
}
