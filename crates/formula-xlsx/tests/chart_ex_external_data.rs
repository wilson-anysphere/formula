use formula_xlsx::drawingml::charts::parse_chart_ex;

#[test]
fn chart_ex_parses_external_data_link() {
    let xml = r#"<cx:chartSpace
        xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <cx:externalData r:id="rId9">
        <cx:autoUpdate val="0"/>
      </cx:externalData>
      <cx:chart>
        <cx:plotArea>
          <cx:histogramChart/>
        </cx:plotArea>
      </cx:chart>
    </cx:chartSpace>"#;

    let model = parse_chart_ex(xml.as_bytes(), "chartEx1.xml").expect("parse chartEx");
    assert_eq!(model.external_data_rel_id.as_deref(), Some("rId9"));
    assert_eq!(model.external_data_auto_update, Some(false));
}

#[test]
fn chart_ex_external_data_auto_update_defaults_true_when_val_missing() {
    let xml = r#"<cx:chartSpace
        xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <cx:externalData r:id="rId7">
        <cx:autoUpdate/>
      </cx:externalData>
      <cx:chart>
        <cx:plotArea>
          <cx:histogramChart/>
        </cx:plotArea>
      </cx:chart>
    </cx:chartSpace>"#;

    let model = parse_chart_ex(xml.as_bytes(), "chartEx1.xml").expect("parse chartEx");
    assert_eq!(model.external_data_rel_id.as_deref(), Some("rId7"));
    assert_eq!(model.external_data_auto_update, Some(true));
}

#[test]
fn chart_ex_external_data_inside_alternate_content_fallback() {
    let xml = r#"<cx:chartSpace
        xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
      <mc:AlternateContent>
        <mc:Fallback>
          <cx:externalData r:id="rId8">
            <cx:autoUpdate val="0"/>
          </cx:externalData>
        </mc:Fallback>
      </mc:AlternateContent>
      <cx:chart>
        <cx:plotArea>
          <cx:histogramChart/>
        </cx:plotArea>
      </cx:chart>
    </cx:chartSpace>"#;

    let model = parse_chart_ex(xml.as_bytes(), "chartEx1.xml").expect("parse chartEx");
    assert_eq!(model.external_data_rel_id.as_deref(), Some("rId8"));
    assert_eq!(model.external_data_auto_update, Some(false));
}
