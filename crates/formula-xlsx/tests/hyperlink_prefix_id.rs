use formula_model::HyperlinkTarget;

#[test]
fn parses_hyperlink_relationship_id_with_custom_prefix() {
    let sheet_xml = r#"
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <hyperlinks>
            <hyperlink ref="A1" rel:id="rId5" xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
          </hyperlinks>
        </worksheet>
    "#;

    let rels_xml = r#"
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship
            Id="rId5"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
            Target="https://example.com"
            TargetMode="External"
          />
        </Relationships>
    "#;

    let links =
        formula_xlsx::parse_worksheet_hyperlinks(sheet_xml, Some(rels_xml)).expect("parse links");
    assert_eq!(links.len(), 1);

    let link = &links[0];
    assert_eq!(link.range.to_string(), "A1");
    assert_eq!(link.rel_id.as_deref(), Some("rId5"));
    match &link.target {
        HyperlinkTarget::ExternalUrl { uri } => assert_eq!(uri, "https://example.com"),
        other => panic!("unexpected hyperlink target: {other:?}"),
    }
}
