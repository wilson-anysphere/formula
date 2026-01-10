#[test]
fn hyperlinks_roundtrip_preserves_targets_and_rids() {
    let fixture = std::fs::read(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/xlsx/hyperlinks/hyperlinks.xlsx"
    ))
    .expect("fixture exists");

    let pkg = formula_xlsx::XlsxPackage::from_bytes(&fixture).expect("parse package");
    let sheet_xml = std::str::from_utf8(
        pkg.part("xl/worksheets/sheet1.xml")
            .expect("sheet1.xml present"),
    )
    .expect("sheet1.xml is utf-8");
    let rels_xml = std::str::from_utf8(
        pkg.part("xl/worksheets/_rels/sheet1.xml.rels")
            .expect("sheet1.xml.rels present"),
    )
    .expect("sheet1.xml.rels is utf-8");

    let links =
        formula_xlsx::parse_worksheet_hyperlinks(sheet_xml, Some(rels_xml)).expect("parse links");
    assert_eq!(links.len(), 3);

    let external = &links[0];
    assert_eq!(external.range.to_string(), "A1");
    assert_eq!(external.rel_id.as_deref(), Some("rId1"));
    match &external.target {
        formula_model::HyperlinkTarget::ExternalUrl { uri } => {
            assert_eq!(uri, "https://example.com");
        }
        other => panic!("unexpected external target: {other:?}"),
    }

    let internal = &links[1];
    assert_eq!(internal.range.to_string(), "A2");
    assert!(internal.rel_id.is_none());
    match &internal.target {
        formula_model::HyperlinkTarget::Internal { sheet, cell } => {
            assert_eq!(sheet, "Sheet2");
            assert_eq!(cell.to_a1(), "B2");
        }
        other => panic!("unexpected internal target: {other:?}"),
    }

    let mail = &links[2];
    assert_eq!(mail.range.to_string(), "A3");
    assert_eq!(mail.rel_id.as_deref(), Some("rId2"));
    match &mail.target {
        formula_model::HyperlinkTarget::Email { uri } => {
            assert_eq!(uri, "mailto:test@example.com");
        }
        other => panic!("unexpected mail target: {other:?}"),
    }

    // Write updated XML using the hyperlink writer, then re-parse after a package write.
    let updated_sheet_xml =
        formula_xlsx::update_worksheet_xml(sheet_xml, &links).expect("write sheet xml");
    let updated_rels_xml = formula_xlsx::update_worksheet_relationships(Some(rels_xml), &links)
        .expect("write rels")
        .expect("rels xml present");

    let mut pkg2 = pkg.clone();
    pkg2.set_part("xl/worksheets/sheet1.xml", updated_sheet_xml.into_bytes());
    pkg2.set_part(
        "xl/worksheets/_rels/sheet1.xml.rels",
        updated_rels_xml.into_bytes(),
    );

    let roundtripped = pkg2.write_to_bytes().expect("write package");
    let pkg3 = formula_xlsx::XlsxPackage::from_bytes(&roundtripped).expect("read roundtrip");
    let sheet_xml2 = std::str::from_utf8(
        pkg3.part("xl/worksheets/sheet1.xml")
            .expect("sheet1.xml present"),
    )
    .expect("sheet1.xml utf-8");
    let rels_xml2 = std::str::from_utf8(
        pkg3.part("xl/worksheets/_rels/sheet1.xml.rels")
            .expect("sheet1.xml.rels present"),
    )
    .expect("sheet1.xml.rels utf-8");

    let links2 =
        formula_xlsx::parse_worksheet_hyperlinks(sheet_xml2, Some(rels_xml2)).expect("re-parse");
    assert_eq!(links2, links);
}
