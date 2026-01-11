use formula_xlsx::pivots::preserve::ensure_sheet_xml_has_pivot_tables;

#[test]
fn inserts_pivot_tables_before_worksheet_close_when_no_ext_lst() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>"#;
    let pivot_tables = br#"<pivotTables><pivotTable r:id="rId7"/></pivotTables>"#;

    let updated =
        ensure_sheet_xml_has_pivot_tables(xml, "xl/worksheets/sheet1.xml", pivot_tables).unwrap();
    let updated_str = std::str::from_utf8(&updated).unwrap();

    assert_eq!(updated_str.matches("<pivotTables").count(), 1);
    assert!(updated_str.contains(r#"r:id="rId7""#));

    let sheet_data_pos = updated_str.find("<sheetData").unwrap();
    let pivot_tables_pos = updated_str.find("<pivotTables").unwrap();
    assert!(
        sheet_data_pos < pivot_tables_pos,
        "pivotTables should not be inserted inside <sheetData>"
    );

    assert!(
        updated_str.contains(
            r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#
        ),
        "xmlns:r should be added when inserting r:id attributes"
    );
}

#[test]
fn inserts_pivot_tables_before_ext_lst_when_present() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/><extLst><ext/></extLst></worksheet>"#;
    let pivot_tables = br#"<pivotTables><pivotTable r:id="rId1"/></pivotTables>"#;

    let updated =
        ensure_sheet_xml_has_pivot_tables(xml, "xl/worksheets/sheet1.xml", pivot_tables).unwrap();
    let updated_str = std::str::from_utf8(&updated).unwrap();

    assert_eq!(updated_str.matches("<pivotTables").count(), 1);
    assert!(updated_str.contains(r#"r:id="rId1""#));

    let pivot_tables_pos = updated_str.find("<pivotTables").unwrap();
    let ext_lst_pos = updated_str.find("<extLst").unwrap();
    assert!(
        pivot_tables_pos < ext_lst_pos,
        "<pivotTables> should be inserted before <extLst>"
    );
}

#[test]
fn merges_into_existing_pivot_tables_instead_of_creating_duplicates() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/><pivotTables count="1"><pivotTable r:id="rId1"/></pivotTables><extLst><ext/></extLst></worksheet>"#;
    let pivot_tables =
        br#"<pivotTables><pivotTable r:id="rId1"/><pivotTable r:id="rId2"/></pivotTables>"#;

    let updated =
        ensure_sheet_xml_has_pivot_tables(xml, "xl/worksheets/sheet1.xml", pivot_tables).unwrap();
    let updated_str = std::str::from_utf8(&updated).unwrap();

    assert_eq!(updated_str.matches("<pivotTables").count(), 1);

    // Merge is by relationship Id, so rId1 should not be duplicated.
    assert_eq!(updated_str.matches(r#"r:id="rId1""#).count(), 1);
    assert!(updated_str.contains(r#"r:id="rId2""#));

    let pivot_tables_pos = updated_str.find("<pivotTables").unwrap();
    let ext_lst_pos = updated_str.find("<extLst").unwrap();
    assert!(
        pivot_tables_pos < ext_lst_pos,
        "<pivotTables> should remain before <extLst>"
    );
}
