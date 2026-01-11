use std::io::{Cursor, Write};

use formula_xlsx::XlsxPackage;
use pretty_assertions::assert_eq;

const FIXTURE: &[u8] = include_bytes!("fixtures/pivot-full.xlsx");

#[test]
fn pivot_graph_resolves_full_chain() {
    let pkg = XlsxPackage::from_bytes(FIXTURE).expect("read pkg");
    let graph = pkg.pivot_graph().expect("pivot graph");

    assert_eq!(graph.pivot_tables.len(), 1);
    let pt = &graph.pivot_tables[0];
    assert_eq!(pt.pivot_table_part, "xl/pivotTables/pivotTable1.xml");
    assert_eq!(pt.sheet_part.as_deref(), Some("xl/worksheets/sheet1.xml"));
    assert_eq!(pt.sheet_name.as_deref(), Some("Sheet1"));
    assert_eq!(pt.cache_id, Some(1));
    assert_eq!(
        pt.cache_definition_part.as_deref(),
        Some("xl/pivotCache/pivotCacheDefinition1.xml")
    );
    assert_eq!(
        pt.cache_records_part.as_deref(),
        Some("xl/pivotCache/pivotCacheRecords1.xml")
    );
}

#[test]
fn preserve_pivot_parts_captures_workbook_and_sheet_subtrees() {
    let pkg = XlsxPackage::from_bytes(FIXTURE).expect("read pkg");
    let preserved = pkg.preserve_pivot_parts().expect("preserve pivot parts");

    let workbook = preserved
        .workbook_pivot_caches
        .as_deref()
        .expect("workbook pivotCaches subtree present");
    let workbook_str = std::str::from_utf8(workbook).unwrap();
    assert!(workbook_str.contains("<pivotCaches"));
    assert!(workbook_str.contains("cacheId=\"1\""));
    assert!(workbook_str.contains("r:id=\"rId2\""));

    let sheet = preserved
        .sheet_pivot_tables
        .get("Sheet1")
        .expect("Sheet1 pivotTables subtree");
    let sheet_str = std::str::from_utf8(&sheet.pivot_tables_xml).unwrap();
    assert!(sheet_str.contains("<pivotTables"));
    assert!(sheet_str.contains("pivotTable"));
    assert!(sheet_str.contains("r:id=\"rId2\""));
}

#[test]
fn apply_preserved_pivot_parts_inserts_subtrees_and_relationships() {
    let src = XlsxPackage::from_bytes(FIXTURE).expect("read src pkg");
    let preserved = src.preserve_pivot_parts().expect("preserve");

    let dest_bytes = build_minimal_destination_package();
    let mut dest = XlsxPackage::from_bytes(&dest_bytes).expect("read dest pkg");
    dest.apply_preserved_pivot_parts(&preserved)
        .expect("apply preserved pivot parts");

    let workbook_xml = std::str::from_utf8(dest.part("xl/workbook.xml").unwrap()).unwrap();
    assert!(workbook_xml.contains("xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\""));
    assert!(workbook_xml.contains("<pivotCaches"));
    assert!(workbook_xml.contains("cacheId=\"1\""));

    let workbook_rels = std::str::from_utf8(dest.part("xl/_rels/workbook.xml.rels").unwrap()).unwrap();
    assert!(workbook_rels.contains("Id=\"rId2\""));
    assert!(workbook_rels.contains(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition"
    ));
    assert!(workbook_rels.contains("Target=\"pivotCache/pivotCacheDefinition1.xml\""));

    let sheet_xml = std::str::from_utf8(dest.part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    assert!(sheet_xml.contains("xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\""));
    assert!(sheet_xml.contains("<pivotTables"));
    assert!(sheet_xml.contains("r:id=\"rId2\""));

    let sheet_rels =
        std::str::from_utf8(dest.part("xl/worksheets/_rels/sheet1.xml.rels").unwrap()).unwrap();
    assert!(sheet_rels.contains("Id=\"rId2\""));
    assert!(sheet_rels.contains(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable"
    ));
    assert!(sheet_rels.contains("Target=\"../pivotTables/pivotTable1.xml\""));

    let content_types = std::str::from_utf8(dest.part("[Content_Types].xml").unwrap()).unwrap();
    assert!(content_types.contains("/xl/pivotTables/pivotTable1.xml"));
    assert!(content_types.contains("/xl/pivotCache/pivotCacheDefinition1.xml"));
    assert!(content_types.contains("/xl/pivotCache/pivotCacheRecords1.xml"));

    // The preserved package should be usable by the pivot graph resolver end-to-end.
    let graph = dest.pivot_graph().expect("pivot graph");
    assert_eq!(graph.pivot_tables.len(), 1);
    let pt = &graph.pivot_tables[0];
    assert_eq!(pt.pivot_table_part, "xl/pivotTables/pivotTable1.xml");
    assert_eq!(
        pt.cache_definition_part.as_deref(),
        Some("xl/pivotCache/pivotCacheDefinition1.xml")
    );
    assert_eq!(
        pt.cache_records_part.as_deref(),
        Some("xl/pivotCache/pivotCacheRecords1.xml")
    );
}

fn build_minimal_destination_package() -> Vec<u8> {
    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>
"#;

    let root_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
"#;

    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>
"#;

    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>
"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(content_types.as_bytes()).unwrap();

    zip.start_file("_rels/.rels", options).unwrap();
    zip.write_all(root_rels.as_bytes()).unwrap();

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options).unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(worksheet_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}
