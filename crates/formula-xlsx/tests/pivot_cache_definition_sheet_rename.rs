use std::io::Write;

use formula_xlsx::XlsxPackage;
use rust_xlsxwriter::Workbook;
use zip::write::FileOptions;

fn build_source_package() -> Vec<u8> {
    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/pivotTables/pivotTable1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>
  <Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>
  <Override PartName="/xl/pivotCache/pivotCacheRecords1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/>
</Types>"#;

    let workbook = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <pivotCaches>
    <pivotCache cacheId="1" r:id="rId99"/>
  </pivotCaches>
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    <sheet name="Sheet2" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/>
</Relationships>"#;

    let sheet1 = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>"#;

    let sheet2 = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData/>
  <pivotTables>
    <pivotTable r:id="rId99"/>
  </pivotTables>
</worksheet>"#;

    let sheet2_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/>
</Relationships>"#;

    let pivot_table = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="PivotTable1" cacheId="1"/>"#;

    // Include both the standard `sheet="..."` attribute and the non-standard `ref="Sheet!A1:B2"`
    // encoding so the apply path can rewrite both.
    let cache_def = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cacheSource type="worksheet"><worksheetSource sheet="Sheet1" ref="Sheet1!A1:C5"/></cacheSource></pivotCacheDefinition>"#;
    let cache_records = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0"/>"#;

    let cache_def_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords1.xml"/>
</Relationships>"#;

    let cursor = std::io::Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(content_types.as_bytes()).unwrap();

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options).unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(sheet1).unwrap();

    zip.start_file("xl/worksheets/sheet2.xml", options).unwrap();
    zip.write_all(sheet2).unwrap();

    zip.start_file("xl/worksheets/_rels/sheet2.xml.rels", options)
        .unwrap();
    zip.write_all(sheet2_rels.as_bytes()).unwrap();

    zip.start_file("xl/pivotTables/pivotTable1.xml", options)
        .unwrap();
    zip.write_all(pivot_table).unwrap();

    zip.start_file("xl/pivotCache/pivotCacheDefinition1.xml", options)
        .unwrap();
    zip.write_all(cache_def).unwrap();

    zip.start_file("xl/pivotCache/pivotCacheRecords1.xml", options)
        .unwrap();
    zip.write_all(cache_records).unwrap();

    zip.start_file(
        "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels",
        options,
    )
    .unwrap();
    zip.write_all(cache_def_rels.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

fn build_destination_package() -> Vec<u8> {
    // Generate a baseline workbook using rust_xlsxwriter (matching the intended real-world
    // regenerate path).
    let mut workbook = Workbook::new();
    workbook.add_worksheet();
    workbook.add_worksheet();
    workbook.save_to_buffer().unwrap()
}

#[test]
fn preserved_pivot_cache_definition_worksheet_source_sheet_is_rewritten_after_sheet_rename() {
    let source_bytes = build_source_package();
    let source_pkg = XlsxPackage::from_bytes(&source_bytes).expect("read source package");
    let preserved = source_pkg
        .preserve_pivot_parts()
        .expect("preserve pivot parts");

    let dest_bytes = build_destination_package();
    let mut dest_pkg = XlsxPackage::from_bytes(&dest_bytes).expect("read destination");

    // Simulate a user renaming the worksheet before saving. The sheet index remains the same.
    let workbook_part = "xl/workbook.xml";
    let workbook_xml = std::str::from_utf8(dest_pkg.part(workbook_part).expect("workbook.xml"))
        .expect("workbook xml utf-8");
    dest_pkg.set_part(
        workbook_part,
        workbook_xml
            .replace("name=\"Sheet1\"", "name=\"Data\"")
            .into_bytes(),
    );

    dest_pkg
        .apply_preserved_pivot_parts(&preserved)
        .expect("apply preserved pivot parts");

    let cache_def_xml = std::str::from_utf8(
        dest_pkg
            .part("xl/pivotCache/pivotCacheDefinition1.xml")
            .expect("pivotCacheDefinition exists"),
    )
    .expect("pivotCacheDefinition utf-8");
    assert!(
        cache_def_xml.contains(r#"sheet="Data""#),
        "expected worksheetSource sheet name to be rewritten: {cache_def_xml}"
    );
    assert!(
        !cache_def_xml.contains(r#"sheet="Sheet1""#),
        "expected old worksheetSource sheet name to be removed: {cache_def_xml}"
    );
    assert!(
        cache_def_xml.contains(r#"ref="Data!A1:C5""#),
        "expected worksheetSource ref to be rewritten: {cache_def_xml}"
    );
}

