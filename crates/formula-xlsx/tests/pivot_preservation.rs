use std::io::Write;

use formula_xlsx::XlsxPackage;
use pretty_assertions::assert_eq;
use rust_xlsxwriter::Workbook;
use zip::write::FileOptions;

fn build_source_package() -> Vec<u8> {
    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/pivotTables/pivotTable1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>
  <Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>
  <Override PartName="/xl/pivotCache/pivotCacheRecords1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/>
  <Override PartName="/xl/slicers/slicer1.xml" ContentType="application/vnd.ms-excel.slicer+xml"/>
  <Override PartName="/xl/slicerCaches/slicerCache1.xml" ContentType="application/vnd.ms-excel.slicerCache+xml"/>
  <Override PartName="/xl/timelines/timeline1.xml" ContentType="application/vnd.ms-excel.timeline+xml"/>
  <Override PartName="/xl/timelineCaches/timelineCacheDefinition1.xml" ContentType="application/vnd.ms-excel.timelineCacheDefinition+xml"/>
</Types>"#;

    let workbook = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
  <pivotCaches>
    <pivotCache cacheId="1" r:id="rId99"/>
  </pivotCaches>
  <slicerCaches>
    <slicerCache r:id="rId98"/>
  </slicerCaches>
  <timelineCaches>
    <timelineCache r:id="rId97"/>
  </timelineCaches>
</workbook>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/>
  <Relationship Id="rId98" Type="http://schemas.microsoft.com/office/2007/relationships/slicerCache" Target="slicerCaches/slicerCache1.xml"/>
  <Relationship Id="rId97" Type="http://schemas.microsoft.com/office/2007/relationships/timelineCacheDefinition" Target="timelineCaches/timelineCacheDefinition1.xml"/>
</Relationships>"#;

    let sheet = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData/>
  <pivotTables>
    <pivotTable r:id="rId99"/>
  </pivotTables>
</worksheet>"#;

    let sheet_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/>
</Relationships>"#;

    let pivot_table = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="PivotTable1" cacheId="1"/>"#;
    let cache_def = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" recordCount="1"/>"#;
    let cache_records = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1"/>"#;

    let cache_def_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords1.xml"/>
</Relationships>"#;

    let slicer = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><slicer xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" name="Slicer1"/>"#;
    let slicer_cache = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><slicerCache xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" name="SlicerCache1"/>"#;
    let timeline = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><timeline xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" name="Timeline1"/>"#;
    let timeline_cache = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><timelineCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" name="TimelineCache1"/>"#;

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
    zip.write_all(sheet.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/_rels/sheet1.xml.rels", options)
        .unwrap();
    zip.write_all(sheet_rels.as_bytes()).unwrap();

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

    zip.start_file("xl/slicers/slicer1.xml", options).unwrap();
    zip.write_all(slicer).unwrap();

    zip.start_file("xl/slicerCaches/slicerCache1.xml", options)
        .unwrap();
    zip.write_all(slicer_cache).unwrap();

    zip.start_file("xl/timelines/timeline1.xml", options).unwrap();
    zip.write_all(timeline).unwrap();

    zip.start_file(
        "xl/timelineCaches/timelineCacheDefinition1.xml",
        options,
    )
    .unwrap();
    zip.write_all(timeline_cache).unwrap();

    zip.finish().unwrap().into_inner()
}

fn build_destination_package() -> Vec<u8> {
    // Generate a baseline workbook using rust_xlsxwriter (matching the intended real-world
    // regenerate path).
    let mut workbook = Workbook::new();
    workbook.add_worksheet();
    workbook.save_to_buffer().unwrap()
}

#[test]
fn preserved_pivot_parts_can_be_reapplied_to_regenerated_workbook() {
    let source_bytes = build_source_package();
    let source_pkg = XlsxPackage::from_bytes(&source_bytes).expect("read source package");
    let preserved = source_pkg
        .preserve_pivot_parts()
        .expect("preserve pivot parts");
    assert!(!preserved.is_empty());

    let original_part = source_pkg
        .part("xl/pivotCache/pivotCacheDefinition1.xml")
        .unwrap()
        .to_vec();
    let original_slicer_cache = source_pkg
        .part("xl/slicerCaches/slicerCache1.xml")
        .unwrap()
        .to_vec();
    let original_timeline_cache = source_pkg
        .part("xl/timelineCaches/timelineCacheDefinition1.xml")
        .unwrap()
        .to_vec();

    let dest_bytes = build_destination_package();
    let mut dest_pkg = XlsxPackage::from_bytes(&dest_bytes).expect("read destination");
    dest_pkg
        .apply_preserved_pivot_parts(&preserved)
        .expect("apply preserved pivot parts");

    assert_eq!(
        dest_pkg.part("xl/pivotCache/pivotCacheDefinition1.xml"),
        Some(original_part.as_slice())
    );
    assert_eq!(
        dest_pkg.part("xl/slicerCaches/slicerCache1.xml"),
        Some(original_slicer_cache.as_slice())
    );
    assert_eq!(
        dest_pkg.part("xl/timelineCaches/timelineCacheDefinition1.xml"),
        Some(original_timeline_cache.as_slice())
    );

    let workbook_xml = std::str::from_utf8(dest_pkg.part("xl/workbook.xml").unwrap()).unwrap();
    assert!(workbook_xml.contains("<pivotCaches"));
    assert!(workbook_xml.contains("cacheId=\"1\""));
    assert!(
        workbook_xml.contains("slicerCaches"),
        "expected workbook.xml to contain slicerCaches"
    );
    assert!(
        workbook_xml.contains("timelineCaches"),
        "expected workbook.xml to contain timelineCaches"
    );

    let workbook_rels =
        std::str::from_utf8(dest_pkg.part("xl/_rels/workbook.xml.rels").unwrap()).unwrap();
    assert!(workbook_rels.contains("pivotCacheDefinition"));
    assert!(workbook_rels.contains("Id=\"rId99\""));
    assert!(workbook_rels.contains("Target=\"pivotCache/pivotCacheDefinition1.xml\""));
    assert!(workbook_rels.contains("relationships/slicerCache"));
    assert!(workbook_rels.contains("Target=\"slicerCaches/slicerCache1.xml\""));
    assert!(workbook_rels.contains("relationships/timelineCacheDefinition"));
    assert!(workbook_rels.contains("Target=\"timelineCaches/timelineCacheDefinition1.xml\""));

    let sheet_xml = std::str::from_utf8(dest_pkg.part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    assert!(sheet_xml.contains("xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\""));
    assert!(sheet_xml.contains("<pivotTables"));
    assert!(sheet_xml.contains("r:id=\"rId99\""));
    let pivot_pos = sheet_xml.find("<pivotTables").unwrap();
    if let Some(ext_pos) = sheet_xml.find("<extLst") {
        assert!(pivot_pos < ext_pos, "pivotTables should be inserted before extLst");
    } else {
        let close_pos = sheet_xml.rfind("</worksheet>").unwrap();
        assert!(pivot_pos < close_pos, "pivotTables should be inserted before </worksheet>");
    }

    let sheet_rels = std::str::from_utf8(
        dest_pkg
            .part("xl/worksheets/_rels/sheet1.xml.rels")
            .expect("sheet rels created"),
    )
    .unwrap();
    assert!(sheet_rels.contains("relationships/pivotTable"));
    assert!(sheet_rels.contains("Id=\"rId99\""));
    assert!(sheet_rels.contains("Target=\"../pivotTables/pivotTable1.xml\""));

    let ct = std::str::from_utf8(dest_pkg.part("[Content_Types].xml").unwrap()).unwrap();
    assert!(ct.contains("PartName=\"/xl/pivotTables/pivotTable1.xml\""));
    assert!(ct.contains("PartName=\"/xl/pivotCache/pivotCacheDefinition1.xml\""));
    assert!(ct.contains("PartName=\"/xl/pivotCache/pivotCacheRecords1.xml\""));
    assert!(ct.contains("PartName=\"/xl/slicers/slicer1.xml\""));
    assert!(ct.contains("PartName=\"/xl/slicerCaches/slicerCache1.xml\""));
    assert!(ct.contains("PartName=\"/xl/timelineCaches/timelineCacheDefinition1.xml\""));
}
