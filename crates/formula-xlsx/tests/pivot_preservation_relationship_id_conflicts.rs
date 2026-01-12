use std::io::{Cursor, Write};

use formula_xlsx::openxml;
use formula_xlsx::XlsxPackage;
use roxmltree::Document;
use zip::write::FileOptions;
use zip::ZipWriter;

const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const REL_TYPE_STYLES: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
const REL_TYPE_DRAWING: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
const REL_TYPE_PIVOT_CACHE_DEFINITION: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition";
const REL_TYPE_PIVOT_TABLE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable";

#[test]
fn pivot_preservation_renumbers_relationship_ids_on_conflict() {
    let source_bytes = build_source_with_pivots_using_rid2();
    let source_pkg = XlsxPackage::from_bytes(&source_bytes).expect("read source pkg");
    let preserved = source_pkg.preserve_pivot_parts().expect("preserve pivots");
    assert!(!preserved.is_empty(), "fixture should preserve something");

    let dest_bytes = build_destination_with_rid2_conflicts();
    let mut dest_pkg = XlsxPackage::from_bytes(&dest_bytes).expect("read destination pkg");
    dest_pkg
        .apply_preserved_pivot_parts(&preserved)
        .expect("apply pivots");

    // Round-trip through zip writer to exercise the same code path as production.
    let merged_bytes = dest_pkg.write_to_bytes().expect("write merged pkg");
    let merged_pkg = XlsxPackage::from_bytes(&merged_bytes).expect("read merged pkg");

    let workbook_xml = std::str::from_utf8(merged_pkg.part("xl/workbook.xml").unwrap()).unwrap();
    let workbook_doc = Document::parse(workbook_xml).unwrap();
    let pivot_cache = workbook_doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "pivotCache")
        .expect("workbook should contain <pivotCache>");
    let cache_rid = pivot_cache
        .attribute((REL_NS, "id"))
        .or_else(|| pivot_cache.attribute("r:id"))
        .expect("pivotCache should have r:id");
    assert_eq!(
        cache_rid, "rId3",
        "pivot cache relationship should be renumbered away from destination rId2"
    );

    let workbook_rels_xml =
        std::str::from_utf8(merged_pkg.part("xl/_rels/workbook.xml.rels").unwrap()).unwrap();
    let (cache_rel_type, cache_rel_target, _) =
        find_relationship(workbook_rels_xml, cache_rid).expect("cache relationship exists");
    assert_eq!(cache_rel_type, REL_TYPE_PIVOT_CACHE_DEFINITION);
    assert_eq!(cache_rel_target, "pivotCache/pivotCacheDefinition1.xml");

    // Ensure we did not corrupt unrelated relationships (styles should stay on rId2).
    let (styles_type, styles_target, _) =
        find_relationship(workbook_rels_xml, "rId2").expect("styles relationship exists");
    assert_eq!(styles_type, REL_TYPE_STYLES);
    assert_eq!(styles_target, "styles.xml");

    let sheet_xml =
        std::str::from_utf8(merged_pkg.part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    let sheet_doc = Document::parse(sheet_xml).unwrap();
    let pivot_table = sheet_doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "pivotTable")
        .expect("sheet should contain <pivotTable>");
    let table_rid = pivot_table
        .attribute((REL_NS, "id"))
        .or_else(|| pivot_table.attribute("r:id"))
        .expect("pivotTable should have r:id");
    assert_eq!(
        table_rid, "rId3",
        "pivot table relationship should be renumbered away from destination rId2"
    );

    let sheet_rels_xml = std::str::from_utf8(
        merged_pkg
            .part("xl/worksheets/_rels/sheet1.xml.rels")
            .unwrap(),
    )
    .unwrap();
    let (table_rel_type, table_rel_target, _raw) =
        find_relationship(sheet_rels_xml, table_rid).expect("pivotTable relationship exists");
    assert_eq!(table_rel_type, REL_TYPE_PIVOT_TABLE);
    assert_eq!(table_rel_target, "../pivotTables/pivotTable1.xml");

    // Ensure we didn't rewrite existing relationships and lose attributes like TargetMode.
    assert!(
        sheet_rels_xml.contains("TargetMode=\"External\""),
        "existing TargetMode attribute should remain intact"
    );
    let (drawing_type, drawing_target, _) =
        find_relationship(sheet_rels_xml, "rId2").expect("drawing relationship exists");
    assert_eq!(drawing_type, REL_TYPE_DRAWING);
    assert_eq!(drawing_target, "../drawings/drawing1.xml");

    // pivot_graph(): make sure the relationship chain resolves to real parts.
    let cache_part = openxml::resolve_target("xl/workbook.xml", &cache_rel_target);
    assert!(
        merged_pkg.part(&cache_part).is_some(),
        "workbook pivot cache target should exist: {cache_part}"
    );

    let pivot_table_part = openxml::resolve_target("xl/worksheets/sheet1.xml", &table_rel_target);
    assert!(
        merged_pkg.part(&pivot_table_part).is_some(),
        "worksheet pivot table target should exist: {pivot_table_part}"
    );
}

fn find_relationship(xml: &str, id: &str) -> Option<(String, String, String)> {
    let doc = Document::parse(xml).ok()?;
    for node in doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "Relationship")
    {
        if node.attribute("Id")? != id {
            continue;
        }
        let type_ = node.attribute("Type")?.to_string();
        let target = node.attribute("Target")?.to_string();
        return Some((type_, target, node.text().unwrap_or_default().to_string()));
    }
    None
}

fn build_source_with_pivots_using_rid2() -> Vec<u8> {
    let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
  <pivotCaches>
    <pivotCache cacheId="1" r:id="rId2"/>
  </pivotCaches>
</workbook>"#;

    let workbook_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/>
</Relationships>"#;

    let sheet_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData/>
  <pivotTables>
    <pivotTable r:id="rId2"/>
  </pivotTables>
</worksheet>"#;

    let sheet_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/>
</Relationships>"#;

    let pivot_table_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="PivotTable1"/>"#;

    let pivot_cache_def_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#;

    let pivot_cache_records_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#;

    let content_types = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/pivotTables/pivotTable1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>
  <Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>
  <Override PartName="/xl/pivotCache/pivotCacheRecords1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/>
</Types>"#;

    let root_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#;

    build_zip(&[
        ("[Content_Types].xml", content_types),
        ("_rels/.rels", root_rels),
        ("xl/workbook.xml", workbook_xml),
        ("xl/_rels/workbook.xml.rels", workbook_rels),
        ("xl/worksheets/sheet1.xml", sheet_xml),
        ("xl/worksheets/_rels/sheet1.xml.rels", sheet_rels),
        ("xl/pivotTables/pivotTable1.xml", pivot_table_xml),
        (
            "xl/pivotCache/pivotCacheDefinition1.xml",
            pivot_cache_def_xml,
        ),
        (
            "xl/pivotCache/pivotCacheRecords1.xml",
            pivot_cache_records_xml,
        ),
    ])
}

fn build_destination_with_rid2_conflicts() -> Vec<u8> {
    let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    // Destination already consumes rId2 for styles (common in rust_xlsxwriter output).
    let workbook_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#;

    let sheet_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData/>
</worksheet>"#;

    // And the sheet rels already consume rId2 for drawing (also common).
    // Include a TargetMode-bearing hyperlink relationship to ensure we don't drop extra attrs.
    let sheet_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>"#;

    let styles_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#;
    let drawing_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>"#;

    let content_types = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>"#;

    let root_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#;

    build_zip(&[
        ("[Content_Types].xml", content_types),
        ("_rels/.rels", root_rels),
        ("xl/workbook.xml", workbook_xml),
        ("xl/_rels/workbook.xml.rels", workbook_rels),
        ("xl/worksheets/sheet1.xml", sheet_xml),
        ("xl/worksheets/_rels/sheet1.xml.rels", sheet_rels),
        ("xl/styles.xml", styles_xml),
        ("xl/drawings/drawing1.xml", drawing_xml),
    ])
}

fn build_zip(entries: &[(&str, &[u8])]) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    for (name, bytes) in entries {
        zip.start_file(*name, options).unwrap();
        zip.write_all(bytes).unwrap();
    }

    zip.finish().unwrap().into_inner()
}
