use std::io::{Cursor, Write};

use formula_xlsx::XlsxPackage;
use pretty_assertions::assert_eq;

#[test]
fn resolves_pivot_table_sheet_and_cache_parts() {
    let bytes = build_synthetic_pivot_package();
    let pkg = XlsxPackage::from_bytes(&bytes).expect("read package");

    let graph = pkg.pivot_graph().expect("resolve pivot graph");
    assert_eq!(graph.pivot_tables.len(), 1);

    let table = &graph.pivot_tables[0];
    assert_eq!(table.pivot_table_part, "xl/pivotTables/pivotTable1.xml");
    assert_eq!(
        table.sheet_part.as_deref(),
        Some("xl/worksheets/sheet1.xml")
    );
    assert_eq!(table.sheet_name.as_deref(), Some("Sheet1"));
    assert_eq!(table.cache_id, Some(1));
    assert_eq!(
        table.cache_definition_part.as_deref(),
        Some("xl/pivotCache/pivotCacheDefinition1.xml")
    );
    assert_eq!(
        table.cache_records_part.as_deref(),
        Some("xl/pivotCache/pivotCacheRecords1.xml")
    );
}

#[test]
fn includes_pivot_table_even_without_cache_id() {
    let fixture = include_bytes!("fixtures/pivot_slicers_and_chart.xlsx");
    let pkg = XlsxPackage::from_bytes(fixture).expect("read fixture");

    let graph = pkg.pivot_graph().expect("resolve pivot graph");
    assert_eq!(graph.pivot_tables.len(), 1);

    let table = &graph.pivot_tables[0];
    assert_eq!(table.pivot_table_part, "xl/pivotTables/pivotTable1.xml");
    assert_eq!(
        table.sheet_part.as_deref(),
        Some("xl/worksheets/sheet1.xml")
    );
    assert_eq!(table.sheet_name.as_deref(), Some("Sheet1"));
    assert_eq!(table.cache_id, None);
    assert_eq!(table.cache_definition_part, None);
    assert_eq!(table.cache_records_part, None);
}

#[test]
fn resolves_cache_parts_by_filename_when_workbook_omits_pivot_caches() {
    // This fixture is intentionally missing `workbook.xml` pivotCaches and any relevant `.rels`
    // entries, but still contains `pivotCacheDefinition1.xml` + `pivotCacheRecords1.xml`. We
    // should still resolve the cache parts by the common `...Definition{cacheId}.xml` pattern.
    let fixture = include_bytes!("fixtures/pivot-fixture.xlsx");
    let pkg = XlsxPackage::from_bytes(fixture).expect("read fixture");

    let graph = pkg.pivot_graph().expect("resolve pivot graph");
    assert_eq!(graph.pivot_tables.len(), 1);

    let table = &graph.pivot_tables[0];
    assert_eq!(table.pivot_table_part, "xl/pivotTables/pivotTable1.xml");
    assert_eq!(table.sheet_part, None);
    assert_eq!(table.sheet_name, None);
    assert_eq!(table.cache_id, Some(1));
    assert_eq!(
        table.cache_definition_part.as_deref(),
        Some("xl/pivotCache/pivotCacheDefinition1.xml")
    );
    assert_eq!(
        table.cache_records_part.as_deref(),
        Some("xl/pivotCache/pivotCacheRecords1.xml")
    );
}

#[test]
fn resolves_pivot_cache_parts_for_pivot_table() {
    let fixture = include_bytes!("fixtures/pivot-full.xlsx");
    let pkg = XlsxPackage::from_bytes(fixture).expect("read fixture");

    let parts = pkg
        .pivot_cache_parts_for_pivot_table("xl/pivotTables/pivotTable1.xml")
        .expect("resolve pivot cache parts");

    assert_eq!(
        parts,
        Some((
            "xl/pivotCache/pivotCacheDefinition1.xml".to_string(),
            "xl/pivotCache/pivotCacheRecords1.xml".to_string()
        ))
    );
}

fn build_synthetic_pivot_package() -> Vec<u8> {
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
  <pivotCaches>
    <pivotCache cacheId="1" r:id="rId2"/>
  </pivotCaches>
</workbook>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/>
</Relationships>"#;

    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData/>
</worksheet>"#;

    let worksheet_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/>
</Relationships>"#;

    let pivot_table_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  name="PivotTable1" cacheId="1"/>"#;

    let cache_definition_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" recordCount="0"/>"#;

    let cache_definition_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords1.xml"/>
</Relationships>"#;

    let cache_records_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0"/>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    for (name, xml) in [
        ("xl/workbook.xml", workbook_xml),
        ("xl/_rels/workbook.xml.rels", workbook_rels),
        ("xl/worksheets/sheet1.xml", worksheet_xml),
        ("xl/worksheets/_rels/sheet1.xml.rels", worksheet_rels),
        ("xl/pivotTables/pivotTable1.xml", pivot_table_xml),
        (
            "xl/pivotCache/pivotCacheDefinition1.xml",
            cache_definition_xml,
        ),
        (
            "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels",
            cache_definition_rels,
        ),
        ("xl/pivotCache/pivotCacheRecords1.xml", cache_records_xml),
    ] {
        zip.start_file(name, options).unwrap();
        zip.write_all(xml.as_bytes()).unwrap();
    }

    zip.finish().unwrap().into_inner()
}
