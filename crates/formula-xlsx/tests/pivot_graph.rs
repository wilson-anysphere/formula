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

#[test]
fn pivot_graph_tolerates_malformed_workbook_relationships() {
    // Malformed `xl/_rels/workbook.xml.rels` should not prevent pivot discovery, and cache parts
    // should still be resolved via the `pivotCacheDefinition{cacheId}.xml` naming convention.
    let bytes = build_synthetic_pivot_package_with_malformed_workbook_rels();
    let pkg = XlsxPackage::from_bytes(&bytes).expect("read package");

    let graph = pkg.pivot_graph().expect("resolve pivot graph");
    assert_eq!(graph.pivot_tables.len(), 1);

    let table = &graph.pivot_tables[0];
    assert_eq!(table.pivot_table_part, "xl/pivotTables/pivotTable1.xml");
    assert_eq!(
        table.sheet_part.as_deref(),
        Some("xl/worksheets/sheet1.xml")
    );
    // With malformed workbook `.rels`, we can't resolve `sheet1.xml` -> "Sheet1".
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
fn pivot_graph_falls_back_when_sheet_rels_are_malformed() {
    // Malformed `xl/worksheets/_rels/sheet1.xml.rels` should not error, and the pivot table should
    // still be returned via the fallback scan of `xl/pivotTables/*.xml`.
    let bytes = build_synthetic_pivot_package_with_malformed_sheet_rels();
    let pkg = XlsxPackage::from_bytes(&bytes).expect("read package");

    let graph = pkg.pivot_graph().expect("resolve pivot graph");
    assert_eq!(graph.pivot_tables.len(), 1);

    let table = &graph.pivot_tables[0];
    assert_eq!(table.pivot_table_part, "xl/pivotTables/pivotTable1.xml");
    assert_eq!(table.sheet_part, None);
}

#[test]
fn pivot_graph_tolerates_malformed_cache_definition_rels() {
    // Malformed cache definition `.rels` should not error. When we cannot resolve the cache records
    // part from relationships and the conventional filename doesn't exist, we should return the
    // pivot table but leave `cache_records_part` unset.
    let bytes = build_synthetic_pivot_package_with_malformed_cache_definition_rels();
    let pkg = XlsxPackage::from_bytes(&bytes).expect("read package");

    let graph = pkg.pivot_graph().expect("resolve pivot graph");
    assert_eq!(graph.pivot_tables.len(), 1);

    let table = &graph.pivot_tables[0];
    assert_eq!(table.pivot_table_part, "xl/pivotTables/pivotTable1.xml");
    assert_eq!(
        table.cache_definition_part.as_deref(),
        Some("xl/pivotCache/pivotCacheDefinition1.xml")
    );
    assert_eq!(table.cache_records_part, None);
}

#[test]
fn pivot_graph_resolves_cache_parts_without_workbook_xml() {
    // Even if `xl/workbook.xml` is missing, we should still attempt to resolve cache parts using
    // the conventional `pivotCacheDefinition{cacheId}.xml` / `pivotCacheRecords{cacheId}.xml`
    // naming convention.
    let bytes = build_synthetic_pivot_package_without_workbook_xml();
    let pkg = XlsxPackage::from_bytes(&bytes).expect("read package");

    let graph = pkg.pivot_graph().expect("resolve pivot graph");
    assert_eq!(graph.pivot_tables.len(), 1);

    let table = &graph.pivot_tables[0];
    assert_eq!(table.pivot_table_part, "xl/pivotTables/pivotTable1.xml");
    assert_eq!(
        table.sheet_part.as_deref(),
        Some("xl/worksheets/sheet1.xml")
    );
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
fn pivot_graph_tolerates_malformed_pivot_table_xml() {
    // If a pivot table part exists but its XML is malformed, we should still include it in the
    // graph (best-effort pivot discovery), just without `cache_id` / cache parts.
    let bytes = build_synthetic_pivot_package_with_malformed_pivot_table_xml();
    let pkg = XlsxPackage::from_bytes(&bytes).expect("read package");

    let graph = pkg.pivot_graph().expect("resolve pivot graph");
    assert_eq!(graph.pivot_tables.len(), 1);

    let table = &graph.pivot_tables[0];
    assert_eq!(table.pivot_table_part, "xl/pivotTables/pivotTable1.xml");
    assert_eq!(table.sheet_part, None);
    assert_eq!(table.cache_id, None);
    assert_eq!(table.cache_definition_part, None);
    assert_eq!(table.cache_records_part, None);

    let cache_parts = pkg
        .pivot_cache_parts_for_pivot_table("xl/pivotTables/pivotTable1.xml")
        .expect("resolve cache parts");
    assert_eq!(cache_parts, None);
}

fn build_synthetic_pivot_package() -> Vec<u8> {
    build_synthetic_pivot_package_with_overrides(
        VALID_WORKBOOK_RELS.as_bytes(),
        VALID_WORKSHEET_RELS.as_bytes(),
        VALID_CACHE_DEFINITION_RELS.as_bytes(),
        true,
    )
}

fn build_synthetic_pivot_package_with_malformed_workbook_rels() -> Vec<u8> {
    build_synthetic_pivot_package_with_overrides(
        b"<Relationships",
        VALID_WORKSHEET_RELS.as_bytes(),
        VALID_CACHE_DEFINITION_RELS.as_bytes(),
        true,
    )
}

fn build_synthetic_pivot_package_with_malformed_sheet_rels() -> Vec<u8> {
    build_synthetic_pivot_package_with_overrides(
        VALID_WORKBOOK_RELS.as_bytes(),
        b"<Relationships",
        VALID_CACHE_DEFINITION_RELS.as_bytes(),
        true,
    )
}

fn build_synthetic_pivot_package_with_malformed_cache_definition_rels() -> Vec<u8> {
    build_synthetic_pivot_package_with_overrides(
        VALID_WORKBOOK_RELS.as_bytes(),
        VALID_WORKSHEET_RELS.as_bytes(),
        b"<Relationships",
        false,
    )
}

fn build_synthetic_pivot_package_without_workbook_xml() -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    for (name, bytes) in [
        ("xl/worksheets/sheet1.xml", VALID_WORKSHEET_XML.as_bytes()),
        (
            "xl/worksheets/_rels/sheet1.xml.rels",
            VALID_WORKSHEET_RELS.as_bytes(),
        ),
        ("xl/pivotTables/pivotTable1.xml", VALID_PIVOT_TABLE_XML.as_bytes()),
        (
            "xl/pivotCache/pivotCacheDefinition1.xml",
            VALID_CACHE_DEFINITION_XML.as_bytes(),
        ),
        ("xl/pivotCache/pivotCacheRecords1.xml", VALID_CACHE_RECORDS_XML.as_bytes()),
    ] {
        zip.start_file(name, options).unwrap();
        zip.write_all(bytes).unwrap();
    }

    zip.finish().unwrap().into_inner()
}

fn build_synthetic_pivot_package_with_malformed_pivot_table_xml() -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("xl/pivotTables/pivotTable1.xml", options)
        .unwrap();
    zip.write_all(b"<pivotTableDefinition").unwrap(); // malformed / truncated XML

    zip.finish().unwrap().into_inner()
}

const VALID_WORKBOOK_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
  <pivotCaches>
    <pivotCache cacheId="1" r:id="rId2"/>
  </pivotCaches>
</workbook>"#;

const VALID_WORKBOOK_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/>
</Relationships>"#;

const VALID_WORKSHEET_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData/>
</worksheet>"#;

const VALID_WORKSHEET_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/>
</Relationships>"#;

const VALID_PIVOT_TABLE_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  name="PivotTable1" cacheId="1"/>"#;

const VALID_CACHE_DEFINITION_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" recordCount="0"/>"#;

const VALID_CACHE_DEFINITION_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords1.xml"/>
</Relationships>"#;

const VALID_CACHE_RECORDS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0"/>"#;

fn build_synthetic_pivot_package_with_overrides(
    workbook_rels: &[u8],
    worksheet_rels: &[u8],
    cache_definition_rels: &[u8],
    include_cache_records: bool,
) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    for (name, bytes) in [
        ("xl/workbook.xml", VALID_WORKBOOK_XML.as_bytes()),
        ("xl/_rels/workbook.xml.rels", workbook_rels),
        ("xl/worksheets/sheet1.xml", VALID_WORKSHEET_XML.as_bytes()),
        ("xl/worksheets/_rels/sheet1.xml.rels", worksheet_rels),
        ("xl/pivotTables/pivotTable1.xml", VALID_PIVOT_TABLE_XML.as_bytes()),
        (
            "xl/pivotCache/pivotCacheDefinition1.xml",
            VALID_CACHE_DEFINITION_XML.as_bytes(),
        ),
        (
            "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels",
            cache_definition_rels,
        ),
    ] {
        zip.start_file(name, options).unwrap();
        zip.write_all(bytes).unwrap();
    }

    if include_cache_records {
        zip.start_file("xl/pivotCache/pivotCacheRecords1.xml", options)
            .unwrap();
        zip.write_all(VALID_CACHE_RECORDS_XML.as_bytes()).unwrap();
    }

    zip.finish().unwrap().into_inner()
}
