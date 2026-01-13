use formula_xlsx::{load_from_bytes, PivotCacheSourceType, XlsxPackage};
use std::io::Cursor;
use std::io::Write;
use zip::write::FileOptions;
use zip::ZipWriter;

const FIXTURE: &[u8] = include_bytes!("fixtures/pivot-fixture.xlsx");

#[test]
fn parses_pivot_cache_definition_worksheet_source_and_fields() {
    let pkg = XlsxPackage::from_bytes(FIXTURE).expect("read pkg");
    let defs = pkg
        .pivot_cache_definitions()
        .expect("parse pivot cache definitions");

    assert_eq!(defs.len(), 1);
    assert_eq!(defs[0].0, "xl/pivotCache/pivotCacheDefinition1.xml");

    let def = &defs[0].1;
    assert_eq!(def.record_count, Some(4));
    assert_eq!(def.refresh_on_load, Some(true));
    assert_eq!(def.cache_source_type, PivotCacheSourceType::Worksheet);
    assert_eq!(def.worksheet_source_sheet.as_deref(), Some("Sheet1"));
    assert_eq!(def.worksheet_source_ref.as_deref(), Some("A1:C5"));

    let names: Vec<&str> = def.cache_fields.iter().map(|f| f.name.as_str()).collect();
    assert_eq!(names, vec!["Region", "Product", "Sales"]);
}

#[test]
fn parses_pivot_cache_definition_from_document_parts() {
    let doc = load_from_bytes(FIXTURE).expect("load fixture");
    let defs = doc
        .pivot_cache_definitions()
        .expect("parse pivot cache definitions");

    assert_eq!(defs.len(), 1);
    let def = &defs[0].1;
    assert_eq!(def.cache_source_type, PivotCacheSourceType::Worksheet);
    assert_eq!(def.worksheet_source_sheet.as_deref(), Some("Sheet1"));
    assert_eq!(def.worksheet_source_ref.as_deref(), Some("A1:C5"));
}

#[test]
fn resolves_pivot_cache_definition_for_cache_id_by_filename() {
    let pkg = XlsxPackage::from_bytes(FIXTURE).expect("read pkg");
    let (part_name, def) = pkg
        .pivot_cache_definition_for_cache_id(1)
        .expect("resolve by cacheId")
        .expect("cache present");

    assert_eq!(part_name, "xl/pivotCache/pivotCacheDefinition1.xml");
    assert_eq!(def.cache_source_type, PivotCacheSourceType::Worksheet);
    assert_eq!(def.worksheet_source_sheet.as_deref(), Some("Sheet1"));
    assert_eq!(def.worksheet_source_ref.as_deref(), Some("A1:C5"));
}

#[test]
fn resolves_pivot_cache_definition_for_cache_id_via_workbook_relationships() {
    let bytes = build_synthetic_workbook_cache_id_package();
    let pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");

    let (part_name, def) = pkg
        .pivot_cache_definition_for_cache_id(7)
        .expect("resolve by cacheId")
        .expect("cache present");

    assert_eq!(part_name, "xl/pivotCache/pivotCacheDefinition1.xml");
    assert_eq!(def.cache_source_type, PivotCacheSourceType::Worksheet);
    assert_eq!(def.worksheet_source_sheet.as_deref(), Some("Sheet1"));
    assert_eq!(def.worksheet_source_ref.as_deref(), Some("A1:C5"));
}

#[test]
fn pivot_cache_definition_for_cache_id_returns_none_when_workbook_is_missing() {
    let bytes = build_synthetic_package_without_workbook();
    let pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");

    // This method is intentionally tolerant: if workbook.xml is missing, we don't attempt to
    // guess the definition part name.
    assert_eq!(
        pkg.pivot_cache_definition_for_cache_id(7)
            .expect("resolve by cacheId"),
        None
    );
}

fn build_synthetic_workbook_cache_id_package() -> Vec<u8> {
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <pivotCaches>
    <pivotCache cacheId="7" r:id="rId1"/>
  </pivotCaches>
</workbook>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/>
</Relationships>"#;

    let cache_definition_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" refreshOnLoad="1" recordCount="4">
  <cacheSource type="worksheet">
    <worksheetSource ref="A1:C5" sheet="Sheet1"/>
  </cacheSource>
  <cacheFields count="0"/>
</pivotCacheDefinition>"#;

    let misleading_guess_definition_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cacheSource type="worksheet">
    <worksheetSource ref="Z1:Z2" sheet="WrongSheet"/>
  </cacheSource>
  <cacheFields count="0"/>
</pivotCacheDefinition>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    for (name, xml) in [
        ("xl/workbook.xml", workbook_xml),
        ("xl/_rels/workbook.xml.rels", workbook_rels),
        (
            "xl/pivotCache/pivotCacheDefinition1.xml",
            cache_definition_xml,
        ),
        // A misleading `pivotCacheDefinition{cacheId}.xml` part that should be ignored in favor of
        // the workbook mapping above.
        (
            "xl/pivotCache/pivotCacheDefinition7.xml",
            misleading_guess_definition_xml,
        ),
    ] {
        zip.start_file(name, options).unwrap();
        zip.write_all(xml.as_bytes()).unwrap();
    }

    zip.finish().unwrap().into_inner()
}

fn build_synthetic_package_without_workbook() -> Vec<u8> {
    let cache_definition_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cacheSource type="worksheet">
    <worksheetSource ref="A1:C5" sheet="Sheet1"/>
  </cacheSource>
  <cacheFields count="0"/>
</pivotCacheDefinition>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("xl/pivotCache/pivotCacheDefinition7.xml", options)
        .unwrap();
    zip.write_all(cache_definition_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}
