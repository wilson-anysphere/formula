use formula_xlsx::{PivotCacheValue, XlsxPackage};
use std::io::{Cursor, Write};
use zip::write::FileOptions;
use zip::ZipWriter;

fn build_pkg_with_cache_definition(xml: &str) -> XlsxPackage {
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("xl/pivotCache/pivotCacheDefinition1.xml", options)
        .unwrap();
    zip.write_all(xml.as_bytes()).unwrap();

    let bytes = zip.finish().unwrap().into_inner();
    XlsxPackage::from_bytes(&bytes).unwrap()
}

#[test]
fn parses_shared_items_for_each_cache_field() {
    let xml = r##"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cacheFields count="2">
    <cacheField name="Field1">
      <sharedItems count="8">
        <m/>
        <n v="1"/>
        <s v="East"/>
        <b v="0"/>
        <e v="#DIV/0!"/>
        <d v="2024-01-15T00:00:00Z"/>
        <n><v>42</v></n>
        <x v="3"/>
      </sharedItems>
    </cacheField>
    <cacheField name="Field2">
      <sharedItems>
        <s>West</s>
        <b>1</b>
        <n>3.5</n>
        <m></m>
      </sharedItems>
    </cacheField>
  </cacheFields>
</pivotCacheDefinition>"##;

    let pkg = build_pkg_with_cache_definition(xml);
    let def = pkg
        .pivot_cache_definition("xl/pivotCache/pivotCacheDefinition1.xml")
        .unwrap()
        .unwrap();

    assert_eq!(def.cache_fields.len(), 2);
    assert_eq!(def.cache_fields[0].name, "Field1");
    assert_eq!(
        def.cache_fields[0].shared_items.as_ref().unwrap(),
        &vec![
            PivotCacheValue::Missing,
            PivotCacheValue::Number(1.0),
            PivotCacheValue::String("East".to_string()),
            PivotCacheValue::Bool(false),
            PivotCacheValue::Error("#DIV/0!".to_string()),
            PivotCacheValue::DateTime("2024-01-15T00:00:00Z".to_string()),
            PivotCacheValue::Number(42.0),
        ],
    );

    assert_eq!(def.cache_fields[1].name, "Field2");
    assert_eq!(
        def.cache_fields[1].shared_items.as_ref().unwrap(),
        &vec![
            PivotCacheValue::String("West".to_string()),
            PivotCacheValue::Bool(true),
            PivotCacheValue::Number(3.5),
            PivotCacheValue::Missing,
        ],
    );
}

#[test]
fn parses_namespace_prefixed_shared_items() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pc:pivotCacheDefinition xmlns:pc="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <pc:cacheFields pc:count="2">
    <pc:cacheField pc:name="Field1">
      <pc:sharedItems pc:count="4">
        <pc:s>East</pc:s>
        <pc:n pc:v="2"/>
        <pc:foo pc:v="ignore-me"/>
        <pc:d pc:v="2024-01-15T00:00:00Z"/>
      </pc:sharedItems>
    </pc:cacheField>
    <pc:cacheField pc:name="Field2">
      <pc:sharedItems>
        <pc:n><pc:v>10</pc:v></pc:n>
        <pc:b pc:v="0"/>
      </pc:sharedItems>
    </pc:cacheField>
  </pc:cacheFields>
</pc:pivotCacheDefinition>"#;

    let pkg = build_pkg_with_cache_definition(xml);
    let def = pkg
        .pivot_cache_definition("xl/pivotCache/pivotCacheDefinition1.xml")
        .unwrap()
        .unwrap();

    assert_eq!(def.cache_fields.len(), 2);
    assert_eq!(def.cache_fields[0].name, "Field1");
    assert_eq!(
        def.cache_fields[0].shared_items.as_ref().unwrap(),
        &vec![
            PivotCacheValue::String("East".to_string()),
            PivotCacheValue::Number(2.0),
            PivotCacheValue::DateTime("2024-01-15T00:00:00Z".to_string()),
        ],
    );

    assert_eq!(def.cache_fields[1].name, "Field2");
    assert_eq!(
        def.cache_fields[1].shared_items.as_ref().unwrap(),
        &vec![PivotCacheValue::Number(10.0), PivotCacheValue::Bool(false)],
    );
}
