use formula_xlsx::{PivotCacheSourceType, XlsxPackage};

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

