use std::path::Path;

#[test]
fn load_from_bytes_populates_worksheet_tables() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/table.xlsx");
    let bytes = std::fs::read(&path).expect("read table fixture");
    let doc = formula_xlsx::load_from_bytes(&bytes).expect("load_from_bytes should succeed");

    assert_eq!(doc.workbook.sheets.len(), 1);
    assert_eq!(doc.workbook.sheets[0].tables.len(), 1);

    let table = &doc.workbook.sheets[0].tables[0];
    assert_eq!(table.name, "Table1");
    assert_eq!(table.relationship_id.as_deref(), Some("rId1"));
    assert_eq!(table.part_path.as_deref(), Some("xl/tables/table1.xml"));
}

