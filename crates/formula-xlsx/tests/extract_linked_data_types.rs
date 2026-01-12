use std::path::Path;

use formula_model::CellRef;
use formula_xlsx::rich_data::extract_linked_data_types;
use formula_xlsx::XlsxPackage;

#[test]
fn extracts_linked_data_types_fixture() -> Result<(), Box<dyn std::error::Error>> {
    let fixture_path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/xlsx/richdata/linked-data-types.xlsx");
    let fixture_bytes = std::fs::read(&fixture_path)?;

    let pkg = XlsxPackage::from_bytes(&fixture_bytes)?;
    let extracted = extract_linked_data_types(&pkg)?;

    let a1 = extracted
        .get(&("Sheet1".to_string(), CellRef::from_a1("A1")?))
        .ok_or("missing Sheet1!A1 rich value")?;
    assert_eq!(a1.type_name.as_deref(), Some("com.microsoft.excel.stocks"));
    assert_eq!(a1.display.as_deref(), Some("MSFT"));

    let a2 = extracted
        .get(&("Sheet1".to_string(), CellRef::from_a1("A2")?))
        .ok_or("missing Sheet1!A2 rich value")?;
    assert_eq!(a2.type_name.as_deref(), Some("com.microsoft.excel.geography"));
    assert_eq!(a2.display.as_deref(), Some("Seattle"));

    Ok(())
}

