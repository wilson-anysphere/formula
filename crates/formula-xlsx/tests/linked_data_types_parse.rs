use std::path::Path;

use formula_model::{CellRef, CellValue};

#[test]
fn linked_data_types_fixture_loads_cells_as_entities() -> Result<(), Box<dyn std::error::Error>> {
    let fixture_path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/xlsx/richdata/linked-data-types.xlsx");
    let fixture_bytes = std::fs::read(&fixture_path)?;

    let doc = formula_xlsx::load_from_bytes(&fixture_bytes)?;
    let sheet_id = doc.workbook.sheets[0].id;
    let sheet = doc.workbook.sheet(sheet_id).expect("Sheet1 exists");

    let a1 = sheet.value(CellRef::from_a1("A1")?);
    match a1 {
        CellValue::Entity(entity) => {
            assert_eq!(entity.display_value, "MSFT");
            assert!(
                entity.entity_type.contains("stocks"),
                "expected stock entity type, got {}",
                entity.entity_type
            );
        }
        other => panic!("expected A1 to be an entity, got {other:?}"),
    }

    let a2 = sheet.value(CellRef::from_a1("A2")?);
    match a2 {
        CellValue::Entity(entity) => {
            assert_eq!(entity.display_value, "Seattle");
            assert!(
                entity.entity_type.contains("geography"),
                "expected geography entity type, got {}",
                entity.entity_type
            );
        }
        other => panic!("expected A2 to be an entity, got {other:?}"),
    }

    Ok(())
}

