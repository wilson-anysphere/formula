use std::path::Path;

use formula_model::CellRef;
use formula_xlsx::load_from_bytes;

#[test]
fn rich_value_index_image_in_cell_fixture_maps_expected_vm_cells() -> Result<(), Box<dyn std::error::Error>>
{
    let fixture_path =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx/basic/image-in-cell.xlsx");
    let fixture_bytes = std::fs::read(&fixture_path)?;

    let doc = load_from_bytes(&fixture_bytes)?;
    let sheet_id = doc.workbook.sheets[0].id;

    assert_eq!(
        doc.rich_value_index(sheet_id, CellRef::from_a1("B2")?),
        Some(0)
    );
    assert_eq!(
        doc.rich_value_index(sheet_id, CellRef::from_a1("B3")?),
        Some(0)
    );
    assert_eq!(
        doc.rich_value_index(sheet_id, CellRef::from_a1("B4")?),
        Some(1)
    );

    Ok(())
}

#[test]
fn rich_value_index_image_in_cell_richdata_fixture_supports_vm_zero_based() -> Result<(), Box<dyn std::error::Error>>
{
    let fixture_path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/xlsx/basic/image-in-cell-richdata.xlsx");
    let fixture_bytes = std::fs::read(&fixture_path)?;

    let doc = load_from_bytes(&fixture_bytes)?;
    let sheet_id = doc.workbook.sheets[0].id;

    assert_eq!(
        doc.rich_value_index(sheet_id, CellRef::from_a1("A1")?),
        Some(0)
    );

    Ok(())
}

