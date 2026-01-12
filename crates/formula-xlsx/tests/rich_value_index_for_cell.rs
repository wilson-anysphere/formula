use std::path::Path;

use formula_model::CellRef;
use formula_xlsx::load_from_bytes;

#[test]
fn resolves_rich_value_index_for_vm_cell() -> Result<(), Box<dyn std::error::Error>> {
    let fixture_path =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx/metadata/rich-values-vm.xlsx");
    let fixture_bytes = std::fs::read(&fixture_path)?;

    let doc = load_from_bytes(&fixture_bytes)?;
    let sheet_id = doc.workbook.sheets[0].id;

    assert_eq!(
        doc.rich_value_index_for_cell(sheet_id, CellRef::from_a1("A1")?)?,
        Some(0)
    );

    // Cells without `c/@vm` should return `None`.
    assert_eq!(
        doc.rich_value_index_for_cell(sheet_id, CellRef::from_a1("B1")?)?,
        None
    );

    Ok(())
}

