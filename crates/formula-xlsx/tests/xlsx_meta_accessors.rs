use std::path::Path;

use formula_model::CellRef;

use formula_xlsx::CellValueKind;

#[test]
fn xlsx_document_exposes_cell_meta() -> Result<(), Box<dyn std::error::Error>> {
    let fixture = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/xlsx/metadata/date-iso-cell.xlsx");

    let doc = formula_xlsx::load_from_path(&fixture)?;
    let sheet_id = doc.workbook.sheets[0].id;
    let cell_ref = CellRef::from_a1("A1")?;

    let meta = doc
        .cell_meta(sheet_id, cell_ref)
        .expect("expected cell metadata for A1");

    assert!(
        matches!(meta.value_kind, Some(CellValueKind::Other { ref t }) if t == "d"),
        "expected A1 to preserve an unknown cell type (t=\"d\") via CellMeta, got: {meta:?}"
    );

    // Ensure the full metadata structure is also accessible via the general accessor.
    assert!(
        doc.xlsx_meta().cell_meta.get(&(sheet_id, cell_ref)).is_some(),
        "expected xlsx_meta().cell_meta to contain an entry for A1"
    );

    Ok(())
}

#[test]
fn xlsx_document_exposes_vm_metadata() -> Result<(), Box<dyn std::error::Error>> {
    let fixture = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/xlsx/metadata/rich-values-vm.xlsx");

    let doc = formula_xlsx::load_from_path(&fixture)?;
    let sheet_id = doc.workbook.sheets[0].id;

    let a1 = CellRef::from_a1("A1")?;
    let b1 = CellRef::from_a1("B1")?;

    let a1_meta = doc
        .cell_meta(sheet_id, a1)
        .expect("expected cell metadata for A1");
    assert_eq!(
        a1_meta.vm.as_deref(),
        Some("1"),
        "expected A1 to preserve vm attribute via CellMeta, got: {a1_meta:?}"
    );

    let b1_meta = doc
        .cell_meta(sheet_id, b1)
        .expect("expected cell metadata for B1");
    assert_eq!(b1_meta.vm.as_deref(), None);

    Ok(())
}

#[test]
fn xlsx_document_exposes_vm_cm_metadata() -> Result<(), Box<dyn std::error::Error>> {
    let fixture = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/xlsx/rich-data/images-in-cell.xlsx");

    let doc = formula_xlsx::load_from_path(&fixture)?;
    let sheet_id = doc.workbook.sheets[0].id;

    let a1 = CellRef::from_a1("A1")?;
    let meta = doc
        .cell_meta(sheet_id, a1)
        .expect("expected cell metadata for A1");

    assert_eq!(meta.vm.as_deref(), Some("1"));
    assert_eq!(meta.cm.as_deref(), Some("1"));

    Ok(())
}
