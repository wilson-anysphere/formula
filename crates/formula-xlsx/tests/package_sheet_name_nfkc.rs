use formula_model::{CellRef, CellValue};
use formula_xlsx::{load_from_bytes, PackageCellPatch, XlsxPackage};
use rust_xlsxwriter::Workbook;

#[test]
fn package_cell_patches_match_sheet_names_nfkc_case_insensitively(
) -> Result<(), Box<dyn std::error::Error>> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.set_name("Kelvin")?;
    let bytes = workbook.save_to_buffer()?;

    let pkg = XlsxPackage::from_bytes(&bytes)?;
    let cell = CellRef::from_a1("A1")?;
    let patch = PackageCellPatch::for_sheet_name(
        // U+212A KELVIN SIGN (K) is NFKC-equivalent to ASCII 'K'.
        "Kelvin",
        cell,
        CellValue::String("patched".to_string()),
        None,
    );
    let out_bytes = pkg.apply_cell_patches_to_bytes(&[patch])?;

    let doc = load_from_bytes(&out_bytes)?;
    let sheet = doc.workbook.sheet_by_name("Kelvin").expect("sheet exists");
    assert_eq!(sheet.value(cell), CellValue::String("patched".to_string()));
    Ok(())
}

