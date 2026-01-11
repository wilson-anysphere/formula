//! Compile-time checks for native-only APIs that are cfg-gated out of wasm builds.

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn native_only_exports_still_exist() -> Result<(), Box<dyn std::error::Error>> {
    use std::io::Cursor;
    use std::path::Path;

    use tempfile::tempdir;

    let missing = Path::new("definitely-does-not-exist.xlsx");
    assert!(formula_xlsx::load_from_path(missing).is_err());
    assert!(formula_xlsx::read_workbook(missing).is_err());

    let mut workbook = formula_model::Workbook::new();
    workbook.add_sheet("Sheet1")?;

    // Writer helper.
    let mut buffer = Cursor::new(Vec::new());
    formula_xlsx::write_workbook_to_writer(&workbook, &mut buffer)?;

    // Reader helper.
    let workbook_roundtrip = formula_xlsx::read_workbook_from_reader(Cursor::new(buffer.into_inner()))?;
    assert_eq!(workbook_roundtrip.sheets.len(), 1);

    // Disk-based read/write.
    let dir = tempdir()?;
    let input_path = dir.path().join("workbook.xlsx");
    formula_xlsx::write_workbook(&workbook, &input_path)?;
    let _ = formula_xlsx::read_workbook(&input_path)?;

    // Shared strings helpers that operate on `.xlsx` files on disk.
    let _ = formula_xlsx::shared_strings::read_shared_strings_from_xlsx(&input_path);
    let output_path = dir.path().join("with-shared-strings.xlsx");
    formula_xlsx::shared_strings::write_shared_strings_to_xlsx(
        &input_path,
        &output_path,
        &formula_xlsx::shared_strings::SharedStrings::default(),
    )?;

    // WorkbookPackage has disk-based helpers on native targets.
    let mut package = formula_xlsx::WorkbookPackage::load(&input_path)?;
    let pkg_out = dir.path().join("package-out.xlsx");
    package.save(&pkg_out)?;

    Ok(())
}
