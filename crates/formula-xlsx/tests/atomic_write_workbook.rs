use std::io;

use formula_model::{Cell, CellRef, CellValue, Workbook};

#[test]
fn write_workbook_creates_parent_directories() -> Result<(), Box<dyn std::error::Error>> {
    let mut workbook = Workbook::new();
    workbook.add_sheet("Sheet1")?;

    let tmp = tempfile::tempdir()?;
    let out_path = tmp.path().join("nested/dir/workbook.xlsx");
    assert!(
        !out_path.parent().unwrap().exists(),
        "test precondition: parent dir should not exist"
    );

    formula_xlsx::write_workbook(&workbook, &out_path)?;

    assert!(out_path.is_file(), "expected workbook file to be created");
    assert!(
        out_path.parent().unwrap().exists(),
        "expected parent directories to be created"
    );

    Ok(())
}

#[test]
fn write_workbook_does_not_clobber_existing_file_on_error() -> Result<(), Box<dyn std::error::Error>>
{
    let mut workbook = Workbook::new();
    let sheet_id = workbook.add_sheet("Sheet1")?;
    {
        let sheet = workbook.sheet_mut(sheet_id).unwrap();
        let mut cell = Cell::new(CellValue::Number(1.0));
        cell.style_id = 999; // invalid / unknown
        sheet.set_cell(CellRef::from_a1("A1")?, cell);
    }

    let tmp = tempfile::tempdir()?;
    let out_path = tmp.path().join("out.xlsx");
    let sentinel = b"sentinel-bytes";
    std::fs::write(&out_path, sentinel)?;

    let err = formula_xlsx::write_workbook(&workbook, &out_path).unwrap_err();
    match err {
        formula_xlsx::XlsxWriteError::Invalid(_) => {}
        other => return Err(Box::new(io::Error::new(io::ErrorKind::Other, format!("{other:?}")))),
    }

    let written = std::fs::read(&out_path)?;
    assert_eq!(
        written, sentinel,
        "destination file should remain untouched on error"
    );

    let entries: Vec<_> = std::fs::read_dir(tmp.path())?
        .collect::<Result<Vec<_>, _>>()?
        .into_iter()
        .map(|e| e.path())
        .filter(|p| p.is_file())
        .collect();
    assert_eq!(
        entries,
        vec![out_path.clone()],
        "expected no temp files to remain after failure"
    );

    Ok(())
}

