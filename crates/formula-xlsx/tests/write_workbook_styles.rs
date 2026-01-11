use std::fs;
use std::io::{Cursor, Read};

use formula_model::{Cell, CellRef, CellValue, Font, Style, Workbook};
use tempfile::tempdir;
use zip::ZipArchive;

/// `write_workbook` is the low-level workbook serializer. Ensure it emits a valid `styles.xml`
/// with appended `xf` records and that cells reference the `xf` indices (not raw `style_id`s).
#[test]
fn write_workbook_emits_style_xfs_and_cell_s_mapping() -> Result<(), Box<dyn std::error::Error>> {
    let mut workbook = Workbook::new();
    let sheet_id = workbook.add_sheet("Sheet1")?;

    let italic_style_id = workbook.intern_style(Style {
        font: Some(Font {
            italic: true,
            ..Default::default()
        }),
        ..Default::default()
    });

    {
        let sheet = workbook.sheet_mut(sheet_id).unwrap();
        let mut cell = Cell::new(CellValue::Number(1.0));
        cell.style_id = italic_style_id;
        sheet.set_cell(CellRef::from_a1("A1")?, cell);
    }

    let dir = tempdir()?;
    let out_path = dir.path().join("styled.xlsx");
    formula_xlsx::write_workbook(&workbook, &out_path)?;

    // Reload via XlsxDocument and assert the style is present and applied.
    let doc = formula_xlsx::load_from_path(&out_path)?;
    let sheet_id = doc.workbook.sheets[0].id;
    let sheet = doc.workbook.sheet(sheet_id).unwrap();
    let cell = sheet.cell(CellRef::from_a1("A1")?).unwrap();
    let style = doc.workbook.styles.get(cell.style_id).unwrap();
    assert!(style.font.as_ref().is_some_and(|f| f.italic));

    // Inspect the raw parts to ensure the file is using `xf` indices and has >=2 xfs.
    let bytes = fs::read(&out_path)?;
    assert_sheet_has_style_index(&bytes, "xl/worksheets/sheet1.xml", r#"s="1""#)?;
    assert_cell_xfs_count_at_least(&bytes, "xl/styles.xml", 2)?;

    Ok(())
}

#[test]
fn write_workbook_errors_on_unknown_style_id() -> Result<(), Box<dyn std::error::Error>> {
    let mut workbook = Workbook::new();
    let sheet_id = workbook.add_sheet("Sheet1")?;
    {
        let sheet = workbook.sheet_mut(sheet_id).unwrap();
        let mut cell = Cell::new(CellValue::Number(1.0));
        cell.style_id = 999;
        sheet.set_cell(CellRef::from_a1("A1")?, cell);
    }

    let dir = tempdir()?;
    let out_path = dir.path().join("invalid.xlsx");
    let err = formula_xlsx::write_workbook(&workbook, &out_path).unwrap_err();
    match err {
        formula_xlsx::XlsxWriteError::Invalid(message) => {
            assert!(
                message.contains("unknown style_id"),
                "unexpected error: {message}"
            );
        }
        other => panic!("expected Invalid error, got {other:?}"),
    }

    Ok(())
}

fn assert_sheet_has_style_index(
    bytes: &[u8],
    sheet_part: &str,
    needle: &str,
) -> Result<(), Box<dyn std::error::Error>> {
    let mut archive = ZipArchive::new(Cursor::new(bytes))?;
    let mut xml = String::new();
    archive.by_name(sheet_part)?.read_to_string(&mut xml)?;
    assert!(
        xml.contains(needle),
        "expected {sheet_part} to contain {needle}, got:\n{xml}"
    );
    Ok(())
}

fn assert_cell_xfs_count_at_least(
    bytes: &[u8],
    styles_part: &str,
    min: u32,
) -> Result<(), Box<dyn std::error::Error>> {
    let mut archive = ZipArchive::new(Cursor::new(bytes))?;
    let mut xml = String::new();
    archive.by_name(styles_part)?.read_to_string(&mut xml)?;
    let doc = roxmltree::Document::parse(&xml)?;
    let cell_xfs = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "cellXfs")
        .ok_or_else(|| format!("{styles_part} missing <cellXfs>"))?;
    let count = cell_xfs
        .attribute("count")
        .and_then(|v| v.parse::<u32>().ok())
        .unwrap_or(0);
    assert!(
        count >= min,
        "expected {styles_part} cellXfs count >= {min}, got {count}"
    );
    Ok(())
}
