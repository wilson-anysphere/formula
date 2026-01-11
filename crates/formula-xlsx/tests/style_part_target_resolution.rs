use std::io::{Cursor, Read, Write};

use formula_model::{Cell, CellRef, Font, Style};
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

/// Ensure `XlsxDocument` respects `workbook.xml.rels` targets for the styles part rather than
/// assuming it always lives at `xl/styles.xml`.
#[test]
fn xlsx_document_resolves_styles_part_from_workbook_rels() -> Result<(), Box<dyn std::error::Error>> {
    let fixture = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/xlsx/styles/styles.xlsx");
    let fixture_bytes = std::fs::read(fixture)?;

    // Rewrite the fixture zip so styles.xml is stored at a non-default path.
    let mut archive = ZipArchive::new(Cursor::new(&fixture_bytes))?;
    let mut out = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(&mut out);
    let options = FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);

    for i in 0..archive.len() {
        let mut file = archive.by_index(i)?;
        if file.is_dir() {
            continue;
        }
        let mut buf = Vec::new();
        file.read_to_end(&mut buf)?;
        let mut name = file.name().to_string();

        if name == "xl/styles.xml" {
            name = "xl/customStyles.xml".to_string();
        } else if name == "xl/_rels/workbook.xml.rels" {
            let xml = String::from_utf8(buf)?;
            let updated = xml.replace(r#"Target="styles.xml""#, r#"Target="customStyles.xml""#);
            buf = updated.into_bytes();
        } else if name == "[Content_Types].xml" {
            let xml = String::from_utf8(buf)?;
            let updated = xml.replace(
                r#"PartName="/xl/styles.xml""#,
                r#"PartName="/xl/customStyles.xml""#,
            );
            buf = updated.into_bytes();
        }

        zip.start_file(name, options)?;
        zip.write_all(&buf)?;
    }
    zip.finish()?;

    let modified_bytes = out.into_inner();
    let mut doc = formula_xlsx::load_from_bytes(&modified_bytes)?;

    // Apply a newly-interned style.
    let italic_style_id = doc.workbook.intern_style(Style {
        font: Some(Font {
            italic: true,
            ..Default::default()
        }),
        ..Default::default()
    });

    let sheet_id = doc.workbook.sheets[0].id;
    let sheet = doc.workbook.sheet_mut(sheet_id).unwrap();
    let cell_ref = CellRef::from_a1("A1")?;
    let mut cell = sheet.cell(cell_ref).cloned().unwrap_or_else(Cell::default);
    cell.style_id = italic_style_id;
    sheet.set_cell(cell_ref, cell);

    let saved = doc.save_to_vec()?;

    // The output should update the custom styles part and not synthesize `xl/styles.xml`.
    let mut saved_zip = ZipArchive::new(Cursor::new(&saved))?;
    assert!(saved_zip.by_name("xl/customStyles.xml").is_ok());
    assert!(saved_zip.by_name("xl/styles.xml").is_err());

    let mut rels_xml = String::new();
    saved_zip
        .by_name("xl/_rels/workbook.xml.rels")?
        .read_to_string(&mut rels_xml)?;
    assert!(
        rels_xml.contains(r#"Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="customStyles.xml""#),
        "workbook.xml.rels should still point at customStyles.xml"
    );

    // Reload and ensure the style survived.
    let reloaded = formula_xlsx::load_from_bytes(&saved)?;
    let sheet_id = reloaded.workbook.sheets[0].id;
    let sheet = reloaded.workbook.sheet(sheet_id).unwrap();
    let cell = sheet.cell(cell_ref).unwrap();
    let style = reloaded.workbook.styles.get(cell.style_id).unwrap();
    assert!(style.font.as_ref().is_some_and(|f| f.italic));

    Ok(())
}

