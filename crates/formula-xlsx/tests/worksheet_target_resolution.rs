use std::io::{Cursor, Read, Write};

use formula_model::{CellRef, CellValue};
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

/// `workbook.xml.rels` targets are resolved relative to `xl/workbook.xml` and may contain `.`/`..`
/// segments. Ensure `load_from_bytes` normalizes these correctly when locating worksheet parts.
#[test]
fn xlsx_document_resolves_worksheet_targets_with_dot_segments() -> Result<(), Box<dyn std::error::Error>>
{
    let fixture = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/xlsx/styles/styles.xlsx");
    let fixture_bytes = std::fs::read(fixture)?;

    // Rewrite the fixture zip so the sheet relationship target includes a redundant `./` segment.
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
        let name = file.name().to_string();

        if name == "xl/_rels/workbook.xml.rels" {
            let xml = String::from_utf8(buf)?;
            let updated = xml.replace(
                r#"Target="worksheets/sheet1.xml""#,
                r#"Target="./worksheets/sheet1.xml""#,
            );
            buf = updated.into_bytes();
        }

        zip.start_file(name, options)?;
        zip.write_all(&buf)?;
    }
    zip.finish()?;

    let modified_bytes = out.into_inner();
    let doc = formula_xlsx::load_from_bytes(&modified_bytes)?;

    let sheet_id = doc.workbook.sheets[0].id;
    let sheet = doc.workbook.sheet(sheet_id).unwrap();
    assert_eq!(
        sheet.value(CellRef::from_a1("A1")?),
        CellValue::String("Bold".to_string())
    );

    Ok(())
}

