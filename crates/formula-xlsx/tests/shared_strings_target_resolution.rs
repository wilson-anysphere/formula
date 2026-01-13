use std::io::{Cursor, Read, Write};

use formula_model::CellRef;
use formula_model::CellValue;
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

/// Ensure `XlsxDocument` respects `workbook.xml.rels` targets for the shared strings part rather
/// than assuming it always lives at `xl/sharedStrings.xml`.
#[test]
fn xlsx_document_resolves_shared_strings_part_from_workbook_rels(
) -> Result<(), Box<dyn std::error::Error>> {
    let fixture = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/xlsx/styles/rich-text-shared-strings.xlsx");
    let fixture_bytes = std::fs::read(fixture)?;

    // Rewrite the fixture zip so sharedStrings.xml is stored at a non-default path.
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

        if name == "xl/sharedStrings.xml" {
            name = "xl/customSharedStrings.xml".to_string();
        } else if name == "xl/_rels/workbook.xml.rels" {
            let xml = String::from_utf8(buf)?;
            let updated =
                xml.replace(r#"Target="sharedStrings.xml""#, r#"Target="customSharedStrings.xml""#);
            buf = updated.into_bytes();
        } else if name == "[Content_Types].xml" {
            let xml = String::from_utf8(buf)?;
            let updated = xml.replace(
                r#"PartName="/xl/sharedStrings.xml""#,
                r#"PartName="/xl/customSharedStrings.xml""#,
            );
            buf = updated.into_bytes();
        }

        zip.start_file(name, options)?;
        zip.write_all(&buf)?;
    }
    zip.finish()?;

    let modified_bytes = out.into_inner();
    let doc = formula_xlsx::load_from_bytes(&modified_bytes)?;

    // Ensure we actually loaded the shared strings.
    let sheet_id = doc.workbook.sheets[0].id;
    let sheet = doc.workbook.sheet(sheet_id).unwrap();
    match sheet.value(CellRef::from_a1("A1")?) {
        CellValue::RichText(rich) => {
            assert_eq!(rich.text, "Hello Bold Italic");
            assert!(
                !rich.runs.is_empty(),
                "expected rich text runs to be preserved for shared string"
            );
        }
        other => panic!("expected A1 to be rich text, got {other:?}"),
    }
    assert_eq!(
        sheet.value(CellRef::from_a1("A2")?),
        formula_model::CellValue::String("Plain".to_string())
    );

    let saved = doc.save_to_vec()?;

    // The output should update the custom shared strings part and not synthesize `xl/sharedStrings.xml`.
    let mut saved_zip = ZipArchive::new(Cursor::new(&saved))?;
    assert!(saved_zip.by_name("xl/customSharedStrings.xml").is_ok());
    assert!(saved_zip.by_name("xl/sharedStrings.xml").is_err());

    let mut rels_xml = String::new();
    saved_zip
        .by_name("xl/_rels/workbook.xml.rels")?
        .read_to_string(&mut rels_xml)?;
    assert!(
        rels_xml.contains(r#"Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="customSharedStrings.xml""#),
        "workbook.xml.rels should still point at customSharedStrings.xml"
    );

    Ok(())
}
