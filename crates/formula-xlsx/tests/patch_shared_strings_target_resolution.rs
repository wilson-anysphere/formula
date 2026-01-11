use std::io::{Cursor, Read, Write};

use formula_model::{CellRef, CellValue};
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

use formula_xlsx::{CellPatch, WorkbookCellPatches, XlsxPackage};

/// Ensure the patch pipeline respects `workbook.xml.rels` targets for the shared strings part
/// rather than assuming it always lives at `xl/sharedStrings.xml`.
#[test]
fn patch_pipeline_resolves_shared_strings_part_from_workbook_rels(
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
            let updated = xml.replace(
                r#"Target="sharedStrings.xml""#,
                r#"Target="customSharedStrings.xml""#,
            );
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
    let mut pkg = XlsxPackage::from_bytes(&modified_bytes)?;
    let sheet_name = pkg
        .workbook_sheets()?
        .first()
        .ok_or("fixture should contain at least one worksheet")?
        .name
        .clone();

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        sheet_name,
        CellRef::from_a1("A2")?,
        CellPatch::set_value(CellValue::String("Patched".to_string())),
    );
    pkg.apply_cell_patches(&patches)?;
    let saved = pkg.write_to_bytes()?;

    let mut saved_zip = ZipArchive::new(Cursor::new(&saved))?;
    assert!(saved_zip.by_name("xl/customSharedStrings.xml").is_ok());
    assert!(saved_zip.by_name("xl/sharedStrings.xml").is_err());

    let mut ss_xml = String::new();
    saved_zip
        .by_name("xl/customSharedStrings.xml")?
        .read_to_string(&mut ss_xml)?;
    assert!(
        ss_xml.contains("Patched"),
        "custom shared strings part should include newly inserted string"
    );

    // Ensure the patched cell still uses shared strings (`t="s"`) instead of being rewritten to
    // an inline string due to missing shared string resolution.
    let mut sheet_xml = String::new();
    saved_zip
        .by_name("xl/worksheets/sheet1.xml")?
        .read_to_string(&mut sheet_xml)?;
    assert_cell_type(&sheet_xml, "A2", Some("s"))?;

    let mut rels_xml = String::new();
    saved_zip
        .by_name("xl/_rels/workbook.xml.rels")?
        .read_to_string(&mut rels_xml)?;
    assert!(
        rels_xml.contains(
            r#"Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="customSharedStrings.xml""#
        ),
        "workbook.xml.rels should still point at customSharedStrings.xml"
    );

    Ok(())
}

fn assert_cell_type(
    worksheet_xml: &str,
    a1: &str,
    expected_type: Option<&str>,
) -> Result<(), Box<dyn std::error::Error>> {
    let mut reader = quick_xml::Reader::from_str(worksheet_xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            quick_xml::events::Event::Start(e) | quick_xml::events::Event::Empty(e)
                if e.name().as_ref() == b"c" =>
            {
                let mut r = None;
                let mut t = None;
                for attr in e.attributes() {
                    let attr = attr?;
                    match attr.key.as_ref() {
                        b"r" => r = Some(attr.unescape_value()?.into_owned()),
                        b"t" => t = Some(attr.unescape_value()?.into_owned()),
                        _ => {}
                    }
                }
                if r.as_deref() == Some(a1) {
                    assert_eq!(t.as_deref(), expected_type);
                    return Ok(());
                }
            }
            quick_xml::events::Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Err(format!("cell {a1} not found in worksheet").into())
}
