use std::io::{Cursor, Read, Write};

use formula_model::{Style, Workbook};
use formula_xlsx::{load_from_bytes, XlsxDocument};
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

fn zip_part(zip_bytes: &[u8], name: &str) -> Vec<u8> {
    let mut archive = ZipArchive::new(Cursor::new(zip_bytes)).expect("open zip");
    let mut buf = Vec::new();
    archive
        .by_name(name)
        .unwrap_or_else(|_| panic!("missing zip part: {name}"))
        .read_to_end(&mut buf)
        .expect("read zip part");
    buf
}

fn zip_replace_part(zip_bytes: &[u8], name: &str, new_bytes: &[u8]) -> Vec<u8> {
    let mut archive = ZipArchive::new(Cursor::new(zip_bytes)).expect("open zip");

    let cursor = Cursor::new(Vec::new());
    let mut writer = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);

    for i in 0..archive.len() {
        let mut file = archive.by_index(i).expect("zip by_index");
        let entry_name = file.name().to_string();
        if file.is_dir() {
            writer
                .add_directory(entry_name, options)
                .expect("write dir entry");
            continue;
        }

        let mut buf = Vec::new();
        file.read_to_end(&mut buf).expect("read zip entry");
        writer
            .start_file(entry_name.as_str(), options)
            .expect("start zip entry");
        if entry_name == name {
            writer.write_all(new_bytes).expect("write replacement");
        } else {
            writer.write_all(&buf).expect("write original");
        }
    }

    writer.finish().expect("finish zip").into_inner()
}

#[test]
fn xlsx_document_preserves_unknown_col_attrs_when_col_style_is_unchanged(
) -> Result<(), Box<dyn std::error::Error>> {
    // Create a workbook with a default column style on column B.
    let mut workbook = Workbook::new();
    let sheet_id = workbook.add_sheet("Sheet1")?;

    let style_id = workbook.intern_style(Style {
        number_format: Some("0.00".to_string()),
        ..Default::default()
    });
    assert_ne!(style_id, 0, "expected non-default style id");

    workbook
        .sheet_mut(sheet_id)
        .expect("sheet exists")
        .set_col_style_id(1, Some(style_id)); // col B

    let bytes = XlsxDocument::new(workbook).save_to_vec()?;

    // Inject an attribute that `formula_model` does not currently model (`bestFit`).
    let sheet_xml = zip_part(&bytes, "xl/worksheets/sheet1.xml");
    let sheet_xml_str = std::str::from_utf8(&sheet_xml)?;
    let needle = r#"<col min="2" max="2""#;
    assert!(
        sheet_xml_str.contains(needle),
        "expected generated sheet XML to include col B definition; got: {sheet_xml_str}"
    );
    let patched_xml =
        sheet_xml_str.replacen(needle, r#"<col min="2" max="2" bestFit="1""#, 1);
    assert!(
        patched_xml.contains(r#"bestFit="1""#),
        "expected patched sheet XML to include bestFit=1"
    );

    let patched_bytes =
        zip_replace_part(&bytes, "xl/worksheets/sheet1.xml", patched_xml.as_bytes());

    // Load + save without changing the model. The writer should avoid rewriting `<cols>` (and thus
    // preserve the unknown attribute) when column defaults are unchanged.
    let doc = load_from_bytes(&patched_bytes)?;
    let out = doc.save_to_vec()?;

    let out_sheet_xml = zip_part(&out, "xl/worksheets/sheet1.xml");
    let out_sheet_xml_str = std::str::from_utf8(&out_sheet_xml)?;
    assert!(
        out_sheet_xml_str.contains(r#"bestFit="1""#),
        "expected bestFit=1 to be preserved, got: {out_sheet_xml_str}"
    );

    Ok(())
}

