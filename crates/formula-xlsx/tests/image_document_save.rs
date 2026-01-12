use std::io::{Cursor, Read as _};

use formula_model::drawings::ImageId;
use formula_model::{CellRef, CellValue, ImageValue};
use formula_xlsx::load_from_bytes;
use rust_xlsxwriter::Workbook;
use zip::ZipArchive;

fn zip_part_to_string(bytes: &[u8], name: &str) -> String {
    let cursor = Cursor::new(bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut file = archive.by_name(name).expect("part exists");
    let mut out = String::new();
    file.read_to_string(&mut out).expect("read xml");
    out
}

fn shared_strings_as_vec(shared_strings_xml: &str) -> Vec<String> {
    let doc = roxmltree::Document::parse(shared_strings_xml).expect("parse sharedStrings.xml");
    doc.descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "si")
        .map(|si| {
            si.descendants()
                .filter(|n| n.is_element() && n.tag_name().name() == "t")
                .filter_map(|n| n.text())
                .collect::<String>()
        })
        .collect()
}

fn cell_shared_string_index(sheet_xml: &str, cell_ref: &str) -> Option<u32> {
    let doc = roxmltree::Document::parse(sheet_xml).ok()?;
    let cell = doc.descendants().find(|n| {
        n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some(cell_ref)
    })?;
    if cell.attribute("t") != Some("s") {
        return None;
    }
    cell.descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "v")
        .and_then(|n| n.text())
        .and_then(|s| s.trim().parse::<u32>().ok())
}

#[test]
fn xlsx_document_writes_image_alt_text_as_shared_string() -> Result<(), Box<dyn std::error::Error>>
{
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    // Ensure the input workbook already contains a sharedStrings.xml table.
    worksheet.write_string(0, 0, "Hello")?;
    let bytes = workbook.save_to_buffer()?;

    let mut doc = load_from_bytes(&bytes)?;
    let sheet_id = doc.workbook.sheets[0].id;

    doc.set_cell_value(
        sheet_id,
        CellRef::from_a1("B1")?,
        CellValue::Image(ImageValue {
            image_id: ImageId::new("image1.png"),
            alt_text: Some("AltText".to_string()),
            width: None,
            height: None,
        }),
    );

    let saved = doc.save_to_vec()?;
    let sheet_xml = zip_part_to_string(&saved, "xl/worksheets/sheet1.xml");
    let shared_strings_xml = zip_part_to_string(&saved, "xl/sharedStrings.xml");

    let strings = shared_strings_as_vec(&shared_strings_xml);
    let alt_idx = strings
        .iter()
        .position(|s| s == "AltText")
        .expect("AltText should be present in sharedStrings.xml") as u32;

    let b1_idx = cell_shared_string_index(&sheet_xml, "B1").expect("B1 should be a shared string");
    assert_eq!(
        b1_idx, alt_idx,
        "B1 shared-string index should point at the AltText entry"
    );

    Ok(())
}

