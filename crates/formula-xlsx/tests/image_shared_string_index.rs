use std::io::{Cursor, Read as _, Write as _};

use formula_model::drawings::ImageId;
use formula_model::{CellRef, CellValue, ImageValue};
use formula_xlsx::load_from_bytes;
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

fn build_duplicate_shared_strings_xlsx() -> Vec<u8> {
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>"#;

    let root_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#;

    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>"#;

    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>0</v></c></row>
    <row r="2"><c r="A2" t="s"><v>1</v></c></row>
  </sheetData>
</worksheet>"#;

    // Two `si` entries with identical visible text but different phonetic markers.
    let shared_strings_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
  <si><t>Base</t><rPh sb="0" eb="4"><t>PHONETIC_A</t></rPh></si>
  <si><t>Base</t><rPh sb="0" eb="4"><t>PHONETIC_B</t></rPh></si>
</sst>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);

    zip.start_file("_rels/.rels", options).unwrap();
    zip.write_all(root_rels.as_bytes()).unwrap();

    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(content_types.as_bytes()).unwrap();

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(worksheet_xml.as_bytes()).unwrap();

    zip.start_file("xl/sharedStrings.xml", options).unwrap();
    zip.write_all(shared_strings_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

fn zip_part(zip_bytes: &[u8], name: &str) -> String {
    let cursor = Cursor::new(zip_bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut file = archive.by_name(name).expect("part exists");
    let mut buf = String::new();
    file.read_to_string(&mut buf).expect("read part");
    buf
}

#[test]
fn document_writer_preserves_shared_string_index_for_image_alt_text() -> Result<(), Box<dyn std::error::Error>>
{
    let bytes = build_duplicate_shared_strings_xlsx();
    let mut doc = load_from_bytes(&bytes)?;
    let sheet_id = doc.workbook.sheets[0].id;

    doc.set_cell_value(
        sheet_id,
        CellRef::from_a1("A2")?,
        CellValue::Image(ImageValue {
            image_id: ImageId::new("image1.png"),
            alt_text: Some("Base".to_string()),
            width: None,
            height: None,
        }),
    );

    let saved = doc.save_to_vec()?;
    let sheet_xml = zip_part(&saved, "xl/worksheets/sheet1.xml");
    let doc_xml = roxmltree::Document::parse(&sheet_xml)?;
    let ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    let cell = doc_xml
        .descendants()
        .find(|n| n.has_tag_name((ns, "c")) && n.attribute("r") == Some("A2"))
        .expect("A2 cell should exist");
    assert_eq!(cell.attribute("t"), Some("s"));
    let v = cell
        .children()
        .find(|n| n.has_tag_name((ns, "v")))
        .and_then(|n| n.text())
        .unwrap_or_default();
    assert_eq!(
        v, "1",
        "expected writer to preserve the original shared string index when Image altText is unchanged"
    );
    Ok(())
}

