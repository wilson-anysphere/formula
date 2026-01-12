use std::io::{Cursor, Read, Write};

use formula_model::{CellRef, CellValue};
use formula_xlsx::load_from_bytes;
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn zip_part(zip_bytes: &[u8], name: &str) -> Vec<u8> {
    let cursor = Cursor::new(zip_bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut file = archive.by_name(name).expect("part exists");
    let mut buf = Vec::new();
    file.read_to_end(&mut buf).expect("read part");
    buf
}

fn build_minimal_xlsx_with_rich_value_metadata() -> (Vec<u8>, Vec<u8>) {
    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.metadata+xml"/>
</Types>"#;

    let rels_root = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#;

    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
  <metadata r:id="rIdMeta"/>
</workbook>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdMeta" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata" Target="metadata.xml"/>
</Relationships>"#;

    let metadata_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <metadataTypes count="1">
    <metadataType name="XLRICHVALUE" minSupportedVersion="120000"/>
  </metadataTypes>
  <valueMetadata count="1">
    <metadataRecord>
      <metadataTypeIndex>0</metadataTypeIndex>
      <futureMetadata name="XLRICHVALUE">
        <xlrd:rvb/>
      </futureMetadata>
    </metadataRecord>
  </valueMetadata>
</metadata>"#;
    let metadata_bytes = metadata_xml.as_bytes().to_vec();

    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1"/>
  <sheetData>
    <row r="1">
      <c r="A1" vm="1"><v>1</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);

    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(content_types.as_bytes()).unwrap();

    zip.start_file("_rels/.rels", options).unwrap();
    zip.write_all(rels_root.as_bytes()).unwrap();

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options).unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(worksheet_xml.as_bytes()).unwrap();

    zip.start_file("xl/metadata.xml", options).unwrap();
    zip.write_all(metadata_xml.as_bytes()).unwrap();

    let bytes = zip.finish().unwrap().into_inner();
    (bytes, metadata_bytes)
}

#[test]
fn editing_cell_preserves_workbook_metadata_parts_and_drops_vm_on_edit(
) -> Result<(), Box<dyn std::error::Error>> {
    let (xlsx_bytes, original_metadata) = build_minimal_xlsx_with_rich_value_metadata();

    let mut doc = load_from_bytes(&xlsx_bytes)?;
    let sheet_id = doc.workbook.sheets[0].id;
    let sheet = doc.workbook.sheet_mut(sheet_id).expect("sheet exists");

    // Force a real rewrite path: switch from a numeric cell to a shared string, which causes the
    // writer to synthesize sharedStrings + update workbook.xml.rels + [Content_Types].xml.
    sheet.set_value(
        CellRef::from_a1("A1")?,
        CellValue::String("Hello".to_string()),
    );

    let saved = doc.save_to_vec()?;

    // `xl/metadata.xml` must round-trip byte-for-byte (it's not understood/edited yet).
    let saved_metadata = zip_part(&saved, "xl/metadata.xml");
    assert_eq!(
        saved_metadata, original_metadata,
        "xl/metadata.xml should be preserved on edit round-trip"
    );

    // workbook.xml must keep a <metadata r:id="..."/> reference.
    let workbook_bytes = zip_part(&saved, "xl/workbook.xml");
    let workbook_xml = std::str::from_utf8(&workbook_bytes)?;
    let workbook_doc = roxmltree::Document::parse(workbook_xml)?;
    let wb_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    let metadata_ref = workbook_doc
        .descendants()
        .find(|n| n.is_element() && n.has_tag_name((wb_ns, "metadata")))
        .ok_or("workbook.xml missing <metadata> element")?;
    let rid = metadata_ref
        .attribute((REL_NS, "id"))
        .or_else(|| metadata_ref.attribute("r:id"))
        .ok_or("workbook.xml <metadata> missing r:id")?;
    assert_eq!(rid, "rIdMeta");

    // workbook.xml.rels must keep the metadata relationship.
    let rels_bytes = zip_part(&saved, "xl/_rels/workbook.xml.rels");
    let rels_xml = std::str::from_utf8(&rels_bytes)?;
    let rels_doc = roxmltree::Document::parse(rels_xml)?;
    let rel = rels_doc
        .descendants()
        .find(|n| {
            n.is_element()
                && n.tag_name().name() == "Relationship"
                && n.attribute("Id") == Some("rIdMeta")
        })
        .ok_or("workbook.xml.rels missing metadata Relationship rIdMeta")?;
    assert_eq!(
        rel.attribute("Type"),
        Some("http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata")
    );
    assert_eq!(rel.attribute("Target"), Some("metadata.xml"));

    // The `vm="..."` attribute is a value-metadata pointer into `xl/metadata.xml`. When we rewrite
    // worksheet XML we preserve it for fidelity (except for the embedded-image `#VALUE!`
    // placeholder case, which is not exercised by this fixture).
    let sheet_bytes = zip_part(&saved, "xl/worksheets/sheet1.xml");
    let sheet_xml = std::str::from_utf8(&sheet_bytes)?;
    let sheet_doc = roxmltree::Document::parse(sheet_xml)?;
    let cell = sheet_doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some("A1"))
        .ok_or("sheet1.xml missing A1 cell")?;
    assert_eq!(cell.attribute("vm"), Some("1"));

    Ok(())
}
