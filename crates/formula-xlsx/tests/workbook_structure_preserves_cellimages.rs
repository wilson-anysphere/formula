use std::io::{Cursor, Read, Write};

use quick_xml::events::Event;
use quick_xml::Reader;
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

fn build_cellimages_fixture() -> Vec<u8> {
    let content_types = r#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/cellimages.xml" ContentType="application/vnd.ms-excel.cellimages+xml"/>
</Types>"#;

    let root_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="xl/workbook.xml"/>
</Relationships>"#;

    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    Target="styles.xml"/>
  <Relationship Id="rId3"
    Type="http://schemas.microsoft.com/office/2019/relationships/cellImages"
    Target="cellimages.xml"/>
</Relationships>"#;

    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"#;

    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>"#;

    let cellimages_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cellImages xmlns="http://schemas.microsoft.com/office/spreadsheetml/2019/cellimages">
  <cellImage/>
</cellImages>"#;

    let cellimages_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    Target="media/image1.png"/>
</Relationships>"#;

    let image_bytes: &[u8] = b"\x89PNG\r\n\x1a\nfake-png-payload";

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);

    fn add_file(
        zip: &mut ZipWriter<Cursor<Vec<u8>>>,
        options: FileOptions<()>,
        name: &str,
        bytes: &[u8],
    ) {
        zip.start_file(name, options).unwrap();
        zip.write_all(bytes).unwrap();
    }

    add_file(&mut zip, options, "[Content_Types].xml", content_types.as_bytes());
    add_file(&mut zip, options, "_rels/.rels", root_rels.as_bytes());
    add_file(&mut zip, options, "xl/workbook.xml", workbook_xml.as_bytes());
    add_file(
        &mut zip,
        options,
        "xl/_rels/workbook.xml.rels",
        workbook_rels.as_bytes(),
    );
    add_file(&mut zip, options, "xl/styles.xml", styles_xml.as_bytes());
    add_file(
        &mut zip,
        options,
        "xl/worksheets/sheet1.xml",
        worksheet_xml.as_bytes(),
    );

    add_file(&mut zip, options, "xl/cellimages.xml", cellimages_xml.as_bytes());
    add_file(
        &mut zip,
        options,
        "xl/_rels/cellimages.xml.rels",
        cellimages_rels.as_bytes(),
    );
    add_file(&mut zip, options, "xl/media/image1.png", image_bytes);

    zip.finish().unwrap().into_inner()
}

fn zip_part(zip_bytes: &[u8], name: &str) -> Vec<u8> {
    let cursor = Cursor::new(zip_bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut file = archive.by_name(name).expect("part exists");
    let mut buf = Vec::new();
    file.read_to_end(&mut buf).expect("read part");
    buf
}

fn workbook_rels_has_cellimages(xml: &[u8]) -> bool {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) if e.name().as_ref() == b"Relationship" => {
                let mut target = None;
                for attr in e.attributes().flatten() {
                    if attr.key.as_ref() == b"Target" {
                        target = Some(attr.unescape_value().expect("attr").into_owned());
                    }
                }
                if target
                    .as_deref()
                    .is_some_and(|t| t.trim_end_matches('/').ends_with("cellimages.xml"))
                {
                    return true;
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(_) => break,
        }
        buf.clear();
    }
    false
}

fn assert_cellimages_parts_preserved(before: &[u8], after: &[u8]) {
    for part in [
        "xl/cellimages.xml",
        "xl/_rels/cellimages.xml.rels",
        "xl/media/image1.png",
    ] {
        assert_eq!(
            zip_part(before, part),
            zip_part(after, part),
            "{part} must be preserved byte-for-byte",
        );
    }

    let workbook_rels = zip_part(after, "xl/_rels/workbook.xml.rels");
    assert!(
        workbook_rels_has_cellimages(&workbook_rels),
        "xl/_rels/workbook.xml.rels must retain the cellimages.xml relationship",
    );

    let content_types = String::from_utf8(zip_part(after, "[Content_Types].xml")).expect("utf8");
    assert!(
        content_types.contains(r#"PartName="/xl/cellimages.xml""#),
        r#"[Content_Types].xml must retain Override PartName="/xl/cellimages.xml""#,
    );
}

#[test]
fn add_sheet_preserves_cellimages_relationship_content_types_and_parts(
) -> Result<(), Box<dyn std::error::Error>> {
    let fixture = build_cellimages_fixture();

    let mut doc = formula_xlsx::load_from_bytes(&fixture)?;
    doc.workbook.add_sheet("Added")?;

    let saved = doc.save_to_vec()?;
    assert_cellimages_parts_preserved(&fixture, &saved);

    Ok(())
}

#[test]
fn delete_sheet_preserves_cellimages_relationship_content_types_and_parts(
) -> Result<(), Box<dyn std::error::Error>> {
    let fixture = build_cellimages_fixture();

    let mut doc = formula_xlsx::load_from_bytes(&fixture)?;

    // Ensure workbook still has at least one sheet after deleting the original.
    doc.workbook.add_sheet("Second")?;
    let original_sheet_id = doc.workbook.sheets[0].id;
    doc.workbook.delete_sheet(original_sheet_id)?;

    let saved = doc.save_to_vec()?;
    assert_cellimages_parts_preserved(&fixture, &saved);

    Ok(())
}

