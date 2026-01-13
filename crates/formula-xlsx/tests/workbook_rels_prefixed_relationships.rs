use std::io::{Cursor, Read, Write};

use quick_xml::events::Event;
use quick_xml::Reader;
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

const REL_TYPE_STYLES: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";

fn build_prefixed_workbook_rels_xlsx() -> Vec<u8> {
    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/customSheet.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/customStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>"#;

    let root_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#;

    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    // Important: the worksheet relationship target is NOT the default `sheet1.xml`, so
    // relationship parsing must succeed for the loader to find the worksheet part.
    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rel:Relationships xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships">
  <rel:Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/customSheet.xml"/>
  <rel:Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="customStyles.xml"/>
</rel:Relationships>"#;

    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>"#;

    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);

    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(content_types.as_bytes()).unwrap();

    zip.start_file("_rels/.rels", options).unwrap();
    zip.write_all(root_rels.as_bytes()).unwrap();

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options).unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/customSheet.xml", options)
        .unwrap();
    zip.write_all(worksheet_xml.as_bytes()).unwrap();

    zip.start_file("xl/customStyles.xml", options).unwrap();
    zip.write_all(styles_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

fn build_prefixed_workbook_rels_xlsx_missing_styles_relationship() -> Vec<u8> {
    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/customSheet.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>"#;

    let root_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#;

    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rel:Relationships xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships">
  <rel:Relationship Id = "rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/customSheet.xml"/>
</rel:Relationships>"#;

    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>"#;

    // Include a minimal styles.xml part, but omit the workbook.xml.rels relationship to it.
    // The writer should preserve the existing `rel:` prefix when inserting the missing
    // relationship.
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);

    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(content_types.as_bytes()).unwrap();

    zip.start_file("_rels/.rels", options).unwrap();
    zip.write_all(root_rels.as_bytes()).unwrap();

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options).unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/customSheet.xml", options)
        .unwrap();
    zip.write_all(worksheet_xml.as_bytes()).unwrap();

    zip.start_file("xl/styles.xml", options).unwrap();
    zip.write_all(styles_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

fn zip_part(bytes: &[u8], name: &str) -> Vec<u8> {
    let cursor = Cursor::new(bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut file = archive.by_name(name).expect("part exists");
    let mut buf = Vec::new();
    file.read_to_end(&mut buf).expect("read part");
    buf
}

fn sheet_rid_by_name(xml: &[u8], sheet_name: &str) -> Option<String> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf).ok()? {
            Event::Start(e) | Event::Empty(e) if e.name().as_ref() == b"sheet" => {
                let mut name = None;
                let mut rid = None;
                for attr in e.attributes().flatten() {
                    let value = attr.unescape_value().ok()?.into_owned();
                    match attr.key.as_ref() {
                        b"name" => name = Some(value),
                        b"r:id" => rid = Some(value),
                        _ => {}
                    }
                }
                if name.as_deref() == Some(sheet_name) {
                    return rid;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    None
}

#[test]
fn workbook_rels_prefixed_relationship_elements_load_and_save() -> Result<(), Box<dyn std::error::Error>>
{
    let bytes = build_prefixed_workbook_rels_xlsx();

    // The lightweight workbook reader should resolve the worksheet part using workbook.xml.rels.
    let workbook = formula_xlsx::read_workbook_model_from_bytes(&bytes)?;
    assert!(workbook.sheet_by_name("Sheet1").is_some());

    // Full-fidelity loader should also resolve the worksheet part.
    let mut doc = formula_xlsx::load_from_bytes(&bytes)?;
    assert!(doc.workbook.sheet_by_name("Sheet1").is_some());

    // Saving should not panic, and should respect prefixed Relationship elements when resolving
    // `styles.xml` targets (no stray `xl/styles.xml` should be synthesized).
    let saved = doc.save_to_vec()?;
    let mut saved_zip = ZipArchive::new(Cursor::new(&saved))?;
    assert!(
        saved_zip.by_name("xl/customStyles.xml").is_ok(),
        "expected writer to keep styles part at custom target"
    );
    assert!(
        saved_zip.by_name("xl/styles.xml").is_err(),
        "writer should not synthesize xl/styles.xml when workbook.xml.rels points elsewhere"
    );

    // Exercise sheet-add rel patching: it must find the end of <rel:Relationships> and insert
    // a prefixed <rel:Relationship .../> entry.
    doc.workbook.add_sheet("Added")?;
    let saved_with_sheet = doc.save_to_vec()?;

    let workbook_xml = zip_part(&saved_with_sheet, "xl/workbook.xml");
    let added_rid =
        sheet_rid_by_name(&workbook_xml, "Added").expect("added sheet must exist in workbook.xml");

    let rels_xml = String::from_utf8(zip_part(&saved_with_sheet, "xl/_rels/workbook.xml.rels"))?;
    assert!(
        rels_xml.contains(&format!(r#"<rel:Relationship Id="{added_rid}""#)),
        "expected workbook.xml.rels to contain a prefixed relationship element for the added sheet"
    );

    Ok(())
}

#[test]
fn writer_inserts_missing_styles_relationship_using_existing_prefix(
) -> Result<(), Box<dyn std::error::Error>> {
    let bytes = build_prefixed_workbook_rels_xlsx_missing_styles_relationship();

    let doc = formula_xlsx::load_from_bytes(&bytes)?;
    let saved = doc.save_to_vec()?;

    let rels_xml = String::from_utf8(zip_part(&saved, "xl/_rels/workbook.xml.rels"))?;
    assert!(
        rels_xml.contains(&format!(
            r#"<rel:Relationship Id="rId2" Type="{REL_TYPE_STYLES}" Target="styles.xml"/>"#
        )),
        "expected writer to insert a prefixed styles relationship into workbook.xml.rels, got:\n{rels_xml}"
    );

    let mut archive = ZipArchive::new(Cursor::new(&saved))?;
    assert!(
        archive.by_name("xl/styles.xml").is_ok(),
        "expected writer to synthesize xl/styles.xml when styles relationship was missing"
    );

    Ok(())
}
