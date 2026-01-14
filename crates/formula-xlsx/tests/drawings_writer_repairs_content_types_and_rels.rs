use std::collections::BTreeMap;
use std::io::{Cursor, Read, Write};

use formula_xlsx::drawings::DrawingPart;
use formula_xlsx::load_from_bytes;
use roxmltree::Document;
use zip::write::FileOptions;
use zip::{ZipArchive, ZipWriter};

const DRAWING_CONTENT_TYPE: &str = "application/vnd.openxmlformats-officedocument.drawing+xml";
const CHART_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";

fn build_corrupted_image_fixture() -> Vec<u8> {
    let fixture_bytes = include_bytes!("../../../fixtures/xlsx/basic/image.xlsx");
    let cursor = Cursor::new(fixture_bytes.as_slice());
    let mut archive = ZipArchive::new(cursor).expect("open fixture zip");

    let mut parts: BTreeMap<String, Vec<u8>> = BTreeMap::new();
    for i in 0..archive.len() {
        let mut file = archive.by_index(i).expect("zip file");
        if file.is_dir() {
            continue;
        }
        let name = file.name().to_string();
        let mut buf = Vec::new();
        file.read_to_end(&mut buf).expect("read zip file");
        parts.insert(name, buf);
    }

    // Remove the drawing override and png Default from [Content_Types].xml.
    let ct_name = "[Content_Types].xml";
    let ct = String::from_utf8(parts.get(ct_name).expect("ct part").clone()).expect("ct utf8");
    let ct = ct.replace(r#"<Default Extension="png" ContentType="image/png"/>"#, "");
    let ct = ct.replace(
        r#"<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>"#,
        "",
    );
    parts.insert(ct_name.to_string(), ct.into_bytes());

    // Drop the worksheet .rels so the writer must recreate the drawing relationship.
    parts.remove("xl/worksheets/_rels/sheet1.xml.rels");

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    for (name, bytes) in parts {
        zip.start_file(name, options).expect("start file");
        zip.write_all(&bytes).expect("write file");
    }

    zip.finish().expect("finish zip").into_inner()
}

fn build_corrupted_chart_fixture() -> Vec<u8> {
    let fixture_bytes = include_bytes!("../../../fixtures/xlsx/charts/basic-chart.xlsx");
    let cursor = Cursor::new(fixture_bytes.as_slice());
    let mut archive = ZipArchive::new(cursor).expect("open fixture zip");

    let mut parts: BTreeMap<String, Vec<u8>> = BTreeMap::new();
    for i in 0..archive.len() {
        let mut file = archive.by_index(i).expect("zip file");
        if file.is_dir() {
            continue;
        }
        let name = file.name().to_string();
        let mut buf = Vec::new();
        file.read_to_end(&mut buf).expect("read zip file");
        parts.insert(name, buf);
    }

    // Remove the chart override from [Content_Types].xml.
    let ct_name = "[Content_Types].xml";
    let ct = String::from_utf8(parts.get(ct_name).expect("ct part").clone()).expect("ct utf8");
    let ct = ct.replace(
        r#"<Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>"#,
        "",
    );
    parts.insert(ct_name.to_string(), ct.into_bytes());

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    for (name, bytes) in parts {
        zip.start_file(name, options).expect("start file");
        zip.write_all(&bytes).expect("write file");
    }

    zip.finish().expect("finish zip").into_inner()
}

#[test]
fn save_repairs_drawing_content_types_and_sheet_rels() {
    let corrupted = build_corrupted_image_fixture();
    let mut doc = load_from_bytes(&corrupted).expect("load corrupted xlsx");

    // Populate `Worksheet.drawings` so the writer path treats this worksheet as having drawings.
    let sheet_id = doc.workbook.sheets[0].id;
    let parts = doc.parts().clone();
    let part =
        DrawingPart::parse_from_parts(0, "xl/drawings/drawing1.xml", &parts, &mut doc.workbook)
            .expect("parse drawing");
    doc.workbook
        .sheet_mut(sheet_id)
        .expect("sheet exists")
        .drawings = part.objects;

    let saved = doc.save_to_vec().expect("save repaired xlsx");

    let cursor = Cursor::new(saved);
    let mut archive = ZipArchive::new(cursor).expect("open saved zip");

    // [Content_Types].xml: drawing override and png Default must exist.
    let mut ct_xml = String::new();
    archive
        .by_name("[Content_Types].xml")
        .expect("ct part exists")
        .read_to_string(&mut ct_xml)
        .expect("read ct xml");
    let ct_doc = Document::parse(&ct_xml).expect("parse ct xml");

    assert!(
        ct_doc.descendants().any(|n| {
            n.is_element()
                && n.tag_name().name() == "Override"
                && n.attribute("PartName") == Some("/xl/drawings/drawing1.xml")
                && n.attribute("ContentType") == Some(DRAWING_CONTENT_TYPE)
        }),
        "expected drawing Override content type to be present, got:\n{ct_xml}"
    );

    assert!(
        ct_doc.descendants().any(|n| {
            n.is_element()
                && n.tag_name().name() == "Default"
                && n.attribute("Extension") == Some("png")
                && n.attribute("ContentType") == Some("image/png")
        }),
        "expected png Default content type to be present, got:\n{ct_xml}"
    );

    // Worksheet `.rels`: drawing relationship must exist.
    let mut rels_xml = String::new();
    archive
        .by_name("xl/worksheets/_rels/sheet1.xml.rels")
        .expect("sheet rels exists")
        .read_to_string(&mut rels_xml)
        .expect("read sheet rels");
    let rels_doc = Document::parse(&rels_xml).expect("parse sheet rels");

    assert!(
        rels_doc.descendants().any(|n| {
            n.is_element()
                && n.tag_name().name() == "Relationship"
                && n.attribute("Type")
                    == Some("http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing")
                && n.attribute("Target") == Some("../drawings/drawing1.xml")
        }),
        "expected sheet drawing relationship to exist, got:\n{rels_xml}"
    );
}

#[test]
fn save_repairs_chart_content_types_when_drawing_rels_reference_chart() {
    let corrupted = build_corrupted_chart_fixture();
    let mut doc = load_from_bytes(&corrupted).expect("load corrupted xlsx");

    // Populate `Worksheet.drawings` so the writer path treats this worksheet as having drawings.
    let sheet_id = doc.workbook.sheets[0].id;
    let parts = doc.parts().clone();
    let part =
        DrawingPart::parse_from_parts(0, "xl/drawings/drawing1.xml", &parts, &mut doc.workbook)
            .expect("parse drawing");
    doc.workbook
        .sheet_mut(sheet_id)
        .expect("sheet exists")
        .drawings = part.objects;

    let saved = doc.save_to_vec().expect("save repaired xlsx");

    let cursor = Cursor::new(saved);
    let mut archive = ZipArchive::new(cursor).expect("open saved zip");

    let mut ct_xml = String::new();
    archive
        .by_name("[Content_Types].xml")
        .expect("ct part exists")
        .read_to_string(&mut ct_xml)
        .expect("read ct xml");
    let ct_doc = Document::parse(&ct_xml).expect("parse ct xml");

    assert!(
        ct_doc.descendants().any(|n| {
            n.is_element()
                && n.tag_name().name() == "Override"
                && n.attribute("PartName") == Some("/xl/charts/chart1.xml")
                && n.attribute("ContentType") == Some(CHART_CONTENT_TYPE)
        }),
        "expected chart Override content type to be present, got:\n{ct_xml}"
    );
}
