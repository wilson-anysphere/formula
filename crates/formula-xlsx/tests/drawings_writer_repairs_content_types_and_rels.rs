use std::collections::{BTreeMap, HashMap};
use std::io::{Cursor, Read, Write};

use formula_model::drawings::{
    Anchor, AnchorPoint, CellOffset, DrawingObject, DrawingObjectId, DrawingObjectKind, EmuSize,
    ImageData,
};
use formula_model::CellRef;
use formula_xlsx::drawings::DrawingPart;
use formula_xlsx::load_from_bytes;
use roxmltree::Document;
use zip::write::FileOptions;
use zip::{ZipArchive, ZipWriter};

const DRAWING_CONTENT_TYPE: &str = "application/vnd.openxmlformats-officedocument.drawing+xml";
const CHART_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
const DRAWING_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";

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

#[test]
fn save_preserves_existing_sheet_drawing_relationship_id() {
    let fixture_bytes = include_bytes!("../../../fixtures/xlsx/basic/image.xlsx");
    let mut doc = load_from_bytes(fixture_bytes).expect("load xlsx");

    // Capture the original sheet drawing relationship id/target.
    let sheet_rels_xml = std::str::from_utf8(
        doc.parts()
            .get("xl/worksheets/_rels/sheet1.xml.rels")
            .expect("sheet rels part exists"),
    )
    .expect("sheet rels utf8");
    let rels_doc = Document::parse(sheet_rels_xml).expect("parse sheet rels");
    let original_rel = rels_doc
        .descendants()
        .find(|n| {
            n.is_element()
                && n.tag_name().name() == "Relationship"
                && n.attribute("Type") == Some(DRAWING_REL_TYPE)
        })
        .expect("expected drawing relationship in fixture");
    let original_rid = original_rel
        .attribute("Id")
        .expect("drawing relationship has Id")
        .to_string();
    let original_target = original_rel
        .attribute("Target")
        .expect("drawing relationship has Target")
        .to_string();

    // Populate `Worksheet.drawings` from the drawing part.
    let sheet_id = doc.workbook.sheets[0].id;
    let parts = doc.parts().clone();
    let part =
        DrawingPart::parse_from_parts(0, "xl/drawings/drawing1.xml", &parts, &mut doc.workbook)
            .expect("parse drawing");
    let mut drawings = part.objects;

    // Force the writer down the `drawings_need_emit` path by introducing a new image object
    // whose media part does not exist in the original package.
    let existing_image_id = drawings
        .iter()
        .find_map(|o| match &o.kind {
            DrawingObjectKind::Image { image_id } => Some(image_id.clone()),
            _ => None,
        })
        .expect("fixture should contain an image drawing");
    let existing_bytes = doc
        .workbook
        .images
        .get(&existing_image_id)
        .expect("fixture image bytes exist")
        .bytes
        .clone();

    let new_image_id = doc.workbook.images.ensure_unique_name("added", "png");
    doc.workbook.images.insert(
        new_image_id.clone(),
        ImageData {
            bytes: existing_bytes,
            content_type: Some("image/png".to_string()),
        },
    );

    let next_object_id = drawings
        .iter()
        .map(|o| o.id.0)
        .max()
        .unwrap_or(0)
        .saturating_add(1);
    let next_z = drawings.iter().map(|o| o.z_order).max().unwrap_or(0) + 1;
    let ext = EmuSize::new(914_400, 914_400);
    drawings.push(DrawingObject {
        id: DrawingObjectId(next_object_id),
        kind: DrawingObjectKind::Image {
            image_id: new_image_id,
        },
        anchor: Anchor::OneCell {
            from: AnchorPoint::new(CellRef::new(0, 0), CellOffset::new(0, 0)),
            ext,
        },
        z_order: next_z,
        size: Some(ext),
        preserved: HashMap::new(),
    });

    doc.workbook
        .sheet_mut(sheet_id)
        .expect("sheet exists")
        .drawings = drawings;

    let saved = doc.save_to_vec().expect("save");

    let cursor = Cursor::new(saved);
    let mut archive = ZipArchive::new(cursor).expect("open saved zip");
    let mut rels_xml = String::new();
    archive
        .by_name("xl/worksheets/_rels/sheet1.xml.rels")
        .expect("sheet rels exists")
        .read_to_string(&mut rels_xml)
        .expect("read sheet rels");
    let rels_doc = Document::parse(&rels_xml).expect("parse sheet rels");
    let rel = rels_doc
        .descendants()
        .find(|n| {
            n.is_element()
                && n.tag_name().name() == "Relationship"
                && n.attribute("Type") == Some(DRAWING_REL_TYPE)
        })
        .expect("drawing relationship present after save");
    assert_eq!(
        rel.attribute("Id"),
        Some(original_rid.as_str()),
        "expected writer to preserve existing drawing relationship Id"
    );
    assert_eq!(
        rel.attribute("Target"),
        Some(original_target.as_str()),
        "expected writer to preserve existing drawing relationship Target"
    );
}
