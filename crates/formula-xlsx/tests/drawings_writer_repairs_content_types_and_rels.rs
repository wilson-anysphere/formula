use std::collections::{BTreeMap, HashMap};
use std::io::{Cursor, Read, Write};

use formula_model::drawings::{
    Anchor, AnchorPoint, CellOffset, DrawingObject, DrawingObjectId, DrawingObjectKind, EmuSize,
    ImageData,
};
use formula_model::CellRef;
use formula_xlsx::drawings::DrawingPart;
use formula_xlsx::load_from_bytes;
use formula_xlsx::XlsxDocument;
use roxmltree::Document;
use zip::write::FileOptions;
use zip::{ZipArchive, ZipWriter};

const DRAWING_CONTENT_TYPE: &str = "application/vnd.openxmlformats-officedocument.drawing+xml";
const CHART_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
const DRAWING_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";

fn zip_part(bytes: &[u8], name: &str) -> Vec<u8> {
    let cursor = Cursor::new(bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut file = archive.by_name(name).expect("part exists");
    let mut buf = Vec::new();
    file.read_to_end(&mut buf).expect("read part");
    buf
}

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

fn build_fixture_missing_drawing_content_types() -> Vec<u8> {
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

    // Remove the drawing override and png Default from [Content_Types].xml, but keep the worksheet
    // relationships + drawing parts intact. This exercises the "repair content types without
    // rewriting drawing parts" path.
    let ct_name = "[Content_Types].xml";
    let ct = String::from_utf8(parts.get(ct_name).expect("ct part").clone()).expect("ct utf8");
    let ct = ct.replace(r#"<Default Extension="png" ContentType="image/png"/>"#, "");
    let ct = ct.replace(
        r#"<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>"#,
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
fn save_repairs_content_types_without_rewriting_drawing_parts_when_unchanged() {
    let corrupted = build_fixture_missing_drawing_content_types();

    // Capture drawing part bytes before saving; the writer should only patch [Content_Types].xml
    // and keep the drawing XML + `.rels` parts byte-for-byte stable.
    let original_drawing_xml = zip_part(&corrupted, "xl/drawings/drawing1.xml");
    let original_drawing_rels_xml = zip_part(&corrupted, "xl/drawings/_rels/drawing1.xml.rels");

    let doc = load_from_bytes(&corrupted).expect("load corrupted xlsx");
    let saved = doc.save_to_vec().expect("save repaired xlsx");

    // [Content_Types].xml: drawing override and png Default must exist.
    let cursor = Cursor::new(saved.clone());
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

    // Drawing parts should not be rewritten.
    assert_eq!(
        zip_part(&saved, "xl/drawings/drawing1.xml"),
        original_drawing_xml,
        "expected drawing XML to remain byte-for-byte stable"
    );
    assert_eq!(
        zip_part(&saved, "xl/drawings/_rels/drawing1.xml.rels"),
        original_drawing_rels_xml,
        "expected drawing rels XML to remain byte-for-byte stable"
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

#[test]
fn save_adds_content_types_default_for_jpg_media() {
    let mut workbook = formula_model::Workbook::default();
    let sheet_id = workbook.add_sheet("Sheet1").expect("add sheet");

    let image_id = workbook.images.ensure_unique_name("image", "jpg");
    workbook.images.insert(
        image_id.clone(),
        ImageData {
            // Minimal JPEG header/footer (not intended to render, but sufficient for writer paths).
            bytes: vec![0xFF, 0xD8, 0xFF, 0xD9],
            content_type: Some("image/jpeg".to_string()),
        },
    );

    let ext = EmuSize::new(914_400, 914_400);
    let drawing = DrawingObject {
        id: DrawingObjectId(1),
        kind: DrawingObjectKind::Image { image_id },
        anchor: Anchor::OneCell {
            from: AnchorPoint::new(CellRef::new(0, 0), CellOffset::new(0, 0)),
            ext,
        },
        z_order: 0,
        size: Some(ext),
        preserved: HashMap::new(),
    };

    workbook
        .sheet_mut(sheet_id)
        .expect("sheet exists")
        .drawings = vec![drawing];

    let doc = XlsxDocument::new(workbook);
    let saved = doc.save_to_vec().expect("save xlsx");

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
                && n.tag_name().name() == "Default"
                && n.attribute("Extension") == Some("jpg")
                && n.attribute("ContentType") == Some("image/jpeg")
        }),
        "expected jpg Default content type to be present, got:\n{ct_xml}"
    );
}

#[test]
fn save_adds_content_types_default_for_jpeg_media() {
    let mut workbook = formula_model::Workbook::default();
    let sheet_id = workbook.add_sheet("Sheet1").expect("add sheet");

    let image_id = workbook.images.ensure_unique_name("image", "jpeg");
    workbook.images.insert(
        image_id.clone(),
        ImageData {
            // Minimal JPEG header/footer (not intended to render, but sufficient for writer paths).
            bytes: vec![0xFF, 0xD8, 0xFF, 0xD9],
            content_type: Some("image/jpeg".to_string()),
        },
    );

    let ext = EmuSize::new(914_400, 914_400);
    let drawing = DrawingObject {
        id: DrawingObjectId(1),
        kind: DrawingObjectKind::Image { image_id },
        anchor: Anchor::OneCell {
            from: AnchorPoint::new(CellRef::new(0, 0), CellOffset::new(0, 0)),
            ext,
        },
        z_order: 0,
        size: Some(ext),
        preserved: HashMap::new(),
    };

    workbook
        .sheet_mut(sheet_id)
        .expect("sheet exists")
        .drawings = vec![drawing];

    let doc = XlsxDocument::new(workbook);
    let saved = doc.save_to_vec().expect("save xlsx");

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
                && n.tag_name().name() == "Default"
                && n.attribute("Extension") == Some("jpeg")
                && n.attribute("ContentType") == Some("image/jpeg")
        }),
        "expected jpeg Default content type to be present, got:\n{ct_xml}"
    );
}

#[test]
fn save_adds_content_types_default_for_emf_media() {
    let mut workbook = formula_model::Workbook::default();
    let sheet_id = workbook.add_sheet("Sheet1").expect("add sheet");

    let image_id = workbook.images.ensure_unique_name("image", "emf");
    workbook.images.insert(
        image_id.clone(),
        ImageData {
            // Placeholder bytes: writer does not validate image payloads.
            bytes: vec![0, 1, 2, 3],
            content_type: Some("image/x-emf".to_string()),
        },
    );

    let ext = EmuSize::new(914_400, 914_400);
    let drawing = DrawingObject {
        id: DrawingObjectId(1),
        kind: DrawingObjectKind::Image { image_id },
        anchor: Anchor::OneCell {
            from: AnchorPoint::new(CellRef::new(0, 0), CellOffset::new(0, 0)),
            ext,
        },
        z_order: 0,
        size: Some(ext),
        preserved: HashMap::new(),
    };

    workbook
        .sheet_mut(sheet_id)
        .expect("sheet exists")
        .drawings = vec![drawing];

    let doc = XlsxDocument::new(workbook);
    let saved = doc.save_to_vec().expect("save xlsx");

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
                && n.tag_name().name() == "Default"
                && n.attribute("Extension") == Some("emf")
                && n.attribute("ContentType") == Some("image/x-emf")
        }),
        "expected emf Default content type to be present, got:\n{ct_xml}"
    );
}

#[test]
fn save_adds_content_types_default_for_wmf_media() {
    let mut workbook = formula_model::Workbook::default();
    let sheet_id = workbook.add_sheet("Sheet1").expect("add sheet");

    let image_id = workbook.images.ensure_unique_name("image", "wmf");
    workbook.images.insert(
        image_id.clone(),
        ImageData {
            // Placeholder bytes: writer does not validate image payloads.
            bytes: vec![0, 1, 2, 3],
            content_type: Some("image/x-wmf".to_string()),
        },
    );

    let ext = EmuSize::new(914_400, 914_400);
    let drawing = DrawingObject {
        id: DrawingObjectId(1),
        kind: DrawingObjectKind::Image { image_id },
        anchor: Anchor::OneCell {
            from: AnchorPoint::new(CellRef::new(0, 0), CellOffset::new(0, 0)),
            ext,
        },
        z_order: 0,
        size: Some(ext),
        preserved: HashMap::new(),
    };

    workbook
        .sheet_mut(sheet_id)
        .expect("sheet exists")
        .drawings = vec![drawing];

    let doc = XlsxDocument::new(workbook);
    let saved = doc.save_to_vec().expect("save xlsx");

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
                && n.tag_name().name() == "Default"
                && n.attribute("Extension") == Some("wmf")
                && n.attribute("ContentType") == Some("image/x-wmf")
        }),
        "expected wmf Default content type to be present, got:\n{ct_xml}"
    );
}

#[test]
fn save_adds_content_types_default_for_webp_media() {
    let mut workbook = formula_model::Workbook::default();
    let sheet_id = workbook.add_sheet("Sheet1").expect("add sheet");

    let image_id = workbook.images.ensure_unique_name("image", "webp");
    workbook.images.insert(
        image_id.clone(),
        ImageData {
            // Placeholder bytes: writer does not validate image payloads.
            bytes: vec![0, 1, 2, 3],
            content_type: Some("image/webp".to_string()),
        },
    );

    let ext = EmuSize::new(914_400, 914_400);
    let drawing = DrawingObject {
        id: DrawingObjectId(1),
        kind: DrawingObjectKind::Image { image_id },
        anchor: Anchor::OneCell {
            from: AnchorPoint::new(CellRef::new(0, 0), CellOffset::new(0, 0)),
            ext,
        },
        z_order: 0,
        size: Some(ext),
        preserved: HashMap::new(),
    };

    workbook
        .sheet_mut(sheet_id)
        .expect("sheet exists")
        .drawings = vec![drawing];

    let doc = XlsxDocument::new(workbook);
    let saved = doc.save_to_vec().expect("save xlsx");

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
                && n.tag_name().name() == "Default"
                && n.attribute("Extension") == Some("webp")
                && n.attribute("ContentType") == Some("image/webp")
        }),
        "expected webp Default content type to be present, got:\n{ct_xml}"
    );
}

#[test]
fn save_adds_content_types_default_for_svg_media() {
    let mut workbook = formula_model::Workbook::default();
    let sheet_id = workbook.add_sheet("Sheet1").expect("add sheet");

    let image_id = workbook.images.ensure_unique_name("image", "svg");
    workbook.images.insert(
        image_id.clone(),
        ImageData {
            // Minimal SVG payload (not intended to render, but sufficient for writer paths).
            bytes: br#"<svg xmlns="http://www.w3.org/2000/svg"></svg>"#.to_vec(),
            content_type: Some("image/svg+xml".to_string()),
        },
    );

    let ext = EmuSize::new(914_400, 914_400);
    let drawing = DrawingObject {
        id: DrawingObjectId(1),
        kind: DrawingObjectKind::Image { image_id },
        anchor: Anchor::OneCell {
            from: AnchorPoint::new(CellRef::new(0, 0), CellOffset::new(0, 0)),
            ext,
        },
        z_order: 0,
        size: Some(ext),
        preserved: HashMap::new(),
    };

    workbook
        .sheet_mut(sheet_id)
        .expect("sheet exists")
        .drawings = vec![drawing];

    let doc = XlsxDocument::new(workbook);
    let saved = doc.save_to_vec().expect("save xlsx");

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
                && n.tag_name().name() == "Default"
                && n.attribute("Extension") == Some("svg")
                && n.attribute("ContentType") == Some("image/svg+xml")
        }),
        "expected svg Default content type to be present, got:\n{ct_xml}"
    );
}
