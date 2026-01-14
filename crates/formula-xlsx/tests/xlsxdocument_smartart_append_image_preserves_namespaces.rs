use std::collections::BTreeMap;
use std::io::{Cursor, Read, Write};

use base64::Engine;
use formula_model::drawings::{
    Anchor, AnchorPoint, CellOffset, DrawingObject, DrawingObjectId, DrawingObjectKind, EmuSize,
    ImageData,
};
use formula_model::CellRef;
use formula_xlsx::openxml;
use formula_xlsx::{load_from_bytes, XlsxPackage};
use zip::write::FileOptions;
use zip::{ZipArchive, ZipWriter};

const FIXTURE: &[u8] = include_bytes!("../../../fixtures/xlsx/basic/smartart.xlsx");

const DRAWING_XML: &str = "xl/drawings/drawing1.xml";
const DRAWING_RELS: &str = "xl/drawings/_rels/drawing1.xml.rels";
const REL_TYPE_IMAGE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

fn build_fixture_with_target_mode() -> Vec<u8> {
    let cursor = Cursor::new(FIXTURE);
    let mut archive = ZipArchive::new(cursor).expect("open fixture zip");

    let mut parts: BTreeMap<String, Vec<u8>> = BTreeMap::new();
    for i in 0..archive.len() {
        let mut file = archive.by_index(i).expect("zip entry");
        if file.is_dir() {
            continue;
        }
        let name = file.name().to_string();
        let mut buf = Vec::new();
        file.read_to_end(&mut buf).expect("read zip entry");
        parts.insert(name, buf);
    }

    let rels_xml = String::from_utf8(parts.get(DRAWING_RELS).expect("drawing rels").clone())
        .expect("rels xml utf8");
    let rels_xml = rels_xml.replace(
        r#"Target="../diagrams/data1.xml"/>"#,
        r#"Target="../diagrams/data1.xml" TargetMode="External"/>"#,
    );
    parts.insert(DRAWING_RELS.to_string(), rels_xml.into_bytes());

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
fn xlsxdocument_append_image_to_smartart_preserves_namespaces_and_rels() {
    let modified_fixture = build_fixture_with_target_mode();
    let mut doc = load_from_bytes(&modified_fixture).expect("load modified smartart fixture");

    // 1x1 transparent PNG.
    let png_bytes = base64::engine::general_purpose::STANDARD
        .decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/58HAQUBAO3+2NoAAAAASUVORK5CYII=")
        .expect("valid base64 png");

    let image_id = doc.workbook.images.ensure_unique_name("image", "png");
    doc.workbook.images.insert(
        image_id.clone(),
        ImageData {
            bytes: png_bytes.clone(),
            content_type: Some("image/png".to_string()),
        },
    );

    let sheet = doc
        .workbook
        .sheets
        .get_mut(0)
        .expect("fixture has at least one sheet");
    let next_id = sheet
        .drawings
        .iter()
        .map(|o| o.id.0)
        .max()
        .unwrap_or(0)
        .saturating_add(1);
    let ext = EmuSize::new(914_400, 914_400);
    let anchor = Anchor::OneCell {
        from: AnchorPoint::new(CellRef::new(0, 0), CellOffset::new(0, 0)),
        ext,
    };
    sheet.drawings.push(DrawingObject {
        id: DrawingObjectId(next_id),
        kind: DrawingObjectKind::Image {
            image_id: image_id.clone(),
        },
        anchor,
        z_order: sheet.drawings.len() as i32,
        size: Some(ext),
        preserved: Default::default(),
    });

    let saved = doc.save_to_vec().expect("save");

    let pkg = XlsxPackage::from_bytes(&saved).expect("open saved workbook");

    // The regenerated drawing XML must still be well-formed and include the SmartArt namespace
    // declarations needed for preserved `dgm:*` nodes.
    let drawing_xml = std::str::from_utf8(pkg.part(DRAWING_XML).expect("drawing part")).unwrap();
    roxmltree::Document::parse(drawing_xml).expect("drawing xml should be well-formed");
    assert!(
        drawing_xml.contains("xmlns:dgm="),
        "expected regenerated drawing to retain xmlns:dgm, got:\n{drawing_xml}"
    );

    // The regenerated drawing `.rels` should preserve the SmartArt diagram relationships (including
    // injected TargetMode) and add a new image relationship for the appended picture.
    let rels = openxml::parse_relationships(pkg.part(DRAWING_RELS).expect("drawing rels"))
        .expect("parse drawing rels");
    assert!(
        rels.iter().any(|rel| rel.type_uri.contains("diagramData")),
        "expected drawing rels to include SmartArt diagram relationships"
    );
    assert!(
        rels.iter().any(|rel| rel.target_mode.as_deref() == Some("External")),
        "expected injected TargetMode to be preserved"
    );

    let expected_image_target = format!("../media/{}", image_id.as_str());
    assert!(
        rels.iter()
            .any(|rel| rel.type_uri == REL_TYPE_IMAGE && rel.target == expected_image_target),
        "expected drawing rels to include an image relationship targeting {expected_image_target}, got: {rels:?}"
    );
}

