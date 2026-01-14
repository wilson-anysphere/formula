use std::io::{Cursor, Read as _};

use base64::Engine as _;
use formula_model::drawings::{
    Anchor, AnchorPoint, CellOffset, DrawingObject, DrawingObjectId, DrawingObjectKind, EmuSize,
    ImageData,
};
use formula_model::CellRef;
use formula_xlsx::load_from_bytes;
use roxmltree::Document;
use zip::ZipArchive;

const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL_TYPE_DRAWING: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
const REL_TYPE_IMAGE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

fn zip_part_to_bytes(bytes: &[u8], name: &str) -> Vec<u8> {
    let cursor = Cursor::new(bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut file = archive.by_name(name).expect("part exists");
    let mut out = Vec::new();
    file.read_to_end(&mut out).expect("read part");
    out
}

fn zip_part_to_string(bytes: &[u8], name: &str) -> String {
    String::from_utf8(zip_part_to_bytes(bytes, name)).expect("utf8")
}

#[test]
fn xlsx_document_writes_updated_worksheet_drawings_and_media(
) -> Result<(), Box<dyn std::error::Error>> {
    let fixture = include_bytes!("../../../fixtures/xlsx/basic/image.xlsx");
    let mut doc = load_from_bytes(fixture)?;

    let sheet_id = doc.workbook.sheets[0].id;
    assert!(
        !doc.workbook.sheets[0].drawings.is_empty(),
        "expected fixture to load at least one drawing"
    );

    // 2x2 transparent PNG.
    let png_bytes = base64::engine::general_purpose::STANDARD
        .decode("iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAQAAADZc7J/AAAADElEQVR42mP8z8BQDwAF9QH5m2n1LwAAAABJRU5ErkJggg==")?;

    let inserted_id = doc.workbook.images.ensure_unique_name("image", "png");
    doc.workbook.images.insert(
        inserted_id.clone(),
        ImageData {
            bytes: png_bytes.clone(),
            content_type: Some("image/png".to_string()),
        },
    );

    // Append a second floating image drawing.
    {
        let sheet = doc.workbook.sheet_mut(sheet_id).expect("worksheet exists");
        let next_object_id = sheet
            .drawings
            .iter()
            .map(|d| d.id.0)
            .max()
            .unwrap_or(0)
            .saturating_add(1);
        let z_order = sheet.drawings.len() as i32;

        let ext = EmuSize::new(914_400, 914_400);
        let anchor = Anchor::OneCell {
            from: AnchorPoint::new(CellRef::new(0, 2), CellOffset::new(0, 0)),
            ext,
        };

        sheet.drawings.push(DrawingObject {
            id: DrawingObjectId(next_object_id),
            kind: DrawingObjectKind::Image {
                image_id: inserted_id.clone(),
            },
            anchor,
            z_order,
            size: Some(ext),
            preserved: Default::default(),
        });
    }

    let saved = doc.save_to_vec()?;

    // Assert the new media part exists and matches the bytes we inserted.
    let media_part = format!("xl/media/{}", inserted_id.as_str());
    assert_eq!(zip_part_to_bytes(&saved, &media_part), png_bytes);

    // Resolve the worksheet -> drawing relationship.
    let sheet_xml = zip_part_to_string(&saved, "xl/worksheets/sheet1.xml");
    let sheet_doc = Document::parse(&sheet_xml)?;
    let drawing_rid = sheet_doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "drawing")
        .and_then(|n| n.attribute((REL_NS, "id")).or_else(|| n.attribute("r:id")))
        .expect("worksheet should contain <drawing r:id=...>");

    let sheet_rels_xml = zip_part_to_string(&saved, "xl/worksheets/_rels/sheet1.xml.rels");
    let sheet_rels_doc = Document::parse(&sheet_rels_xml)?;
    let drawing_target = sheet_rels_doc
        .descendants()
        .find(|n| {
            n.is_element()
                && n.tag_name().name() == "Relationship"
                && n.attribute("Id") == Some(drawing_rid)
                && n.attribute("Type") == Some(REL_TYPE_DRAWING)
        })
        .and_then(|n| n.attribute("Target"))
        .expect("worksheet rels should contain drawing relationship");

    let drawing_part =
        formula_xlsx::openxml::resolve_target("xl/worksheets/sheet1.xml", drawing_target);
    let drawing_rels_part = formula_xlsx::openxml::rels_part_name(&drawing_part);

    // Assert the drawing rels contains a relationship for the inserted image, and that the drawing
    // XML references it via `r:embed`.
    let drawing_rels_xml = zip_part_to_string(&saved, &drawing_rels_part);
    let drawing_rels_doc = Document::parse(&drawing_rels_xml)?;
    let embed_rid = drawing_rels_doc
        .descendants()
        .find(|n| {
            n.is_element()
                && n.tag_name().name() == "Relationship"
                && n.attribute("Type") == Some(REL_TYPE_IMAGE)
                && n.attribute("Target") == Some(&format!("../media/{}", inserted_id.as_str()))
        })
        .and_then(|n| n.attribute("Id"))
        .expect("drawing rels should contain image relationship for the inserted image");

    let drawing_xml = zip_part_to_string(&saved, &drawing_part);
    let drawing_doc = Document::parse(&drawing_xml)?;
    let saw_embed = drawing_doc.descendants().any(|n| {
        n.is_element()
            && n.tag_name().name() == "blip"
            && n.attribute((REL_NS, "embed"))
                .or_else(|| n.attribute("r:embed"))
                == Some(embed_rid)
    });
    assert!(saw_embed, "drawing XML should reference new r:embed");

    Ok(())
}
