use base64::Engine;
use formula_model::drawings::{
    Anchor, AnchorPoint, CellOffset, DrawingObject, DrawingObjectId, DrawingObjectKind, EmuSize,
    ImageData,
};
use formula_model::CellRef;
use formula_xlsx::drawings::DrawingPart;
use formula_xlsx::{load_from_bytes, XlsxDocument, XlsxPackage};
use roxmltree::Document;
use std::collections::HashMap;

const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL_TYPE_DRAWING: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";

#[test]
fn xlsxdocument_roundtrip_preserves_floating_images_fixture() {
    let fixture = include_bytes!("../../../fixtures/xlsx/basic/image.xlsx");

    let doc = load_from_bytes(fixture).expect("load fixture");
    let saved = doc.save_to_vec().expect("save");

    let pkg = XlsxPackage::from_bytes(&saved).expect("open saved workbook");
    assert!(
        pkg.part("xl/drawings/drawing1.xml").is_some(),
        "expected drawing part to be present after save"
    );
    assert!(
        pkg.part("xl/media/image1.png").is_some(),
        "expected image media part to be present after save"
    );

    let reloaded = load_from_bytes(&saved).expect("reload");
    let sheet = &reloaded.workbook.sheets[0];
    assert!(
        sheet
            .drawings
            .iter()
            .any(|o| matches!(o.kind, DrawingObjectKind::Image { .. })),
        "expected worksheet.drawings to contain an image object"
    );
}

#[test]
fn xlsxdocument_writes_newly_inserted_sheet_drawing_image() {
    // 1x1 transparent PNG.
    let png_bytes = base64::engine::general_purpose::STANDARD
        .decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/58HAQUBAO3+2NoAAAAASUVORK5CYII=")
        .expect("valid base64 png");

    let mut workbook = formula_model::Workbook::new();
    let sheet_id = workbook.add_sheet("Sheet1").expect("add sheet");

    let image_id = workbook.images.ensure_unique_name("image", "png");
    workbook.images.insert(
        image_id.clone(),
        ImageData {
            bytes: png_bytes.clone(),
            content_type: Some("image/png".to_string()),
        },
    );

    // Build a drawing object via the existing `DrawingPart::insert_image_object` API.
    let (mut drawing_part, _sheet_drawing_rid) = DrawingPart::create_new(0).expect("create part");
    let anchor = Anchor::OneCell {
        from: AnchorPoint::new(CellRef::new(0, 0), CellOffset::new(0, 0)),
        ext: EmuSize::new(914_400, 914_400),
    };
    drawing_part.insert_image_object(&image_id, anchor);

    workbook.sheet_mut(sheet_id).expect("sheet exists").drawings = drawing_part.objects.clone();

    let doc = XlsxDocument::new(workbook);
    let saved = doc.save_to_vec().expect("save");

    let pkg = XlsxPackage::from_bytes(&saved).expect("open saved workbook");
    assert!(
        pkg.part("xl/drawings/drawing1.xml").is_some(),
        "expected drawing part to be emitted for newly inserted image"
    );

    let media_part_name = format!("xl/media/{}", image_id.as_str());
    assert_eq!(
        pkg.part(&media_part_name),
        Some(png_bytes.as_slice()),
        "expected image bytes to be present in {media_part_name}"
    );

    let reloaded = load_from_bytes(&saved).expect("reload");
    let sheet = &reloaded.workbook.sheets[0];
    assert!(
        sheet.drawings.iter().any(|o| matches!(
            &o.kind,
            DrawingObjectKind::Image { image_id: got } if got == &image_id
        )),
        "expected worksheet.drawings to contain the inserted image object"
    );
}

#[test]
fn xlsxdocument_can_append_image_to_existing_sheet_drawing() {
    let fixture = include_bytes!("../../../fixtures/xlsx/basic/image.xlsx");

    // 1x1 transparent PNG.
    let png_bytes = base64::engine::general_purpose::STANDARD
        .decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/58HAQUBAO3+2NoAAAAASUVORK5CYII=")
        .expect("valid base64 png");

    let mut doc = load_from_bytes(fixture).expect("load fixture");

    // Add a second image to the workbook image store.
    let image_id = doc.workbook.images.ensure_unique_name("image", "png");
    doc.workbook.images.insert(
        image_id.clone(),
        ImageData {
            bytes: png_bytes.clone(),
            content_type: Some("image/png".to_string()),
        },
    );

    // Append a new image drawing object to the existing sheet drawings.
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
        from: AnchorPoint::new(CellRef::new(1, 1), CellOffset::new(0, 0)),
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
        preserved: HashMap::new(),
    });

    let saved = doc.save_to_vec().expect("save");

    let pkg = XlsxPackage::from_bytes(&saved).expect("open saved workbook");
    assert!(
        pkg.part("xl/drawings/drawing1.xml").is_some(),
        "expected drawing part to still exist after adding an image"
    );
    assert!(
        pkg.part("xl/media/image1.png").is_some(),
        "expected original image media part to remain present"
    );
    assert_eq!(
        pkg.part(&format!("xl/media/{}", image_id.as_str())),
        Some(png_bytes.as_slice()),
        "expected appended image bytes to be present"
    );

    let reloaded = load_from_bytes(&saved).expect("reload");
    let sheet = &reloaded.workbook.sheets[0];
    assert!(
        sheet.drawings.iter().any(|o| matches!(
            &o.kind,
            DrawingObjectKind::Image { image_id: got } if got == &image_id
        )),
        "expected worksheet.drawings to contain the appended image object"
    );
}

#[test]
fn xlsxdocument_can_duplicate_image_without_adding_new_media() {
    let fixture = include_bytes!("../../../fixtures/xlsx/basic/image.xlsx");

    let mut doc = load_from_bytes(fixture).expect("load fixture");

    let sheet = doc
        .workbook
        .sheets
        .get_mut(0)
        .expect("fixture has at least one sheet");

    let existing_image_id = sheet
        .drawings
        .iter()
        .find_map(|o| match &o.kind {
            DrawingObjectKind::Image { image_id } => Some(image_id.clone()),
            _ => None,
        })
        .expect("fixture should contain at least one image drawing");

    // Duplicate the existing image drawing, reusing the same `image_id` (no new `xl/media/*`
    // part is introduced). This should still cause the drawing part to be rewritten.
    let next_id = sheet
        .drawings
        .iter()
        .map(|o| o.id.0)
        .max()
        .unwrap_or(0)
        .saturating_add(1);

    let ext = EmuSize::new(914_400, 914_400);
    let anchor = Anchor::OneCell {
        from: AnchorPoint::new(CellRef::new(2, 2), CellOffset::new(0, 0)),
        ext,
    };
    sheet.drawings.push(DrawingObject {
        id: DrawingObjectId(next_id),
        kind: DrawingObjectKind::Image {
            image_id: existing_image_id.clone(),
        },
        anchor,
        z_order: sheet.drawings.len() as i32,
        size: Some(ext),
        preserved: HashMap::new(),
    });

    let saved = doc.save_to_vec().expect("save");

    // Original media should still exist (no new media part should be required).
    let pkg = XlsxPackage::from_bytes(&saved).expect("open saved workbook");
    assert!(
        pkg.part("xl/media/image1.png").is_some(),
        "expected original image media part to remain present"
    );

    let reloaded = load_from_bytes(&saved).expect("reload");
    let sheet = &reloaded.workbook.sheets[0];
    let image_count = sheet
        .drawings
        .iter()
        .filter(|o| matches!(&o.kind, DrawingObjectKind::Image { image_id } if image_id == &existing_image_id))
        .count();
    assert!(
        image_count >= 2,
        "expected duplicated image to round-trip; got {image_count} occurrences of {existing_image_id:?}"
    );
}

#[test]
fn xlsxdocument_can_remove_sheet_drawings() {
    let fixture = include_bytes!("../../../fixtures/xlsx/basic/image.xlsx");

    let mut doc = load_from_bytes(fixture).expect("load fixture");
    doc.workbook.sheets[0].drawings.clear();

    let saved = doc.save_to_vec().expect("save");

    let pkg = XlsxPackage::from_bytes(&saved).expect("open saved workbook");
    let sheet_xml = std::str::from_utf8(
        pkg.part("xl/worksheets/sheet1.xml")
            .expect("sheet xml should exist"),
    )
    .expect("sheet xml utf8");
    let sheet_doc = Document::parse(sheet_xml).expect("parse sheet xml");
    assert!(
        !sheet_doc
            .descendants()
            .any(|n| n.is_element() && n.tag_name().name() == "drawing"),
        "expected worksheet XML to not contain a <drawing> element after removal, got:\n{sheet_xml}"
    );

    if let Some(rels_bytes) = pkg.part("xl/worksheets/_rels/sheet1.xml.rels") {
        let rels_xml = std::str::from_utf8(rels_bytes).expect("rels utf8");
        let rels_doc = Document::parse(rels_xml).expect("parse rels xml");
        assert!(
            !rels_doc.descendants().any(|n| {
                n.is_element()
                    && n.tag_name().name() == "Relationship"
                    && n.attribute("Type")
                        == Some(
                            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
                        )
            }),
            "expected worksheet .rels to not contain a drawing relationship after removal, got:\n{rels_xml}"
        );
    }

    let reloaded = load_from_bytes(&saved).expect("reload");
    assert!(
        reloaded.workbook.sheets[0].drawings.is_empty(),
        "expected drawings to be removed on reload"
    );
}

#[test]
fn xlsxdocument_remove_drawings_only_removes_referenced_relationship() {
    let fixture = include_bytes!("../../../fixtures/xlsx/basic/image.xlsx");

    let mut pkg = XlsxPackage::from_bytes(fixture).expect("open fixture");
    let rels_part = "xl/worksheets/_rels/sheet1.xml.rels";
    let rels_xml = std::str::from_utf8(pkg.part(rels_part).expect("sheet rels")).expect("utf8");
    let injected = format!(
        "Relationship Id=\"rId999\" Type=\"{}\" Target=\"../drawings/drawing999.xml\"",
        REL_TYPE_DRAWING
    );
    let rels_xml = inject_relationship_first(rels_xml, &injected);
    pkg.set_part(rels_part, rels_xml.into_bytes());

    let sheet_xml = std::str::from_utf8(
        pkg.part("xl/worksheets/sheet1.xml")
            .expect("sheet xml exists"),
    )
    .expect("utf8");
    let drawing_rid = drawing_rel_id_from_sheet_xml(sheet_xml).expect("sheet has <drawing>");

    let modified_bytes = pkg.write_to_bytes().expect("write modified fixture");
    let mut doc = load_from_bytes(&modified_bytes).expect("load modified fixture");

    // Clear drawings from the model, which should remove the sheet-level <drawing> relationship.
    doc.workbook.sheets[0].drawings.clear();
    let saved = doc.save_to_vec().expect("save");

    let saved_pkg = XlsxPackage::from_bytes(&saved).expect("open saved workbook");
    let saved_rels_xml =
        std::str::from_utf8(saved_pkg.part(rels_part).expect("rels exists")).expect("rels utf8");

    assert!(
        find_relationship(saved_rels_xml, "rId999").is_some(),
        "expected unrelated drawing relationship rId999 to remain, got:\n{saved_rels_xml}"
    );
    assert!(
        find_relationship(saved_rels_xml, &drawing_rid).is_none(),
        "expected referenced drawing relationship {drawing_rid} to be removed, got:\n{saved_rels_xml}"
    );
}

#[test]
fn xlsxdocument_preserves_existing_drawing_relationship_by_id() {
    let fixture = include_bytes!("../../../fixtures/xlsx/basic/image.xlsx");

    let mut pkg = XlsxPackage::from_bytes(fixture).expect("open fixture");
    let rels_part = "xl/worksheets/_rels/sheet1.xml.rels";
    let rels_xml = std::str::from_utf8(pkg.part(rels_part).expect("sheet rels")).expect("utf8");
    let injected = format!(
        "Relationship Id=\"rId999\" Type=\"{}\" Target=\"../drawings/drawing999.xml\"",
        REL_TYPE_DRAWING
    );
    let rels_xml = inject_relationship_first(rels_xml, &injected);
    pkg.set_part(rels_part, rels_xml.into_bytes());

    let sheet_xml = std::str::from_utf8(
        pkg.part("xl/worksheets/sheet1.xml")
            .expect("sheet xml exists"),
    )
    .expect("utf8");
    let drawing_rid = drawing_rel_id_from_sheet_xml(sheet_xml).expect("sheet has <drawing>");
    let (_, drawing_target) = find_relationship(
        std::str::from_utf8(pkg.part(rels_part).unwrap()).unwrap(),
        &drawing_rid,
    )
    .expect("referenced drawing relationship exists");

    let modified_bytes = pkg.write_to_bytes().expect("write modified fixture");
    let mut doc = load_from_bytes(&modified_bytes).expect("load modified fixture");

    // Duplicate an existing image drawing without introducing new media. This should trigger
    // `drawings_need_emit`, and the writer should preserve the relationship referenced by the
    // sheet's <drawing r:id="..."> instead of picking the first relationship of Type=drawing.
    let sheet = doc
        .workbook
        .sheets
        .get_mut(0)
        .expect("fixture has at least one sheet");
    let existing_image_id = sheet
        .drawings
        .iter()
        .find_map(|o| match &o.kind {
            DrawingObjectKind::Image { image_id } => Some(image_id.clone()),
            _ => None,
        })
        .expect("fixture should contain at least one image drawing");

    let next_id = sheet
        .drawings
        .iter()
        .map(|o| o.id.0)
        .max()
        .unwrap_or(0)
        .saturating_add(1);

    let ext = EmuSize::new(914_400, 914_400);
    let anchor = Anchor::OneCell {
        from: AnchorPoint::new(CellRef::new(2, 2), CellOffset::new(0, 0)),
        ext,
    };
    sheet.drawings.push(DrawingObject {
        id: DrawingObjectId(next_id),
        kind: DrawingObjectKind::Image {
            image_id: existing_image_id.clone(),
        },
        anchor,
        z_order: sheet.drawings.len() as i32,
        size: Some(ext),
        preserved: HashMap::new(),
    });

    let saved = doc.save_to_vec().expect("save");

    let saved_pkg = XlsxPackage::from_bytes(&saved).expect("open saved workbook");
    let saved_sheet_xml = std::str::from_utf8(
        saved_pkg
            .part("xl/worksheets/sheet1.xml")
            .expect("saved sheet xml exists"),
    )
    .expect("utf8");
    let saved_rid =
        drawing_rel_id_from_sheet_xml(saved_sheet_xml).expect("saved sheet has <drawing>");
    assert_eq!(
        saved_rid, drawing_rid,
        "expected writer to preserve sheet drawing r:id; got {saved_rid}, expected {drawing_rid}\n{saved_sheet_xml}"
    );

    let saved_rels_xml =
        std::str::from_utf8(saved_pkg.part(rels_part).expect("rels exists")).expect("rels utf8");
    let (saved_ty, saved_target) =
        find_relationship(saved_rels_xml, &drawing_rid).expect("drawing relationship exists");
    assert_eq!(saved_ty, REL_TYPE_DRAWING);
    assert_eq!(
        saved_target, drawing_target,
        "expected drawing relationship target to be preserved"
    );

    assert!(
        saved_pkg.part("xl/drawings/drawing999.xml").is_none(),
        "unexpectedly created xl/drawings/drawing999.xml due to relationship-by-type lookup"
    );
}

fn drawing_rel_id_from_sheet_xml(sheet_xml: &str) -> Option<String> {
    let doc = Document::parse(sheet_xml).ok()?;
    let drawing = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "drawing")?;
    Some(
        drawing
            .attribute((REL_NS, "id"))
            .or_else(|| drawing.attribute("r:id"))
            .or_else(|| drawing.attribute("id"))?
            .to_string(),
    )
}

fn find_relationship(xml: &str, id: &str) -> Option<(String, String)> {
    let doc = Document::parse(xml).ok()?;
    for node in doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "Relationship")
    {
        if node.attribute("Id")? != id {
            continue;
        }
        return Some((
            node.attribute("Type")?.to_string(),
            node.attribute("Target")?.to_string(),
        ));
    }
    None
}

fn inject_relationship_first(rels_xml: &str, relationship_attrs: &str) -> String {
    // Insert a new <Relationship .../> element before the first existing relationship element so
    // relationship-by-type lookups would return it.
    let insertion = format!("<{relationship_attrs}/>\n");
    // Match the actual `<Relationship ...>` elements, not the `<Relationships>` root element.
    if let Some(idx) = rels_xml.find("<Relationship ") {
        let mut out = String::new();
        out.push_str(&rels_xml[..idx]);
        out.push_str(&insertion);
        out.push_str(&rels_xml[idx..]);
        return out;
    }

    // Fall back to inserting before the closing root element.
    if let Some(idx) = rels_xml.rfind("</Relationships>") {
        let mut out = String::new();
        out.push_str(&rels_xml[..idx]);
        out.push_str(&insertion);
        out.push_str(&rels_xml[idx..]);
        return out;
    }

    format!("{rels_xml}\n{insertion}")
}
