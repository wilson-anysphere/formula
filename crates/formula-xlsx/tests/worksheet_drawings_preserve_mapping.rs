use std::collections::BTreeSet;
use std::io::{Cursor, Read, Write};

use formula_model::drawings::DrawingObjectId;
use formula_model::{CellRef, CellValue};
use formula_xlsx::load_from_bytes;
use roxmltree::Document;
use zip::write::FileOptions;
use zip::ZipArchive;
use zip::ZipWriter;

const REL_TYPE_DRAWING: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";

fn sheet_drawing_relationship(xml: &str) -> Option<(String, String)> {
    let doc = Document::parse(xml).ok()?;
    for node in doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "Relationship")
    {
        if node.attribute("Type")? != REL_TYPE_DRAWING {
            continue;
        }
        let id = node.attribute("Id")?.to_string();
        let target = node.attribute("Target")?.to_string();
        return Some((id, target));
    }
    None
}

fn drawing_part_names(bytes: &[u8]) -> BTreeSet<String> {
    let cursor = Cursor::new(bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut out = BTreeSet::new();
    for idx in 0..archive.len() {
        let Ok(file) = archive.by_index(idx) else {
            continue;
        };
        if file.is_dir() {
            continue;
        }
        let name = file.name();
        let name = name.strip_prefix('/').unwrap_or(name);
        if name.starts_with("xl/drawings/") && name.ends_with(".xml") && !name.contains("/_rels/") {
            out.insert(name.to_string());
        }
    }
    out
}

fn zip_part_bytes(bytes: &[u8], part_name: &str) -> Option<Vec<u8>> {
    let cursor = Cursor::new(bytes);
    let mut archive = ZipArchive::new(cursor).ok()?;
    let part_name = part_name.strip_prefix('/').unwrap_or(part_name);
    let mut file = archive.by_name(part_name).ok()?;
    let mut out = Vec::new();
    file.read_to_end(&mut out).ok()?;
    Some(out)
}

fn zip_part_string(bytes: &[u8], part_name: &str) -> Option<String> {
    let bytes = zip_part_bytes(bytes, part_name)?;
    String::from_utf8(bytes).ok()
}

fn build_two_sheet_drawing_workbook() -> Vec<u8> {
    let base = include_bytes!("../../../fixtures/xlsx/basic/image.xlsx");
    let cursor = Cursor::new(base);
    let mut archive = ZipArchive::new(cursor).expect("open base fixture zip");

    let mut parts = std::collections::BTreeMap::<String, Vec<u8>>::new();
    for idx in 0..archive.len() {
        let mut file = archive.by_index(idx).expect("zip entry");
        if file.is_dir() {
            continue;
        }
        let name = file.name().trim_start_matches('/').to_string();
        let mut buf = Vec::new();
        file.read_to_end(&mut buf).expect("read zip entry");
        parts.insert(name, buf);
    }

    // Duplicate the worksheet + its `.rels` but point at a new drawing part.
    let sheet1 = parts
        .get("xl/worksheets/sheet1.xml")
        .expect("base sheet1.xml")
        .clone();
    parts.insert("xl/worksheets/sheet2.xml".to_string(), sheet1);

    let sheet1_rels = String::from_utf8(
        parts
            .get("xl/worksheets/_rels/sheet1.xml.rels")
            .expect("base sheet1 rels")
            .clone(),
    )
    .expect("utf8 sheet rels");
    let sheet2_rels = sheet1_rels.replace("drawing1.xml", "drawing2.xml");
    parts.insert(
        "xl/worksheets/_rels/sheet2.xml.rels".to_string(),
        sheet2_rels.into_bytes(),
    );

    // Duplicate drawing1 -> drawing2 (and its `.rels`).
    let drawing1 = parts
        .get("xl/drawings/drawing1.xml")
        .expect("base drawing1.xml")
        .clone();
    parts.insert("xl/drawings/drawing2.xml".to_string(), drawing1);
    let drawing1_rels = parts
        .get("xl/drawings/_rels/drawing1.xml.rels")
        .expect("base drawing1 rels")
        .clone();
    parts.insert("xl/drawings/_rels/drawing2.xml.rels".to_string(), drawing1_rels);

    // Replace workbook.xml to reference both worksheets.
    let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    <sheet name="Sheet2" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>
"#;
    parts.insert("xl/workbook.xml".to_string(), workbook_xml.to_vec());

    // Replace workbook.xml.rels with a consistent relationship table.
    let workbook_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
"#;
    parts.insert("xl/_rels/workbook.xml.rels".to_string(), workbook_rels.to_vec());

    // Update `[Content_Types].xml` so the new parts have overrides.
    let content_types = String::from_utf8(
        parts
            .get("[Content_Types].xml")
            .expect("base content types")
            .clone(),
    )
    .expect("utf8 [Content_Types].xml");
    let content_types = content_types.replace(
        r#"<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#,
        r#"<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#,
    );
    let content_types = content_types.replace(
        r#"<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>"#,
        r#"<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  <Override PartName="/xl/drawings/drawing2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>"#,
    );
    parts.insert("[Content_Types].xml".to_string(), content_types.into_bytes());

    // Repack into a new XLSX zip.
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);
    for (name, bytes) in parts {
        zip.start_file(name, options).expect("zip start_file");
        zip.write_all(&bytes).expect("zip write_all");
    }
    zip.finish().expect("zip finish").into_inner()
}

#[test]
fn xlsx_document_roundtrip_preserves_sheet_drawing_relationship_mapping() {
    let input = include_bytes!("../../../fixtures/xlsx/basic/image.xlsx");
    let doc = load_from_bytes(input).expect("load fixture");

    let orig_rels_xml = zip_part_string(input, "xl/worksheets/_rels/sheet1.xml.rels")
        .expect("fixture sheet1 rels");
    let (orig_id, orig_target) =
        sheet_drawing_relationship(&orig_rels_xml).expect("fixture drawing relationship");

    let saved = doc.save_to_vec().expect("save");
    let _doc2 = load_from_bytes(&saved).expect("reload");

    let out_rels_xml = zip_part_string(&saved, "xl/worksheets/_rels/sheet1.xml.rels")
        .expect("output sheet1 rels");
    let (out_id, out_target) =
        sheet_drawing_relationship(&out_rels_xml).expect("output drawing relationship");

    assert_eq!(
        out_id, orig_id,
        "drawing relationship id should be preserved"
    );
    assert_eq!(
        out_target, orig_target,
        "drawing relationship target should be preserved"
    );

    let in_parts = drawing_part_names(input);
    let out_parts = drawing_part_names(&saved);
    assert_eq!(
        out_parts, in_parts,
        "drawing part names should be preserved"
    );
}

#[test]
fn xlsx_document_roundtrip_preserves_multi_sheet_drawing_relationship_mapping() {
    let input = build_two_sheet_drawing_workbook();
    let doc = load_from_bytes(&input).expect("load synthetic workbook");

    let sheet1_rels = zip_part_string(&input, "xl/worksheets/_rels/sheet1.xml.rels")
        .expect("input sheet1 rels");
    let (sheet1_orig_id, sheet1_orig_target) =
        sheet_drawing_relationship(&sheet1_rels).expect("sheet1 drawing rel");
    let sheet2_rels = zip_part_string(&input, "xl/worksheets/_rels/sheet2.xml.rels")
        .expect("input sheet2 rels");
    let (sheet2_orig_id, sheet2_orig_target) =
        sheet_drawing_relationship(&sheet2_rels).expect("sheet2 drawing rel");

    assert_eq!(sheet1_orig_target, "../drawings/drawing1.xml");
    assert_eq!(sheet2_orig_target, "../drawings/drawing2.xml");

    let saved = doc.save_to_vec().expect("save");
    let _doc2 = load_from_bytes(&saved).expect("reload saved workbook");

    let sheet1_out_rels = zip_part_string(&saved, "xl/worksheets/_rels/sheet1.xml.rels")
        .expect("output sheet1 rels");
    let (sheet1_out_id, sheet1_out_target) =
        sheet_drawing_relationship(&sheet1_out_rels).expect("output sheet1 drawing rel");
    let sheet2_out_rels = zip_part_string(&saved, "xl/worksheets/_rels/sheet2.xml.rels")
        .expect("output sheet2 rels");
    let (sheet2_out_id, sheet2_out_target) =
        sheet_drawing_relationship(&sheet2_out_rels).expect("output sheet2 drawing rel");

    assert_eq!(sheet1_out_id, sheet1_orig_id);
    assert_eq!(sheet1_out_target, sheet1_orig_target);
    assert_eq!(sheet2_out_id, sheet2_orig_id);
    assert_eq!(sheet2_out_target, sheet2_orig_target);

    let in_parts = drawing_part_names(&input);
    let out_parts = drawing_part_names(&saved);
    assert_eq!(out_parts, in_parts, "drawing part names should be preserved");
}

#[test]
fn xlsx_document_roundtrip_preserves_drawing_mapping_after_sheet_reorder() {
    let input = build_two_sheet_drawing_workbook();
    let mut doc = load_from_bytes(&input).expect("load synthetic workbook");

    let sheet2_id = doc
        .workbook
        .sheets
        .iter()
        .find(|s| s.name == "Sheet2")
        .expect("Sheet2")
        .id;
    assert!(
        doc.workbook.reorder_sheet(sheet2_id, 0),
        "expected reorder_sheet to succeed"
    );
    assert_eq!(
        doc.workbook.sheets[0].name, "Sheet2",
        "sanity check: sheet order should change in the in-memory model"
    );

    let saved = doc.save_to_vec().expect("save");
    let workbook_xml = zip_part_string(&saved, "xl/workbook.xml").expect("workbook.xml");
    let sheet1_pos = workbook_xml
        .find("name=\"Sheet1\"")
        .expect("Sheet1 in workbook.xml");
    let sheet2_pos = workbook_xml
        .find("name=\"Sheet2\"")
        .expect("Sheet2 in workbook.xml");
    assert!(
        sheet2_pos < sheet1_pos,
        "expected Sheet2 to appear before Sheet1 after reordering"
    );

    let sheet1_rels = zip_part_string(&saved, "xl/worksheets/_rels/sheet1.xml.rels")
        .expect("output sheet1 rels");
    let (sheet1_out_id, sheet1_out_target) =
        sheet_drawing_relationship(&sheet1_rels).expect("output sheet1 drawing rel");
    let sheet2_rels = zip_part_string(&saved, "xl/worksheets/_rels/sheet2.xml.rels")
        .expect("output sheet2 rels");
    let (sheet2_out_id, sheet2_out_target) =
        sheet_drawing_relationship(&sheet2_rels).expect("output sheet2 drawing rel");

    assert_eq!(sheet1_out_id, "rId1");
    assert_eq!(sheet1_out_target, "../drawings/drawing1.xml");
    assert_eq!(sheet2_out_id, "rId1");
    assert_eq!(sheet2_out_target, "../drawings/drawing2.xml");
}

#[test]
fn xlsx_document_roundtrip_preserves_drawing_mapping_after_cell_edit() {
    let input = build_two_sheet_drawing_workbook();
    let mut doc = load_from_bytes(&input).expect("load synthetic workbook");

    let sheet1_id = doc
        .workbook
        .sheets
        .iter()
        .find(|s| s.name == "Sheet1")
        .expect("Sheet1")
        .id;
    assert!(
        doc.set_cell_value(sheet1_id, CellRef::new(0, 1), CellValue::Number(42.0)),
        "expected set_cell_value to succeed"
    );

    let saved = doc.save_to_vec().expect("save");
    let _doc2 = load_from_bytes(&saved).expect("reload");

    let sheet1_rels = zip_part_string(&saved, "xl/worksheets/_rels/sheet1.xml.rels")
        .expect("output sheet1 rels");
    let (sheet1_out_id, sheet1_out_target) =
        sheet_drawing_relationship(&sheet1_rels).expect("output sheet1 drawing rel");
    let sheet2_rels = zip_part_string(&saved, "xl/worksheets/_rels/sheet2.xml.rels")
        .expect("output sheet2 rels");
    let (sheet2_out_id, sheet2_out_target) =
        sheet_drawing_relationship(&sheet2_rels).expect("output sheet2 drawing rel");

    assert_eq!(sheet1_out_id, "rId1");
    assert_eq!(sheet1_out_target, "../drawings/drawing1.xml");
    assert_eq!(sheet2_out_id, "rId1");
    assert_eq!(sheet2_out_target, "../drawings/drawing2.xml");

    let in_parts = drawing_part_names(&input);
    let out_parts = drawing_part_names(&saved);
    assert_eq!(out_parts, in_parts, "drawing part names should be preserved");
}

#[test]
fn xlsx_document_roundtrip_preserves_drawing_mapping_after_drawing_edit() {
    let input = build_two_sheet_drawing_workbook();
    let mut doc = load_from_bytes(&input).expect("load synthetic workbook");

    let sheet1 = doc
        .workbook
        .sheets
        .iter_mut()
        .find(|s| s.name == "Sheet1")
        .expect("Sheet1");
    let existing = sheet1
        .drawings
        .first()
        .cloned()
        .expect("expected Sheet1 to contain at least one drawing object");
    let mut cloned = existing.clone();
    cloned.id = DrawingObjectId(existing.id.0.saturating_add(1));
    // Force the writer to allocate a new embed relationship + pic XML for this cloned object so we
    // exercise the drawing-part rewrite path, not just a no-op clone.
    cloned.preserved.remove("xlsx.embed_rel_id");
    cloned.preserved.remove("xlsx.pic_xml");
    sheet1.drawings.push(cloned);

    let saved = doc.save_to_vec().expect("save");
    let _doc2 = load_from_bytes(&saved).expect("reload");

    let sheet1_rels = zip_part_string(&saved, "xl/worksheets/_rels/sheet1.xml.rels")
        .expect("output sheet1 rels");
    let (sheet1_out_id, sheet1_out_target) =
        sheet_drawing_relationship(&sheet1_rels).expect("output sheet1 drawing rel");
    let sheet2_rels = zip_part_string(&saved, "xl/worksheets/_rels/sheet2.xml.rels")
        .expect("output sheet2 rels");
    let (sheet2_out_id, sheet2_out_target) =
        sheet_drawing_relationship(&sheet2_rels).expect("output sheet2 drawing rel");

    assert_eq!(sheet1_out_id, "rId1");
    assert_eq!(sheet1_out_target, "../drawings/drawing1.xml");
    assert_eq!(sheet2_out_id, "rId1");
    assert_eq!(sheet2_out_target, "../drawings/drawing2.xml");

    let in_parts = drawing_part_names(&input);
    let out_parts = drawing_part_names(&saved);
    assert_eq!(out_parts, in_parts, "drawing part names should be preserved");
}
