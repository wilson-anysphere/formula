use std::collections::HashSet;
use std::io::{Cursor, Write};

use base64::Engine;
use formula_model::{CellRef, CellValue};
use formula_xlsx::{
    patch_xlsx_streaming_workbook_cell_patches, CellPatch, WorkbookCellPatches, XlsxPackage,
};
use rust_xlsxwriter::{Format, Image, Workbook};

fn richdata_part_names(pkg: &XlsxPackage) -> Vec<String> {
    pkg.part_names()
        .filter(|name| {
            name.starts_with("xl/richData/")
                // Relationship parts are covered by the standard `.rels` default content type and
                // typically do not appear as explicit `<Override>` entries.
                && !name.starts_with("xl/richData/_rels/")
                && !name.ends_with(".rels")
        })
        .map(|name| name.to_string())
        .collect()
}

fn content_type_overrides(pkg: &XlsxPackage) -> HashSet<String> {
    let ct_xml = std::str::from_utf8(pkg.part("[Content_Types].xml").unwrap())
        .expect("[Content_Types].xml must be utf-8");
    let doc = roxmltree::Document::parse(ct_xml).expect("parse [Content_Types].xml");
    doc.descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "Override")
        .filter_map(|n| n.attribute("PartName"))
        .map(|s| s.to_string())
        .collect()
}

fn assert_embedded_image_richdata_present(pkg: &XlsxPackage) {
    let rich_parts = richdata_part_names(pkg);
    assert!(
        !rich_parts.is_empty(),
        "expected embedded image workbook to contain xl/richData/* parts, but none were found. Parts: {:?}",
        pkg.part_names().collect::<Vec<_>>()
    );

    assert!(
        pkg.part("xl/metadata.xml").is_some(),
        "expected embedded image workbook to contain xl/metadata.xml"
    );

    assert!(
        pkg.part_names().any(|name| name.starts_with("xl/media/")),
        "expected embedded image workbook to contain xl/media/* image parts"
    );

    let rels_xml = std::str::from_utf8(
        pkg.part("xl/_rels/workbook.xml.rels")
            .expect("workbook rels must exist"),
    )
    .expect("workbook.xml.rels must be utf-8");
    let doc = roxmltree::Document::parse(rels_xml).expect("parse workbook.xml.rels");
    let rel_types: Vec<String> = doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "Relationship")
        .filter_map(|n| n.attribute("Type"))
        .map(|t| t.to_string())
        .collect();
    assert!(
        rel_types.iter().any(|t| t.contains("sheetMetadata")),
        "expected workbook.xml.rels to contain a sheetMetadata relationship type, got: {rel_types:?}"
    );
    assert!(
        rel_types.iter().any(|t| t.contains("rdRichValue")),
        "expected workbook.xml.rels to contain at least one rdRichValue* relationship type, got: {rel_types:?}"
    );
    assert!(
        rel_types.iter().any(|t| t.contains("richValueRel")),
        "expected workbook.xml.rels to contain a richValueRel relationship type, got: {rel_types:?}"
    );

    let overrides = content_type_overrides(pkg);
    assert!(
        overrides.contains("/xl/metadata.xml"),
        "expected [Content_Types].xml to contain an override for /xl/metadata.xml; got: {overrides:?}"
    );
    for part in &rich_parts {
        let part_name = format!("/{part}");
        assert!(
            overrides.contains(&part_name),
            "expected [Content_Types].xml to contain an override for {part_name} (embedded image richData part)"
        );
    }
}

#[test]
fn streaming_formula_recalc_preserves_embedded_images_richdata_relationships_and_content_types(
) -> Result<(), Box<dyn std::error::Error>> {
    // 1) Generate an XLSX with an embedded image-in-cell (richData) via rust_xlsxwriter.
    let png_bytes = base64::engine::general_purpose::STANDARD.decode(
        "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAQAAADZc7J/AAAADElEQVR42mP8z8BQDwAF9QH5m2n1LwAAAABJRU5ErkJggg==",
    )?;
    let mut png_file = tempfile::NamedTempFile::new()?;
    png_file.write_all(&png_bytes)?;

    let image = Image::new(png_file.path())?;
    let format = Format::new();

    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.embed_image_with_format(1, 1, &image, &format)?;

    let input_bytes = workbook.save_to_buffer()?;

    let input_pkg = XlsxPackage::from_bytes(&input_bytes)?;
    assert_embedded_image_richdata_present(&input_pkg);

    // 2) Apply a streaming patch containing a formula edit.
    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("A1")?,
        CellPatch::set_value_with_formula(CellValue::Number(2.0), "=1+1"),
    );

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming_workbook_cell_patches(Cursor::new(input_bytes), &mut out, &patches)?;
    let out_bytes = out.into_inner();

    // 3) Re-load output and assert recalc-policy rewrites didn't drop richData support.
    let out_pkg = XlsxPackage::from_bytes(&out_bytes)?;
    let out_workbook_xml = std::str::from_utf8(out_pkg.part("xl/workbook.xml").unwrap())
        .expect("xl/workbook.xml must be utf-8");
    assert!(
        out_workbook_xml.contains(r#"fullCalcOnLoad="1""#),
        "expected formula edit to trigger recalc policy (fullCalcOnLoad=1), got:\n{out_workbook_xml}"
    );
    assert_embedded_image_richdata_present(&out_pkg);

    Ok(())
}
