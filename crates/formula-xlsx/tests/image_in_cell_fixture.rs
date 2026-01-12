use std::collections::HashSet;
use std::io::{Cursor, Read};

use formula_model::drawings::ImageId;
use zip::ZipArchive;

/// Validates the `fixtures/xlsx/basic/image-in-cell.xlsx` fixture structure.
///
/// This fixture demonstrates Excel "images in cells" via the Rich Value pipeline:
/// `vm=` on cells + `xl/metadata.xml` + `xl/richData/*` + `xl/media/*`.
#[test]
fn image_in_cell_fixture_has_expected_rich_value_parts() {
    let fixture_path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/xlsx/basic/image-in-cell.xlsx");
    let bytes = std::fs::read(&fixture_path)
        .unwrap_or_else(|e| panic!("read fixture {}: {e}", fixture_path.display()));

    let mut archive = ZipArchive::new(Cursor::new(&bytes)).expect("open xlsx as zip");
    let mut names: HashSet<String> = HashSet::new();
    for i in 0..archive.len() {
        let file = archive.by_index(i).expect("zip entry");
        names.insert(file.name().to_string());
    }

    // Core parts for the image-in-cell (rich value) pipeline.
    for expected in [
        "xl/metadata.xml",
        "xl/richData/richValueRel.xml",
        "xl/richData/_rels/richValueRel.xml.rels",
        "xl/richData/rdrichvalue.xml",
        "xl/richData/rdrichvaluestructure.xml",
        "xl/richData/rdRichValueTypes.xml",
        "xl/media/image1.png",
        "xl/media/image2.png",
    ] {
        assert!(names.contains(expected), "missing expected part: {expected}");
    }

    // Document the observed structure: this Excel fixture does *not* use xl/cellimages.xml.
    assert!(
        !names.contains("xl/cellimages.xml"),
        "fixture unexpectedly contains xl/cellimages.xml; update test expectations"
    );

    // Confirm this is an Excel-produced workbook (not a synthetic fixture).
    let mut app_props = String::new();
    archive
        .by_name("docProps/app.xml")
        .expect("docProps/app.xml")
        .read_to_string(&mut app_props)
        .expect("read docProps/app.xml");
    assert!(
        app_props.contains("<Application>Microsoft Excel</Application>"),
        "expected docProps/app.xml Application=Microsoft Excel, got: {app_props}"
    );

    // Ensure our rich-data parser can actually resolve the embedded images for this fixture.
    let pkg = formula_xlsx::XlsxPackage::from_bytes(&bytes).expect("parse xlsx package");
    let embedded = formula_xlsx::extract_embedded_images(&pkg).expect("extract embedded images");
    assert_eq!(
        embedded.len(),
        3,
        "expected 3 embedded image cells (Sheet1!B2,B3,B4)"
    );

    let image1 = pkg
        .part("xl/media/image1.png")
        .expect("xl/media/image1.png")
        .to_vec();
    let image2 = pkg
        .part("xl/media/image2.png")
        .expect("xl/media/image2.png")
        .to_vec();

    let mut by_cell: std::collections::HashMap<formula_model::CellRef, formula_xlsx::EmbeddedImageCell> =
        std::collections::HashMap::new();
    for entry in embedded {
        assert_eq!(entry.sheet_part, "xl/worksheets/sheet1.xml");
        by_cell.insert(entry.cell, entry);
    }

    let b2 = formula_model::CellRef::from_a1("B2").unwrap();
    let b3 = formula_model::CellRef::from_a1("B3").unwrap();
    let b4 = formula_model::CellRef::from_a1("B4").unwrap();

    let e_b2 = by_cell.get(&b2).expect("expected embedded image at B2");
    let e_b3 = by_cell.get(&b3).expect("expected embedded image at B3");
    let e_b4 = by_cell.get(&b4).expect("expected embedded image at B4");

    assert_eq!(e_b2.image_target, "xl/media/image1.png");
    assert_eq!(e_b3.image_target, "xl/media/image1.png");
    assert_eq!(e_b4.image_target, "xl/media/image2.png");

    assert_eq!(e_b2.bytes, image1);
    assert_eq!(e_b3.bytes, image1);
    assert_eq!(e_b4.bytes, image2);

    // This fixture stores "Place in Cell" images as decorative (CalcOrigin=5) with no alt text.
    assert!(e_b2.decorative);
    assert!(e_b3.decorative);
    assert!(e_b4.decorative);
    assert_eq!(e_b2.alt_text, None);
    assert_eq!(e_b3.alt_text, None);
    assert_eq!(e_b4.alt_text, None);

    // workbook.xml.rels should link the metadata + richData parts at the workbook level.
    let mut workbook_rels = String::new();
    archive
        .by_name("xl/_rels/workbook.xml.rels")
        .expect("xl/_rels/workbook.xml.rels")
        .read_to_string(&mut workbook_rels)
        .expect("read workbook rels");
    assert!(workbook_rels.contains("sheetMetadata"));
    assert!(workbook_rels.contains("richValueRel"));
    assert!(workbook_rels.contains("rdRichValue"));
    assert!(workbook_rels.contains("rdRichValueStructure"));
    assert!(workbook_rels.contains("rdRichValueTypes"));

    // Worksheet should contain cells with `vm=` pointing into `xl/metadata.xml`.
    let mut sheet1 = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")
        .expect("xl/worksheets/sheet1.xml")
        .read_to_string(&mut sheet1)
        .expect("read sheet1.xml");
    assert!(sheet1.contains("vm=\"1\""), "expected vm=\"1\" in sheet1.xml");
    assert!(sheet1.contains("vm=\"2\""), "expected vm=\"2\" in sheet1.xml");

    // metadata.xml should contain rich value bundle indices (xlrd:rvb).
    let mut metadata = String::new();
    archive
        .by_name("xl/metadata.xml")
        .expect("xl/metadata.xml")
        .read_to_string(&mut metadata)
        .expect("read metadata.xml");
    assert!(
        metadata.contains("XLRICHVALUE"),
        "expected XLRICHVALUE metadata type"
    );
    assert!(
        metadata.contains("xlrd:rvb i=\"0\""),
        "expected xlrd:rvb i=\"0\" in metadata.xml"
    );
    assert!(
        metadata.contains("xlrd:rvb i=\"1\""),
        "expected xlrd:rvb i=\"1\" in metadata.xml"
    );

    // richValueRel.xml.rels should resolve rel IDs to image binaries.
    let mut rich_value_rels = String::new();
    archive
        .by_name("xl/richData/_rels/richValueRel.xml.rels")
        .expect("xl/richData/_rels/richValueRel.xml.rels")
        .read_to_string(&mut rich_value_rels)
        .expect("read richValueRel.xml.rels");
    assert!(
        rich_value_rels.contains("../media/image1.png"),
        "expected image1.png target in richValueRel.xml.rels"
    );
    assert!(
        rich_value_rels.contains("../media/image2.png"),
        "expected image2.png target in richValueRel.xml.rels"
    );

    // `load_from_bytes` should opportunistically load RichData-backed in-cell images into
    // `workbook.images`, even when the workbook does not include `xl/cellimages.xml`.
    let doc = formula_xlsx::load_from_bytes(&bytes).expect("load xlsx document");
    let stored_image1 = doc
        .workbook
        .images
        .get(&ImageId::new("image1.png"))
        .expect("expected workbook.images to contain image1.png");
    assert_eq!(
        stored_image1.bytes, image1,
        "expected workbook.images image1.png bytes to match xl/media/image1.png"
    );
    let stored_image2 = doc
        .workbook
        .images
        .get(&ImageId::new("image2.png"))
        .expect("expected workbook.images to contain image2.png");
    assert_eq!(
        stored_image2.bytes, image2,
        "expected workbook.images image2.png bytes to match xl/media/image2.png"
    );

    // `vm` should resolve to rich value indices and be captured in `XlsxDocument.meta.rich_value_cells`.
    // In this fixture:
    // - vm="1" -> rich value index 0 (image1)
    // - vm="2" -> rich value index 1 (image2)
    let sheet_id = doc.workbook.sheets[0].id;
    assert_eq!(doc.rich_value_index(sheet_id, b2), Some(0));
    assert_eq!(doc.rich_value_index(sheet_id, b3), Some(0));
    assert_eq!(doc.rich_value_index(sheet_id, b4), Some(1));
}
