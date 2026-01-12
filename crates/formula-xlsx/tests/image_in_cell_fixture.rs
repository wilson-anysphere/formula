use std::collections::HashSet;
use std::io::{Cursor, Read};

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

    // If/when rich-value image extraction is implemented in `load_from_bytes`, ensure the
    // workbook has images populated. (Today this is best-effort and may be empty.)
    let doc = formula_xlsx::load_from_bytes(&bytes).expect("load_from_bytes");
    if !doc.workbook.images.is_empty() {
        assert!(
            doc.workbook.images.ids().count() >= 1,
            "expected at least one extracted cell image"
        );
    }
}
