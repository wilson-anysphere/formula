use std::collections::BTreeMap;
use std::io::{Cursor, Read};
use std::path::Path;

use formula_model::drawings::ImageId;
use zip::ZipArchive;

const REQUIRED_PARTS: &[&str] = &[
    "xl/cellimages.xml",
    "xl/_rels/cellimages.xml.rels",
    "xl/media/image1.png",
];

const XML_PARTS: &[&str] = &["xl/cellimages.xml", "xl/_rels/cellimages.xml.rels"];

fn read_zip_parts(
    bytes: &[u8],
    part_names: &[&str],
) -> Result<BTreeMap<String, Vec<u8>>, zip::result::ZipError> {
    let mut archive = ZipArchive::new(Cursor::new(bytes))?;
    let mut out = BTreeMap::new();
    for name in part_names {
        let mut f = archive.by_name(name)?;
        // Do not trust `ZipFile::size()` for allocation; ZIP metadata is untrusted and can
        // advertise enormous uncompressed sizes (zip-bomb style OOM).
        let mut buf = Vec::new();
        f.read_to_end(&mut buf)?;
        out.insert((*name).to_string(), buf);
    }
    Ok(out)
}

fn assert_xml_bytes_equal_or_semantic(part_name: &str, expected: &[u8], actual: &[u8]) {
    if expected == actual {
        return;
    }

    let expected_norm = formula_xlsx::normalize_xml(expected).expect("normalize expected xml");
    let actual_norm = formula_xlsx::normalize_xml(actual).expect("normalize actual xml");
    assert_eq!(
        expected_norm, actual_norm,
        "XML part changed after round-trip: {part_name}"
    );
}

#[test]
fn roundtrip_preserves_cellimages_parts_for_fixture() -> Result<(), Box<dyn std::error::Error>> {
    let fixture =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx/basic/cellimages.xlsx");

    let original_bytes = std::fs::read(&fixture)?;

    let doc = formula_xlsx::load_from_path(&fixture)?;
    let out_bytes = doc.save_to_vec()?;

    let original_parts = read_zip_parts(&original_bytes, REQUIRED_PARTS)?;
    let out_parts = read_zip_parts(&out_bytes, REQUIRED_PARTS)?;

    // Ensure the in-cell image store is also surfaced through `Workbook.images`.
    let image = doc
        .workbook
        .images
        .get(&ImageId::new("image1.png"))
        .expect("expected Workbook.images to contain image1.png from xl/cellimages.xml");
    assert_eq!(
        image.bytes,
        *original_parts
            .get("xl/media/image1.png")
            .expect("missing xl/media/image1.png in fixture zip"),
        "expected xl/cellimages.xml image relationship to load image bytes into Workbook.images"
    );

    for part_name in REQUIRED_PARTS {
        assert!(
            out_parts.contains_key(*part_name),
            "missing expected part in output zip: {part_name}"
        );
    }

    for part_name in XML_PARTS {
        let expected = original_parts
            .get(*part_name)
            .unwrap_or_else(|| panic!("missing part in fixture zip: {part_name}"));
        let actual = out_parts
            .get(*part_name)
            .unwrap_or_else(|| panic!("missing part in output zip: {part_name}"));
        assert_xml_bytes_equal_or_semantic(part_name, expected, actual);
    }

    assert_eq!(
        out_parts
            .get("xl/media/image1.png")
            .expect("missing xl/media/image1.png in output zip"),
        original_parts
            .get("xl/media/image1.png")
            .expect("missing xl/media/image1.png in fixture zip"),
        "expected image payload to be preserved byte-for-byte"
    );

    Ok(())
}
