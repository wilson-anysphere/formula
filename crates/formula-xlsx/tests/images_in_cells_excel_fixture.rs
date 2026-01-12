use std::collections::{BTreeMap, HashMap, HashSet};
use std::io::{Cursor, Read};

use formula_model::drawings::ImageId;
use roxmltree::Document;
use zip::ZipArchive;

/// Ground-truth-ish `.xlsx` fixture meant to reflect Excel's real part graph for
/// "Images in Cells" (Place-in-Cell picture + `IMAGE()`).
///
/// See `docs/20-images-in-cells.md` for the documented part graph and the
/// expected content types / relationships.
const FIXTURE: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../..",
    "/fixtures/xlsx/images-in-cells/image-in-cell.xlsx"
));

fn resolve_target(source_part: &str, target: &str) -> String {
    if let Some(t) = target.strip_prefix('/') {
        return normalize_path(t);
    }

    let base_dir = source_part
        .rsplit_once('/')
        .map(|(dir, _)| dir)
        .unwrap_or("");
    normalize_path(&format!("{base_dir}/{target}"))
}

fn normalize_path(path: &str) -> String {
    let mut out: Vec<&str> = Vec::new();
    for part in path.split('/') {
        match part {
            "" | "." => {}
            ".." => {
                out.pop();
            }
            other => out.push(other),
        }
    }
    out.join("/")
}

fn open_fixture_zip() -> ZipArchive<Cursor<&'static [u8]>> {
    ZipArchive::new(Cursor::new(FIXTURE)).expect("fixture must be a valid zip")
}

fn zip_part(zip: &mut ZipArchive<Cursor<&'static [u8]>>, name: &str) -> Vec<u8> {
    let mut file = zip.by_name(name).unwrap_or_else(|_| panic!("missing part {name}"));
    let mut buf = Vec::new();
    file.read_to_end(&mut buf)
        .unwrap_or_else(|_| panic!("failed reading part {name}"));
    buf
}

fn zip_names(zip: &mut ZipArchive<Cursor<&'static [u8]>>) -> HashSet<String> {
    (0..zip.len())
        .map(|i| zip.by_index(i).expect("zip index").name().to_string())
        .collect()
}

fn parse_relationship_targets(rels_xml: &str) -> Vec<(String, String)> {
    // Returns (type_uri, target)
    let doc = Document::parse(rels_xml).expect("rels XML must parse");
    doc.descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "Relationship")
        .filter_map(|n| {
            let ty = n.attribute("Type")?.to_string();
            let target = n.attribute("Target")?.to_string();
            Some((ty, target))
        })
        .collect()
}

fn parse_content_types_overrides(ct_xml: &str) -> HashMap<String, String> {
    let doc = Document::parse(ct_xml).expect("content types XML must parse");
    doc.descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "Override")
        .filter_map(|n| {
            let part_name = n.attribute("PartName")?.to_string();
            let content_type = n.attribute("ContentType")?.to_string();
            Some((part_name, content_type))
        })
        .collect()
}

#[test]
fn excel_images_in_cells_fixture_has_expected_part_graph() {
    let mut zip = open_fixture_zip();
    let names = zip_names(&mut zip);

    // Core expected parts for "images in cells".
    assert!(names.contains("xl/cellimages.xml"));
    assert!(names.contains("xl/_rels/cellimages.xml.rels"));
    assert!(names.contains("xl/metadata.xml"));

    // RichData expected part set (Excel version-dependent, but these are expected in this fixture).
    for part in [
        "xl/richData/richValue.xml",
        "xl/richData/richValueRel.xml",
        "xl/richData/richValueTypes.xml",
        "xl/richData/richValueStructure.xml",
        "xl/richData/_rels/richValueRel.xml.rels",
    ] {
        assert!(
            names.contains(part),
            "expected fixture to contain richData part {part}"
        );
    }

    // Verify sheet1.xml has `vm` and/or `cm` metadata attributes on at least one cell.
    let sheet1_xml = zip_part(&mut zip, "xl/worksheets/sheet1.xml");
    let sheet1_xml = std::str::from_utf8(&sheet1_xml).expect("sheet1.xml must be utf-8");
    let sheet_doc = Document::parse(sheet1_xml).expect("sheet1.xml must parse");
    let vm_values: Vec<u32> = sheet_doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "c")
        .filter_map(|cell| cell.attribute("vm").and_then(|v| v.parse::<u32>().ok()))
        .collect();
    let cm_values: Vec<u32> = sheet_doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "c")
        .filter_map(|cell| cell.attribute("cm").and_then(|v| v.parse::<u32>().ok()))
        .collect();
    assert!(
        !vm_values.is_empty() || !cm_values.is_empty(),
        "expected at least one <c> with vm/cn metadata index (vm={vm_values:?}, cm={cm_values:?})"
    );

    // Verify `cellimages.xml.rels` image targets exist as parts.
    let cellimages_rels = zip_part(&mut zip, "xl/_rels/cellimages.xml.rels");
    let cellimages_rels = std::str::from_utf8(&cellimages_rels).expect("cellimages.xml.rels utf-8");
    for (ty, target) in parse_relationship_targets(cellimages_rels) {
        if ty == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" {
            let resolved = resolve_target("xl/cellimages.xml", &target);
            assert!(
                names.contains(resolved.as_str()),
                "cellimages.xml.rels references missing target part {resolved} (Target={target})"
            );
        }
    }

    // Verify `richValueRel.xml.rels` image targets exist as parts.
    let rv_rels = zip_part(&mut zip, "xl/richData/_rels/richValueRel.xml.rels");
    let rv_rels = std::str::from_utf8(&rv_rels).expect("richValueRel.xml.rels utf-8");
    for (ty, target) in parse_relationship_targets(rv_rels) {
        if ty == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" {
            let resolved = resolve_target("xl/richData/richValueRel.xml", &target);
            assert!(
                names.contains(resolved.as_str()),
                "richValueRel.xml.rels references missing target part {resolved} (Target={target})"
            );
        }
    }
}

#[test]
fn excel_images_in_cells_fixture_loads_images_and_metadata_mappings() {
    // Ensure our high-level loader succeeds on a file containing real-ish images-in-cells parts.
    let doc = formula_xlsx::load_from_bytes(FIXTURE).expect("load_from_bytes must succeed");

    assert!(
        !doc.workbook.images.is_empty(),
        "expected Workbook.images to contain at least one image loaded from xl/cellimages.xml"
    );
    assert!(
        doc.workbook.images.get(&ImageId::new("image1.png")).is_some(),
        "expected Workbook.images to contain xl/media/image1.png"
    );

    let mut zip = open_fixture_zip();
    let sheet1_xml = zip_part(&mut zip, "xl/worksheets/sheet1.xml");
    let sheet1_xml = std::str::from_utf8(&sheet1_xml).expect("sheet1.xml must be utf-8");
    let sheet_doc = Document::parse(sheet1_xml).expect("sheet1.xml must parse");
    let vm_values: Vec<u32> = sheet_doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "c")
        .filter_map(|cell| cell.attribute("vm").and_then(|v| v.parse::<u32>().ok()))
        .collect();

    if !vm_values.is_empty() {
        let metadata_xml = zip_part(&mut zip, "xl/metadata.xml");
        let map = formula_xlsx::parse_value_metadata_vm_to_rich_value_index_map(&metadata_xml)
            .expect("metadata.xml must parse");
        assert!(
            !map.is_empty(),
            "expected metadata.xml vm->richValue map to be non-empty"
        );
        for vm in vm_values {
            assert!(
                map.contains_key(&vm),
                "expected vm value {vm} from sheet1.xml to exist in metadata.xml map (map keys: {:?})",
                map.keys().collect::<Vec<_>>()
            );
        }
    }
}

#[test]
fn excel_images_in_cells_fixture_documents_content_types_and_workbook_relationships() {
    let mut zip = open_fixture_zip();

    let content_types_xml = zip_part(&mut zip, "[Content_Types].xml");
    let content_types_xml =
        std::str::from_utf8(&content_types_xml).expect("[Content_Types].xml utf-8");
    let overrides = parse_content_types_overrides(content_types_xml);

    assert_eq!(
        overrides.get("/xl/cellimages.xml").map(String::as_str),
        Some("application/vnd.ms-excel.cellimages+xml")
    );
    assert_eq!(
        overrides.get("/xl/metadata.xml").map(String::as_str),
        Some("application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml")
    );
    assert_eq!(
        overrides
            .get("/xl/richData/richValue.xml")
            .map(String::as_str),
        Some("application/vnd.ms-excel.richvalue+xml")
    );
    assert_eq!(
        overrides
            .get("/xl/richData/richValueRel.xml")
            .map(String::as_str),
        Some("application/vnd.ms-excel.richvaluerel+xml")
    );
    assert_eq!(
        overrides
            .get("/xl/richData/richValueTypes.xml")
            .map(String::as_str),
        Some("application/vnd.ms-excel.richvaluetypes+xml")
    );
    assert_eq!(
        overrides
            .get("/xl/richData/richValueStructure.xml")
            .map(String::as_str),
        Some("application/vnd.ms-excel.richvaluestructure+xml")
    );

    let workbook_rels = zip_part(&mut zip, "xl/_rels/workbook.xml.rels");
    let workbook_rels = std::str::from_utf8(&workbook_rels).expect("workbook.xml.rels utf-8");
    let rels: BTreeMap<(String, String), String> = parse_relationship_targets(workbook_rels)
        .into_iter()
        .map(|(ty, target)| ((ty, target.clone()), target))
        .collect();

    assert!(
        rels.contains_key(&(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata"
                .to_string(),
            "metadata.xml".to_string()
        )),
        "expected workbook.xml.rels to contain workbook->metadata.xml relationship, got:\n{workbook_rels}"
    );
    assert!(
        rels.contains_key(&(
            "http://schemas.microsoft.com/office/2017/06/relationships/cellImages".to_string(),
            "cellimages.xml".to_string()
        )),
        "expected workbook.xml.rels to contain workbook->cellimages.xml relationship, got:\n{workbook_rels}"
    );
    assert!(
        rels.contains_key(&(
            "http://schemas.microsoft.com/office/2017/06/relationships/richValue".to_string(),
            "richData/richValue.xml".to_string()
        )),
        "expected workbook.xml.rels to contain workbook->richData/richValue.xml relationship, got:\n{workbook_rels}"
    );
    assert!(
        rels.contains_key(&(
            "http://schemas.microsoft.com/office/2017/06/relationships/richValueRel".to_string(),
            "richData/richValueRel.xml".to_string()
        )),
        "expected workbook.xml.rels to contain workbook->richData/richValueRel.xml relationship, got:\n{workbook_rels}"
    );
    assert!(
        rels.contains_key(&(
            "http://schemas.microsoft.com/office/2017/06/relationships/richValueTypes".to_string(),
            "richData/richValueTypes.xml".to_string()
        )),
        "expected workbook.xml.rels to contain workbook->richData/richValueTypes.xml relationship, got:\n{workbook_rels}"
    );
    assert!(
        rels.contains_key(&(
            "http://schemas.microsoft.com/office/2017/06/relationships/richValueStructure".to_string(),
            "richData/richValueStructure.xml".to_string()
        )),
        "expected workbook.xml.rels to contain workbook->richData/richValueStructure.xml relationship, got:\n{workbook_rels}"
    );
}
