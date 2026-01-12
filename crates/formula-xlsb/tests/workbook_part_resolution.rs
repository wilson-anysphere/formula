use std::collections::BTreeMap;
use std::fs::File;
use std::io::{Cursor, Read, Write};
use std::path::Path;

use formula_xlsb::{CellValue, XlsbWorkbook};
use pretty_assertions::assert_eq;
use tempfile::tempdir;
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

mod fixture_builder;
use fixture_builder::XlsbFixtureBuilder;

const WORKBOOK_PART: &str = "xl/workbook1.bin";
const WORKBOOK_RELS_PART: &str = "xl/_rels/workbook1.bin.rels";

fn build_nonstandard_workbook_part_fixture(with_calc_chain: bool) -> Vec<u8> {
    let mut builder = XlsbFixtureBuilder::new();
    builder.add_shared_string("Hello");
    builder.set_cell_sst(0, 0, 0);
    let base = builder.build_bytes();

    let mut zip = ZipArchive::new(Cursor::new(base)).expect("open base xlsb zip");
    let mut parts: BTreeMap<String, Vec<u8>> = BTreeMap::new();
    for i in 0..zip.len() {
        let mut file = zip.by_index(i).expect("read zip entry");
        if file.is_dir() {
            continue;
        }
        let mut buf = Vec::new();
        file.read_to_end(&mut buf).expect("read part bytes");
        parts.insert(file.name().to_string(), buf);
    }

    let workbook_bin = parts
        .remove("xl/workbook.bin")
        .expect("base fixture has xl/workbook.bin");
    parts.insert(WORKBOOK_PART.to_string(), workbook_bin);

    let workbook_rels_bin = parts
        .remove("xl/_rels/workbook.bin.rels")
        .expect("base fixture has xl/_rels/workbook.bin.rels");
    let mut workbook_rels_xml =
        String::from_utf8(workbook_rels_bin).expect("utf8 workbook relationships");
    if with_calc_chain {
        workbook_rels_xml = workbook_rels_xml.replace(
            "</Relationships>",
            "  <Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain\" Target=\"calcChain.bin\"/>\n</Relationships>",
        );
        parts.insert("xl/calcChain.bin".to_string(), b"dummy".to_vec());
    }
    parts.insert(WORKBOOK_RELS_PART.to_string(), workbook_rels_xml.into_bytes());

    let root_rels = String::from_utf8(parts["_rels/.rels"].clone()).expect("utf8 root rels");
    let root_rels = root_rels.replace("Target=\"xl/workbook.bin\"", &format!("Target=\"{WORKBOOK_PART}\""));
    parts.insert("_rels/.rels".to_string(), root_rels.into_bytes());

    let content_types =
        String::from_utf8(parts["[Content_Types].xml"].clone()).expect("utf8 content types");
    let mut content_types = content_types.replace("/xl/workbook.bin", &format!("/{WORKBOOK_PART}"));
    if with_calc_chain {
        content_types = content_types.replace(
            "</Types>",
            "  <Override PartName=\"/xl/calcChain.bin\" ContentType=\"application/vnd.ms-excel.calcChain\"/>\n</Types>",
        );
    }
    parts.insert("[Content_Types].xml".to_string(), content_types.into_bytes());

    let mut zip_out = ZipWriter::new(Cursor::new(Vec::new()));
    let options = FileOptions::default().compression_method(CompressionMethod::Stored);
    for (name, bytes) in parts {
        zip_out.start_file(name, options).expect("start file");
        zip_out.write_all(&bytes).expect("write bytes");
    }

    zip_out.finish().expect("finish zip").into_inner()
}

fn zip_has_part(path: &Path, part: &str) -> bool {
    let file = File::open(path).expect("open zip");
    let zip = ZipArchive::new(file).expect("read zip");
    let has = zip.file_names().any(|name| name == part);
    has
}

fn zip_read_to_string(path: &Path, part: &str) -> String {
    let file = File::open(path).expect("open zip");
    let mut zip = ZipArchive::new(file).expect("read zip");
    let mut entry = zip.by_name(part).expect("find part");
    let mut bytes = Vec::new();
    entry.read_to_end(&mut bytes).expect("read part bytes");
    String::from_utf8(bytes).expect("utf8")
}

#[test]
fn open_resolves_workbook_part_from_root_rels() {
    let bytes = build_nonstandard_workbook_part_fixture(false);

    let tmpdir = tempdir().expect("create temp dir");
    let input_path = tmpdir.path().join("input.xlsb");
    std::fs::write(&input_path, bytes).expect("write input workbook");

    let wb = XlsbWorkbook::open(&input_path).expect("open workbook");
    let sheet = wb.read_sheet(0).expect("read sheet");
    let a1 = sheet
        .cells
        .iter()
        .find(|c| c.row == 0 && c.col == 0)
        .expect("A1 exists");
    assert_eq!(a1.value, CellValue::Text("Hello".to_string()));
}

#[test]
fn edited_save_updates_calc_chain_plumbing_for_nonstandard_workbook_rels_part() {
    let bytes = build_nonstandard_workbook_part_fixture(true);

    let tmpdir = tempdir().expect("create temp dir");
    let input_path = tmpdir.path().join("input.xlsb");
    let output_path = tmpdir.path().join("output.xlsb");
    std::fs::write(&input_path, bytes).expect("write input workbook");

    let wb = XlsbWorkbook::open(&input_path).expect("open workbook");
    wb.save_with_edits(&output_path, 0, 0, 1, 123.0)
        .expect("save_with_edits");

    assert!(
        zip_has_part(&output_path, WORKBOOK_PART),
        "expected nonstandard workbook part to be present in output"
    );
    assert!(
        zip_has_part(&output_path, WORKBOOK_RELS_PART),
        "expected nonstandard workbook relationships part to be present in output"
    );
    assert!(
        !zip_has_part(&output_path, "xl/workbook.bin"),
        "did not expect default workbook part to appear in output"
    );
    assert!(
        !zip_has_part(&output_path, "xl/_rels/workbook.bin.rels"),
        "did not expect default workbook relationships part to appear in output"
    );

    assert!(
        !zip_has_part(&output_path, "xl/calcChain.bin"),
        "expected calcChain part to be removed in edited save"
    );

    let content_types = zip_read_to_string(&output_path, "[Content_Types].xml");
    assert!(
        !content_types.to_ascii_lowercase().contains("calcchain"),
        "expected calcChain override to be removed from [Content_Types].xml, got:\n{content_types}"
    );

    let workbook_rels = zip_read_to_string(&output_path, WORKBOOK_RELS_PART);
    assert!(
        !workbook_rels.to_ascii_lowercase().contains("calcchain"),
        "expected calcChain relationship to be removed from workbook rels, got:\n{workbook_rels}"
    );

    let wb2 = XlsbWorkbook::open(&output_path).expect("open output workbook");
    let sheet = wb2.read_sheet(0).expect("read sheet");
    let b1 = sheet
        .cells
        .iter()
        .find(|c| c.row == 0 && c.col == 1)
        .expect("B1 exists");
    assert_eq!(b1.value, CellValue::Number(123.0));
}
