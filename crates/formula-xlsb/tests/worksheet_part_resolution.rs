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

fn build_case_mismatched_sheet_part_fixture() -> Vec<u8> {
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

    // Rename the worksheet part to a different case, without updating workbook.bin.rels.
    let sheet_bytes = parts
        .remove("xl/worksheets/sheet1.bin")
        .expect("base fixture has sheet1.bin");
    parts.insert("xl/worksheets/Sheet1.bin".to_string(), sheet_bytes);

    // Keep the content types consistent with the renamed part.
    let content_types =
        String::from_utf8(parts["[Content_Types].xml"].clone()).expect("utf8 content types");
    let content_types =
        content_types.replace("/xl/worksheets/sheet1.bin", "/xl/worksheets/Sheet1.bin");
    parts.insert("[Content_Types].xml".to_string(), content_types.into_bytes());

    let mut zip_out = ZipWriter::new(Cursor::new(Vec::new()));
    let options = FileOptions::<()>::default().compression_method(CompressionMethod::Stored);
    for (name, bytes) in parts {
        zip_out
            .start_file(name, options.clone())
            .expect("start file");
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

#[test]
fn open_reads_sheet_when_zip_entry_case_differs_from_relationship_target() {
    let bytes = build_case_mismatched_sheet_part_fixture();

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
fn save_with_edits_overrides_the_actual_sheet_part_name() {
    let bytes = build_case_mismatched_sheet_part_fixture();

    let tmpdir = tempdir().expect("create temp dir");
    let input_path = tmpdir.path().join("input.xlsb");
    let output_path = tmpdir.path().join("output.xlsb");
    std::fs::write(&input_path, bytes).expect("write input workbook");

    let wb = XlsbWorkbook::open(&input_path).expect("open workbook");
    wb.save_with_edits(&output_path, 0, 0, 1, 123.0)
        .expect("save_with_edits");

    let wb2 = XlsbWorkbook::open(&output_path).expect("open output workbook");
    let sheet = wb2.read_sheet(0).expect("read sheet");
    let b1 = sheet
        .cells
        .iter()
        .find(|c| c.row == 0 && c.col == 1)
        .expect("B1 exists");
    assert_eq!(b1.value, CellValue::Number(123.0));

    assert!(
        zip_has_part(&output_path, "xl/worksheets/Sheet1.bin"),
        "expected renamed worksheet part to be present in output"
    );
    assert!(
        !zip_has_part(&output_path, "xl/worksheets/sheet1.bin"),
        "did not expect lowercase worksheet part to appear in output"
    );
}
