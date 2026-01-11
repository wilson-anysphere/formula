use std::collections::BTreeMap;
use std::fs::File;
use std::io::{Cursor, Read, Write};
use std::path::Path;

use formula_xlsb::{biff12_varint, CellEdit, CellValue, XlsbWorkbook};
use pretty_assertions::assert_eq;
use tempfile::tempdir;
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

mod fixture_builder;
use fixture_builder::XlsbFixtureBuilder;

fn build_nonstandard_shared_strings_fixture() -> Vec<u8> {
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

    let shared_strings = parts
        .remove("xl/sharedStrings.bin")
        .expect("base fixture has sharedStrings.bin");
    parts.insert("xl/strings/sharedStrings.bin".to_string(), shared_strings);

    let workbook_rels_xml =
        String::from_utf8(parts["xl/_rels/workbook.bin.rels"].clone()).expect("utf8 workbook rels");
    let workbook_rels_xml = workbook_rels_xml.replace(
        "Target=\"sharedStrings.bin\"",
        "Target=\"strings/sharedStrings.bin\"",
    );
    parts.insert(
        "xl/_rels/workbook.bin.rels".to_string(),
        workbook_rels_xml.into_bytes(),
    );

    let content_types_xml =
        String::from_utf8(parts["[Content_Types].xml"].clone()).expect("utf8 content types");
    let content_types_xml = content_types_xml.replace(
        "PartName=\"/xl/sharedStrings.bin\"",
        "PartName=\"/xl/strings/sharedStrings.bin\"",
    );
    parts.insert(
        "[Content_Types].xml".to_string(),
        content_types_xml.into_bytes(),
    );

    let mut zip_out = ZipWriter::new(Cursor::new(Vec::new()));
    let options = FileOptions::default().compression_method(CompressionMethod::Stored);
    for (name, bytes) in parts {
        zip_out.start_file(name, options).expect("start file");
        zip_out.write_all(&bytes).expect("write bytes");
    }

    zip_out.finish().expect("finish zip").into_inner()
}

fn read_zip_part(path: &Path, part_path: &str) -> Vec<u8> {
    let file = File::open(path).expect("open xlsb");
    let mut zip = ZipArchive::new(file).expect("read zip");
    let mut entry = zip.by_name(part_path).expect("find part");
    let mut bytes = Vec::with_capacity(entry.size() as usize);
    entry.read_to_end(&mut bytes).expect("read part bytes");
    bytes
}

struct SharedStringsInfo {
    total_count: Option<u32>,
    unique_count: Option<u32>,
    strings: Vec<String>,
}

fn read_shared_strings_info(shared_strings_bin: &[u8]) -> SharedStringsInfo {
    const SST: u32 = 0x019F;
    const SI: u32 = 0x0013;
    const SST_END: u32 = 0x01A0;

    let mut cursor = Cursor::new(shared_strings_bin);
    let mut total_count = None;
    let mut unique_count = None;
    let mut strings = Vec::new();

    loop {
        let id = match biff12_varint::read_record_id(&mut cursor).ok().flatten() {
            Some(id) => id,
            None => break,
        };
        let len = match biff12_varint::read_record_len(&mut cursor).ok().flatten() {
            Some(len) => len as usize,
            None => break,
        };
        let mut payload = vec![0u8; len];
        cursor
            .read_exact(&mut payload)
            .expect("read record payload");

        match id {
            SST if payload.len() >= 8 => {
                total_count = Some(u32::from_le_bytes(payload[0..4].try_into().unwrap()));
                unique_count = Some(u32::from_le_bytes(payload[4..8].try_into().unwrap()));
            }
            SI if payload.len() >= 5 => {
                let flags = payload[0];
                let cch = u32::from_le_bytes(payload[1..5].try_into().unwrap()) as usize;
                let byte_len = cch.saturating_mul(2);
                let raw = payload.get(5..5 + byte_len).unwrap_or(&[]);
                let mut units = Vec::with_capacity(cch);
                for chunk in raw.chunks_exact(2) {
                    units.push(u16::from_le_bytes([chunk[0], chunk[1]]));
                }
                let text = String::from_utf16_lossy(&units);
                if flags == 0 {
                    strings.push(text);
                }
            }
            SST_END => break,
            _ => {}
        }
    }

    SharedStringsInfo {
        total_count,
        unique_count,
        strings,
    }
}

#[test]
fn open_resolves_shared_strings_part_from_workbook_rels() {
    let bytes = build_nonstandard_shared_strings_fixture();

    let tmpdir = tempdir().expect("create temp dir");
    let input_path = tmpdir.path().join("input.xlsb");
    std::fs::write(&input_path, bytes).expect("write input workbook");

    let wb = XlsbWorkbook::open(&input_path).expect("open xlsb");
    let sheet = wb.read_sheet(0).expect("read sheet");
    let a1 = sheet
        .cells
        .iter()
        .find(|c| c.row == 0 && c.col == 0)
        .expect("A1 exists");
    assert_eq!(a1.value, CellValue::Text("Hello".to_string()));
}

#[test]
fn save_as_roundtrips_nonstandard_shared_strings_part_losslessly() {
    let bytes = build_nonstandard_shared_strings_fixture();

    let tmpdir = tempdir().expect("create temp dir");
    let input_path = tmpdir.path().join("input.xlsb");
    let output_path = tmpdir.path().join("output.xlsb");
    std::fs::write(&input_path, bytes).expect("write input workbook");

    let wb = XlsbWorkbook::open(&input_path).expect("open xlsb");
    wb.save_as(&output_path).expect("save_as");

    let report = xlsx_diff::diff_workbooks(&input_path, &output_path).expect("diff workbooks");
    assert!(
        report.is_empty(),
        "expected no OPC part diffs, got {} diffs",
        report.differences.len()
    );
}

#[test]
fn save_with_cell_edits_shared_strings_updates_nonstandard_shared_strings_part() {
    let bytes = build_nonstandard_shared_strings_fixture();

    let tmpdir = tempdir().expect("create temp dir");
    let input_path = tmpdir.path().join("input.xlsb");
    let output_path = tmpdir.path().join("output.xlsb");
    std::fs::write(&input_path, bytes).expect("write input workbook");

    let wb = XlsbWorkbook::open(&input_path).expect("open xlsb");
    wb.save_with_cell_edits_shared_strings(
        &output_path,
        0,
        &[CellEdit {
            row: 0,
            col: 0,
            new_value: CellValue::Text("New".to_string()),
            new_formula: None,
            new_rgcb: None,
            shared_string_index: None,
        }],
    )
    .expect("save_with_cell_edits_shared_strings");

    let wb2 = XlsbWorkbook::open(&output_path).expect("open output workbook");
    let sheet = wb2.read_sheet(0).expect("read sheet");
    let a1 = sheet
        .cells
        .iter()
        .find(|c| c.row == 0 && c.col == 0)
        .expect("A1 exists");
    assert_eq!(a1.value, CellValue::Text("New".to_string()));

    let shared_strings_bin = read_zip_part(&output_path, "xl/strings/sharedStrings.bin");
    let info = read_shared_strings_info(&shared_strings_bin);
    assert_eq!(info.total_count, Some(1));
    assert_eq!(info.unique_count, Some(2));
    assert_eq!(info.strings.len(), 2);
    assert_eq!(info.strings[1], "New");
}
