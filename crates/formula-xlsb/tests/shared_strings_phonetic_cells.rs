use formula_xlsb::{CellValue, OpenOptions, XlsbWorkbook};
use pretty_assertions::assert_eq;
use tempfile::tempdir;

mod fixture_builder;
use fixture_builder::XlsbFixtureBuilder;

#[test]
fn preserves_shared_string_phonetic_bytes_on_string_cells() {
    let mut builder = XlsbFixtureBuilder::new();

    let phonetic_bytes = vec![0xDE, 0xAD, 0xBE, 0xEF];
    let sst_idx = builder.add_shared_string_with_phonetic("Hi", phonetic_bytes.clone());
    builder.set_cell_sst(0, 0, sst_idx);

    let bytes = builder.build_bytes();

    let tmpdir = tempdir().expect("create temp dir");
    let path = tmpdir.path().join("phonetic.xlsb");
    std::fs::write(&path, bytes).expect("write xlsb bytes");

    let wb = XlsbWorkbook::open_with_options(
        &path,
        OpenOptions {
            preserve_parsed_parts: true,
            ..OpenOptions::default()
        },
    )
    .expect("open xlsb");

    let table = wb.shared_strings_table();
    assert_eq!(table.len(), 1);
    assert_eq!(
        table[sst_idx as usize].phonetic.as_deref(),
        Some(phonetic_bytes.as_slice())
    );

    let sheet = wb.read_sheet(0).expect("read sheet");
    let a1 = sheet
        .cells
        .iter()
        .find(|c| c.row == 0 && c.col == 0)
        .expect("A1 exists");
    assert_eq!(a1.value, CellValue::Text("Hi".to_string()));

    let preserved = a1.preserved_string.as_ref().expect("preserved string");
    assert_eq!(preserved.text, "Hi");
    assert_eq!(
        preserved.phonetic.as_deref(),
        Some(phonetic_bytes.as_slice())
    );
    assert_eq!(preserved.rich, None);
}

#[test]
fn preserves_shared_string_rich_runs_on_string_cells() {
    let mut builder = XlsbFixtureBuilder::new();

    // One StrRun (MS-XLSB): [ich: u32][ifnt: u16][reserved: u16]
    // Use a non-zero `ifnt`/reserved payload so we can assert the opaque bytes are preserved.
    let mut run_bytes = Vec::new();
    run_bytes.extend_from_slice(&0u32.to_le_bytes()); // ich
    run_bytes.extend_from_slice(&0x1234u16.to_le_bytes()); // ifnt
    run_bytes.extend_from_slice(&0x5678u16.to_le_bytes()); // reserved

    let sst_idx = builder.add_shared_string_with_rich_runs("Hi", run_bytes.clone());
    builder.set_cell_sst(0, 0, sst_idx);

    let bytes = builder.build_bytes();

    let tmpdir = tempdir().expect("create temp dir");
    let path = tmpdir.path().join("rich.xlsb");
    std::fs::write(&path, bytes).expect("write xlsb bytes");

    let wb = XlsbWorkbook::open_with_options(
        &path,
        OpenOptions {
            preserve_parsed_parts: true,
            ..OpenOptions::default()
        },
    )
    .expect("open xlsb");

    let table = wb.shared_strings_table();
    assert_eq!(table.len(), 1);
    assert!(!table[sst_idx as usize].rich_text.is_plain());
    assert_eq!(table[sst_idx as usize].phonetic, None);

    let sheet = wb.read_sheet(0).expect("read sheet");
    let a1 = sheet
        .cells
        .iter()
        .find(|c| c.row == 0 && c.col == 0)
        .expect("A1 exists");
    assert_eq!(a1.value, CellValue::Text("Hi".to_string()));

    let preserved = a1.preserved_string.as_ref().expect("preserved string");
    assert_eq!(preserved.text, "Hi");
    assert_eq!(preserved.phonetic, None);

    let rich = preserved.rich.as_ref().expect("preserved rich runs");
    assert_eq!(rich.runs, run_bytes);
}
