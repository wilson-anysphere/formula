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
fn preserves_shared_string_phonetic_bytes_without_preserve_parsed_parts() {
    let mut builder = XlsbFixtureBuilder::new();

    let phonetic_bytes = vec![0xDE, 0xAD, 0xBE, 0xEF];
    let sst_idx = builder.add_shared_string_with_phonetic("Hi", phonetic_bytes.clone());
    builder.set_cell_sst(0, 0, sst_idx);

    let bytes = builder.build_bytes();

    let tmpdir = tempdir().expect("create temp dir");
    let path = tmpdir.path().join("phonetic_no_preserve.xlsb");
    std::fs::write(&path, bytes).expect("write xlsb bytes");

    // `formula-io` opens XLSB files with `preserve_parsed_parts=false`. Ensure the phonetic tail is
    // still surfaced on the parsed cell so downstream importers can extract it.
    let wb = XlsbWorkbook::open_with_options(
        &path,
        OpenOptions {
            preserve_parsed_parts: false,
            ..OpenOptions::default()
        },
    )
    .expect("open xlsb");

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
fn for_each_cell_preserves_shared_string_phonetic_bytes() {
    let mut builder = XlsbFixtureBuilder::new();

    let phonetic_bytes = vec![0xDE, 0xAD, 0xBE, 0xEF];
    let sst_idx = builder.add_shared_string_with_phonetic("Hi", phonetic_bytes.clone());
    builder.set_cell_sst(0, 0, sst_idx);

    let bytes = builder.build_bytes();

    let tmpdir = tempdir().expect("create temp dir");
    let path = tmpdir.path().join("phonetic_for_each_cell.xlsb");
    std::fs::write(&path, bytes).expect("write xlsb bytes");

    // Exercise both worksheet-reading paths:
    // - preserve_worksheets=false: stream from the ZIP entry
    // - preserve_worksheets=true: parse from preserved worksheet bytes
    for preserve_worksheets in [false, true] {
        let wb = XlsbWorkbook::open_with_options(
            &path,
            OpenOptions {
                preserve_parsed_parts: false,
                preserve_worksheets,
                ..OpenOptions::default()
            },
        )
        .expect("open xlsb");

        let mut cells = Vec::new();
        wb.for_each_cell(0, |cell| cells.push(cell))
            .expect("for_each_cell");
        assert_eq!(cells.len(), 1);

        let a1 = &cells[0];
        assert_eq!(a1.value, CellValue::Text("Hi".to_string()));

        let preserved = a1.preserved_string.as_ref().expect("preserved string");
        assert_eq!(preserved.text, "Hi");
        assert_eq!(
            preserved.phonetic.as_deref(),
            Some(phonetic_bytes.as_slice())
        );
    }
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

#[test]
fn does_not_preserve_shared_string_rich_runs_without_preserve_parsed_parts() {
    let mut builder = XlsbFixtureBuilder::new();

    // One StrRun (MS-XLSB): [ich: u32][ifnt: u16][reserved: u16]
    let mut run_bytes = Vec::new();
    run_bytes.extend_from_slice(&0u32.to_le_bytes()); // ich
    run_bytes.extend_from_slice(&0x1234u16.to_le_bytes()); // ifnt
    run_bytes.extend_from_slice(&0x5678u16.to_le_bytes()); // reserved

    let sst_idx = builder.add_shared_string_with_rich_runs("Hi", run_bytes);
    builder.set_cell_sst(0, 0, sst_idx);

    let bytes = builder.build_bytes();

    let tmpdir = tempdir().expect("create temp dir");
    let path = tmpdir.path().join("rich_no_preserve.xlsb");
    std::fs::write(&path, bytes).expect("write xlsb bytes");

    // Rich text runs can be large; they are only preserved when `preserve_parsed_parts=true`.
    let wb = XlsbWorkbook::open_with_options(
        &path,
        OpenOptions {
            preserve_parsed_parts: false,
            ..OpenOptions::default()
        },
    )
    .expect("open xlsb");

    let sheet = wb.read_sheet(0).expect("read sheet");
    let a1 = sheet
        .cells
        .iter()
        .find(|c| c.row == 0 && c.col == 0)
        .expect("A1 exists");
    assert_eq!(a1.value, CellValue::Text("Hi".to_string()));
    assert_eq!(a1.preserved_string, None);
}

#[test]
fn preserves_shared_string_rich_run_offsets_with_surrogate_pairs() {
    let mut builder = XlsbFixtureBuilder::new();

    // Text containing a surrogate pair (😀 is 2 UTF-16 code units).
    let text = "A😀B";

    // Two StrRun entries (MS-XLSB): [ich: u32][ifnt: u16][reserved: u16]
    // Run 0 starts at ich=0; run 1 starts at ich=3 (A=1 code unit, 😀=2).
    let mut run_bytes = Vec::new();

    run_bytes.extend_from_slice(&0u32.to_le_bytes()); // ich = 0
    run_bytes.extend_from_slice(&0x0102u16.to_le_bytes()); // ifnt
    run_bytes.extend_from_slice(&0x0304u16.to_le_bytes()); // reserved

    run_bytes.extend_from_slice(&3u32.to_le_bytes()); // ich = 3 (after A😀)
    run_bytes.extend_from_slice(&0x0506u16.to_le_bytes()); // ifnt
    run_bytes.extend_from_slice(&0x0708u16.to_le_bytes()); // reserved

    let sst_idx = builder.add_shared_string_with_rich_runs(text, run_bytes.clone());
    builder.set_cell_sst(0, 0, sst_idx);

    let bytes = builder.build_bytes();

    let tmpdir = tempdir().expect("create temp dir");
    let path = tmpdir.path().join("rich-surrogate.xlsb");
    std::fs::write(&path, bytes).expect("write xlsb bytes");

    let wb = XlsbWorkbook::open_with_options(
        &path,
        OpenOptions {
            preserve_parsed_parts: true,
            ..OpenOptions::default()
        },
    )
    .expect("open xlsb");

    let sheet = wb.read_sheet(0).expect("read sheet");
    let a1 = sheet
        .cells
        .iter()
        .find(|c| c.row == 0 && c.col == 0)
        .expect("A1 exists");
    assert_eq!(a1.value, CellValue::Text(text.to_string()));

    let preserved = a1.preserved_string.as_ref().expect("preserved string");
    let rich = preserved.rich.as_ref().expect("preserved rich runs");

    // The preserved run bytes should use UTF-16 offsets (not Rust `char` indices), so the second
    // run starts at 3.
    assert_eq!(rich.runs, run_bytes);
}

#[test]
fn preserves_shared_string_rich_and_phonetic_metadata_on_string_cells() {
    let mut builder = XlsbFixtureBuilder::new();

    let phonetic_bytes = vec![0xAA, 0xBB, 0xCC];

    // One StrRun (MS-XLSB): [ich: u32][ifnt: u16][reserved: u16]
    let mut run_bytes = Vec::new();
    run_bytes.extend_from_slice(&0u32.to_le_bytes()); // ich
    run_bytes.extend_from_slice(&0x1111u16.to_le_bytes()); // ifnt
    run_bytes.extend_from_slice(&0x2222u16.to_le_bytes()); // reserved

    let sst_idx = builder.add_shared_string_with_rich_runs_and_phonetic(
        "Hi",
        run_bytes.clone(),
        phonetic_bytes.clone(),
    );
    builder.set_cell_sst(0, 0, sst_idx);

    let bytes = builder.build_bytes();

    let tmpdir = tempdir().expect("create temp dir");
    let path = tmpdir.path().join("rich-phonetic.xlsb");
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

    let rich = preserved.rich.as_ref().expect("preserved rich runs");
    assert_eq!(rich.runs, run_bytes);
}
