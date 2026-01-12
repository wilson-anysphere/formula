use std::collections::HashMap;
use std::fs;

use formula_xlsb::XlsbWorkbook;

#[test]
fn save_as_creates_parent_directories() {
    let path = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/simple.xlsb");
    let wb = XlsbWorkbook::open(path).expect("open xlsb fixture");

    let tmpdir = tempfile::tempdir().expect("tempdir");
    let dest = tmpdir.path().join("nested/dir/out.xlsb");

    wb.save_as(&dest).expect("save_as should create parent dirs");
    assert!(dest.is_file(), "expected {dest:?} to be created");
}

#[test]
fn save_with_part_overrides_does_not_clobber_existing_file_on_error() {
    let path = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/simple.xlsb");
    let wb = XlsbWorkbook::open(path).expect("open xlsb fixture");

    let tmpdir = tempfile::tempdir().expect("tempdir");
    let dest = tmpdir.path().join("out.xlsb");

    let sentinel = b"sentinel-xlsb";
    fs::write(&dest, sentinel).expect("write sentinel");

    // Use an override part name that does not exist in the source package. The writer will only
    // detect this after streaming the ZIP to the output, so a non-atomic implementation would
    // leave a partially-written workbook on disk.
    let overrides = HashMap::from([("xl/does-not-exist.bin".to_string(), vec![1, 2, 3])]);
    let err = wb
        .save_with_part_overrides(&dest, &overrides)
        .expect_err("expected override error");
    assert!(
        format!("{err}").contains("override parts not found"),
        "unexpected error: {err}"
    );

    assert_eq!(fs::read(&dest).expect("read dest"), sentinel);

    let mut entries: Vec<_> = fs::read_dir(tmpdir.path())
        .expect("read_dir")
        .map(|e| e.expect("dir entry").path())
        .collect();
    entries.sort();
    assert_eq!(entries, vec![dest], "expected no temp files to remain");
}

#[test]
fn save_with_part_overrides_streaming_does_not_clobber_existing_file_on_error() {
    let path = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/simple.xlsb");
    let wb = XlsbWorkbook::open(path).expect("open xlsb fixture");

    let tmpdir = tempfile::tempdir().expect("tempdir");
    let dest = tmpdir.path().join("out.xlsb");

    let sentinel = b"sentinel-streaming";
    fs::write(&dest, sentinel).expect("write sentinel");

    let overrides: HashMap<String, Vec<u8>> = HashMap::new();
    let err = wb
        .save_with_part_overrides_streaming(&dest, &overrides, "xl/does-not-exist.bin", |_i, _o| {
            Ok(false)
        })
        .expect_err("expected override error");
    assert!(
        format!("{err}").contains("override parts not found"),
        "unexpected error: {err}"
    );

    assert_eq!(fs::read(&dest).expect("read dest"), sentinel);

    let mut entries: Vec<_> = fs::read_dir(tmpdir.path())
        .expect("read_dir")
        .map(|e| e.expect("dir entry").path())
        .collect();
    entries.sort();
    assert_eq!(entries, vec![dest], "expected no temp files to remain");
}

