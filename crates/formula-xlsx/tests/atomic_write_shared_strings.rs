use std::fs;
use std::io::{Cursor, Write};

use formula_xlsx::shared_strings::{write_shared_strings_to_xlsx, SharedStrings};
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipWriter};

fn build_zip_bytes(entries: &[(&str, &[u8])]) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(CompressionMethod::Stored);

    for (name, bytes) in entries {
        zip.start_file(*name, options).expect("start_file");
        zip.write_all(bytes).expect("write entry bytes");
    }

    zip.finish().expect("finish").into_inner()
}

#[test]
fn write_shared_strings_creates_parent_directories() {
    let dir = tempfile::tempdir().expect("temp dir");
    let input_path = dir.path().join("input.xlsx");
    fs::write(&input_path, build_zip_bytes(&[("good.txt", b"ok")])).expect("write input zip");

    let out_path = dir.path().join("nested/dir/out.xlsx");
    assert!(
        !out_path.parent().unwrap().exists(),
        "test precondition: parent dir should not exist"
    );

    write_shared_strings_to_xlsx(&input_path, &out_path, &SharedStrings::default())
        .expect("write_shared_strings_to_xlsx");

    assert!(out_path.is_file(), "expected {out_path:?} to be created");
}

#[test]
fn write_shared_strings_does_not_clobber_existing_file_on_error() {
    let dir = tempfile::tempdir().expect("temp dir");
    let input_path = dir.path().join("input.xlsx");
    let output_path = dir.path().join("out.xlsx");

    let sentinel = b"sentinel-shared-strings";
    fs::write(&output_path, sentinel).expect("seed existing output");

    // Build a zip with two stored entries, then corrupt the payload bytes of the second entry so
    // the ZIP CRC check fails during `read_to_end`. The implementation creates the output file
    // before reading entries, so a non-atomic implementation would truncate the existing output.
    let bad_payload = b"unique-bad-payload-0123456789abcdef";
    let mut bytes = build_zip_bytes(&[
        ("good.txt", b"good"),
        ("bad.txt", bad_payload.as_slice()),
    ]);

    let hits: Vec<usize> = bytes
        .windows(bad_payload.len())
        .enumerate()
        .filter_map(|(i, window)| (window == bad_payload).then_some(i))
        .collect();
    assert_eq!(
        hits.len(),
        1,
        "expected bad payload to appear exactly once in zip bytes"
    );
    let idx = hits[0];
    bytes[idx] ^= 0xFF;

    fs::write(&input_path, bytes).expect("write corrupted input zip");

    let _err = write_shared_strings_to_xlsx(&input_path, &output_path, &SharedStrings::default())
        .expect_err("expected write_shared_strings_to_xlsx to fail");

    assert_eq!(fs::read(&output_path).expect("read output"), sentinel);

    let mut entries: Vec<_> = fs::read_dir(dir.path())
        .expect("read_dir")
        .map(|e| e.expect("dir entry").path())
        .collect();
    entries.sort();
    assert_eq!(
        entries,
        vec![input_path, output_path],
        "expected no temp files to remain"
    );
}

