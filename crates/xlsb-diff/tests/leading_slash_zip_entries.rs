use std::io::{Cursor, Write};

use pretty_assertions::assert_eq;
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipWriter};

fn build_zip(parts: &[(&str, &[u8])]) -> Vec<u8> {
    let mut writer = ZipWriter::new(Cursor::new(Vec::<u8>::new()));
    let options = FileOptions::<()>::default().compression_method(CompressionMethod::Stored);

    for (name, bytes) in parts {
        writer.start_file(*name, options).expect("start file");
        writer.write_all(bytes).expect("write bytes");
    }

    writer.finish().expect("finish zip").into_inner()
}

#[test]
fn diff_treats_leading_slash_zip_entry_names_as_equivalent() {
    let payload = b"workbook-bytes";
    let expected_zip = build_zip(&[("xl/workbook.bin", payload)]);
    let actual_zip = build_zip(&[("/xl/workbook.bin", payload)]);

    let expected = xlsb_diff::WorkbookArchive::from_bytes(&expected_zip).expect("open expected");
    let actual = xlsb_diff::WorkbookArchive::from_bytes(&actual_zip).expect("open actual");

    assert_eq!(actual.get("xl/workbook.bin"), Some(payload.as_slice()));

    let report = xlsb_diff::diff_archives(&expected, &actual);
    assert!(
        report.is_empty(),
        "expected no diffs when entry names only differ by leading '/'; got:\n{}",
        report
            .differences
            .iter()
            .map(|d| d.to_string())
            .collect::<String>()
    );
}

