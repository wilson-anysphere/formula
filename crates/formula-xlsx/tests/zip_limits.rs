use std::io::Cursor;
use std::io::Write;

use formula_xlsx::{read_part_from_reader_limited, XlsxError, XlsxPackage, XlsxPackageLimits};
use zip::write::FileOptions;
use zip::ZipWriter;

fn build_zip_bytes(entries: &[(&str, &[u8])]) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    for (name, bytes) in entries {
        zip.start_file(*name, options).unwrap();
        zip.write_all(bytes).unwrap();
    }

    zip.finish().unwrap().into_inner()
}

#[test]
fn read_part_from_reader_limited_rejects_oversized_parts() {
    let bytes = build_zip_bytes(&[("xl/vbaProject.bin", b"0123456789A")]); // 11 bytes

    let err = read_part_from_reader_limited(Cursor::new(bytes), "xl/vbaProject.bin", 10)
        .expect_err("expected part-too-large error");

    match err {
        XlsxError::PartTooLarge { part, size, max } => {
            assert_eq!(part, "xl/vbaProject.bin");
            assert_eq!(size, 11);
            assert_eq!(max, 10);
        }
        other => panic!("expected XlsxError::PartTooLarge, got {other:?}"),
    }
}

#[test]
fn read_part_from_reader_limited_reads_small_parts() {
    let bytes = build_zip_bytes(&[("xl/workbook.xml", b"hello")]);

    let out =
        read_part_from_reader_limited(Cursor::new(bytes), "xl/workbook.xml", 10).unwrap();
    assert_eq!(out.as_deref(), Some(b"hello".as_slice()));
}

#[test]
fn xlsxpackage_from_bytes_limited_enforces_total_budget() {
    let bytes = build_zip_bytes(&[
        ("xl/a.bin", b"0123456789"), // 10 bytes
        ("xl/b.bin", b"0123456789"), // 10 bytes
        ("xl/c.bin", b"0123456789"), // 10 bytes
    ]);

    let limits = XlsxPackageLimits {
        max_part_bytes: 10,
        max_total_bytes: 20,
    };
    let err = XlsxPackage::from_bytes_limited(&bytes, limits)
        .expect_err("expected total-budget error");
    match err {
        XlsxError::PackageTooLarge { total, max } => {
            assert_eq!(max, 20);
            assert!(
                total > max,
                "expected reported total ({total}) to exceed max ({max})"
            );
        }
        other => panic!(
            "expected XlsxError::PackageTooLarge, got {other:?}"
        ),
    }
}
