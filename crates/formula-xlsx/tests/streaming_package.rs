use std::io::{Cursor, Read, Seek, SeekFrom, Write};

use formula_xlsx::StreamingXlsxPackage;
use zip::write::FileOptions;
use zip::{CompressionMethod, DateTime, ZipArchive, ZipWriter};

fn build_zip(entries: &[(&str, CompressionMethod, &[u8])]) -> Vec<u8> {
    let mut cursor = Cursor::new(Vec::new());
    {
        let mut zip = ZipWriter::new(&mut cursor);
        for (name, method, bytes) in entries {
            let options = FileOptions::<()>::default().compression_method(*method);
            zip.start_file(name, options).unwrap();
            zip.write_all(bytes).unwrap();
        }
        zip.finish().unwrap();
    }
    cursor.into_inner()
}

fn read_zip_part(zip_bytes: &[u8], name: &str) -> Vec<u8> {
    let mut zip = ZipArchive::new(Cursor::new(zip_bytes)).unwrap();
    let mut file = zip.by_name(name).unwrap();
    let mut out = Vec::new();
    file.read_to_end(&mut out).unwrap();
    out
}

fn zip_part_last_modified(zip_bytes: &[u8], name: &str) -> DateTime {
    let mut zip = ZipArchive::new(Cursor::new(zip_bytes)).unwrap();
    let ts = zip
        .by_name(name)
        .unwrap()
        .last_modified()
        .expect("zip entry missing last_modified time");
    ts
}

fn read_zip_part_compressed_bytes(zip_bytes: &[u8], name: &str) -> Vec<u8> {
    let mut zip = ZipArchive::new(Cursor::new(zip_bytes)).unwrap();
    let file = zip.by_name(name).unwrap();
    let start = file.data_start();
    let len = file.compressed_size();
    drop(file);

    let mut reader = zip.into_inner();
    reader.seek(SeekFrom::Start(start)).unwrap();
    let mut out = vec![0u8; len as usize];
    reader.read_exact(&mut out).unwrap();
    out
}

#[test]
fn streaming_package_write_to_raw_copies_large_part() {
    // Highly compressible payload so that "raw-copy" vs "store uncompressed" is obvious.
    let big = vec![0u8; 5 * 1024 * 1024];

    // Also set a non-default timestamp on the entry so we can detect whether it was
    // raw-copied (timestamp preserved) vs rewritten (timestamp would change).
    let big_ts = DateTime::from_date_and_time(2001, 2, 3, 4, 5, 6).unwrap();
    let input = {
        let mut cursor = Cursor::new(Vec::new());
        {
            let mut zip = ZipWriter::new(&mut cursor);
            let options = FileOptions::<()>::default()
                .compression_method(CompressionMethod::Deflated)
                .last_modified_time(big_ts);
            zip.start_file("xl/big.bin", options).unwrap();
            zip.write_all(&big).unwrap();

            let options =
                FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);
            zip.start_file("xl/other.txt", options).unwrap();
            zip.write_all(b"hello world").unwrap();

            zip.finish().unwrap();
        }
        cursor.into_inner()
    };

    let input_compressed = read_zip_part_compressed_bytes(&input, "xl/big.bin");
    let input_ts = zip_part_last_modified(&input, "xl/big.bin");

    let pkg = StreamingXlsxPackage::from_reader(Cursor::new(input.clone())).unwrap();
    let mut out = Cursor::new(Vec::new());
    pkg.write_to(&mut out).unwrap();
    let output = out.into_inner();

    // Uncompressed bytes must match.
    let output_big = read_zip_part(&output, "xl/big.bin");
    assert_eq!(output_big, big);

    // The compressed bytes should be identical when the entry is raw-copied.
    let output_compressed = read_zip_part_compressed_bytes(&output, "xl/big.bin");
    assert_eq!(output_compressed, input_compressed);

    // Raw-copy should also preserve entry metadata like timestamps.
    let output_ts = zip_part_last_modified(&output, "xl/big.bin");
    assert_eq!(output_ts, input_ts);

    // And the overall file size should not balloon toward the uncompressed payload size.
    assert!(
        output.len() < 200_000,
        "output ZIP unexpectedly large: {} bytes",
        output.len()
    );
}

#[test]
fn streaming_package_set_part_and_remove_part() {
    let input = build_zip(&[
        ("xl/workbook.xml", CompressionMethod::Deflated, b"old"),
        ("xl/to_remove.bin", CompressionMethod::Deflated, b"bye"),
    ]);

    let mut pkg = StreamingXlsxPackage::from_reader(Cursor::new(input)).unwrap();
    pkg.set_part("xl/workbook.xml", b"new".to_vec());
    pkg.remove_part("xl/to_remove.bin");

    let mut out = Cursor::new(Vec::new());
    pkg.write_to(&mut out).unwrap();
    let output = out.into_inner();

    assert_eq!(read_zip_part(&output, "xl/workbook.xml"), b"new");

    let mut zip = ZipArchive::new(Cursor::new(output)).unwrap();
    assert!(zip.by_name("xl/to_remove.bin").is_err());
}

#[test]
fn streaming_package_normalizes_backslashes_and_leading_slash() {
    let input = build_zip(&[
        // Non-canonical ZIP entry name (`\\` separator) seen in some broken producers.
        ("xl\\workbook.xml", CompressionMethod::Deflated, b"old"),
        // Also exercise leading `/` mismatch.
        ("/xl/keep.txt", CompressionMethod::Deflated, b"keep"),
    ]);

    let mut pkg = StreamingXlsxPackage::from_reader(Cursor::new(input)).unwrap();

    assert_eq!(
        pkg.read_part("xl/workbook.xml").unwrap().as_deref(),
        Some(b"old".as_slice())
    );
    // Canonical part names should be surfaced through part_names().
    let names: Vec<String> = pkg.part_names().collect();
    assert!(names.iter().any(|n| n == "xl/workbook.xml"));
    assert!(names.iter().any(|n| n == "xl/keep.txt"));

    pkg.set_part("/xl/workbook.xml", b"new".to_vec());
    pkg.remove_part("xl/keep.txt");

    let mut out = Cursor::new(Vec::new());
    pkg.write_to(&mut out).unwrap();
    let output = out.into_inner();

    // Replaced entry should still be found under its original ZIP name.
    assert_eq!(read_zip_part(&output, "xl\\workbook.xml"), b"new");

    // Removed entry should be absent (regardless of leading `/`).
    let mut zip = ZipArchive::new(Cursor::new(output)).unwrap();
    assert!(zip.by_name("/xl/keep.txt").is_err());
    assert!(zip.by_name("xl/keep.txt").is_err());
}
