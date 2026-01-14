use std::io::{Cursor, Read, Seek, SeekFrom, Write};

use formula_xlsx::StreamingXlsxPackage;
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

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

    let input = build_zip(&[
        ("xl/big.bin", CompressionMethod::Deflated, &big),
        ("xl/other.txt", CompressionMethod::Deflated, b"hello world"),
    ]);

    let input_compressed = read_zip_part_compressed_bytes(&input, "xl/big.bin");

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

