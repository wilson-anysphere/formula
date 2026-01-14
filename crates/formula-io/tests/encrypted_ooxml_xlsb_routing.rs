use std::io::{Cursor, Write};

use formula_io::{open_workbook_with_password, Error};

fn xlsb_like_zip_bytes() -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options =
        zip::write::FileOptions::<()>::default().compression_method(zip::CompressionMethod::Stored);
    zip.start_file("xl/workbook.bin", options)
        .expect("start xl/workbook.bin");
    zip.write_all(b"not a real workbook.bin payload")
        .expect("write workbook.bin");
    zip.finish().expect("finish zip").into_inner()
}

#[test]
fn routes_workbook_bin_packages_to_xlsb_open_path() {
    // Build a minimal ZIP that looks like an XLSB package (`xl/workbook.bin` is present) but is
    // intentionally invalid so we can assert that `formula-io` routed it to the XLSB opener.
    let zip_bytes = xlsb_like_zip_bytes();

    // Wrap it in the standard Office-encrypted OOXML OLE container shape: `EncryptionInfo` +
    // `EncryptedPackage`. This fixture places the plaintext ZIP bytes directly in `EncryptedPackage`
    // (with the usual 8-byte length prefix), which `formula-io` supports for synthetic fixtures.
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");

    {
        let mut info = ole
            .create_stream("/EncryptionInfo")
            .expect("create EncryptionInfo stream");
        // Minimal EncryptionInfo header:
        // - VersionMajor = 4
        // - VersionMinor = 4 (Agile encryption)
        // - Flags = 0
        info.write_all(&[4, 0, 4, 0, 0, 0, 0, 0])
            .expect("write EncryptionInfo header");
    }

    {
        let mut pkg = ole
            .create_stream("/EncryptedPackage")
            .expect("create EncryptedPackage stream");
        // MS-OFFCRYPTO uses a u64 little-endian plaintext size prefix.
        pkg.write_all(&(zip_bytes.len() as u64).to_le_bytes())
            .expect("write EncryptedPackage size prefix");
        pkg.write_all(&zip_bytes).expect("write EncryptedPackage payload");
    }

    let ole_bytes = ole.into_inner().into_inner();

    let tmp = tempfile::tempdir().expect("tempdir");
    let path = tmp.path().join("encrypted.xlsb");
    std::fs::write(&path, &ole_bytes).expect("write OLE container");

    match open_workbook_with_password(&path, Some("password")) {
        Err(Error::OpenXlsb { path: err_path, .. }) => {
            assert_eq!(err_path, path);
        }
        other => panic!("expected OpenXlsb error, got {other:?}"),
    }
}
