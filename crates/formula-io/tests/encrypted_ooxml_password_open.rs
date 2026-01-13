use std::io::{Cursor, Write};

use formula_io::{open_workbook_with_password, Workbook};

fn simple_xlsx_bytes() -> Vec<u8> {
    let mut wb = formula_model::Workbook::new();
    wb.add_sheet("Sheet1").expect("add sheet");

    let mut cursor = Cursor::new(Vec::new());
    formula_xlsx::write_workbook_to_writer(&wb, &mut cursor).expect("write workbook");
    cursor.into_inner()
}

#[test]
fn opens_encrypted_ooxml_via_password_api_when_streams_use_leading_slash_paths() {
    // Build an OOXML payload (a valid `.xlsx` ZIP).
    let xlsx_bytes = simple_xlsx_bytes();

    // Wrap it in the standard OLE encryption container shape: `EncryptionInfo` + `EncryptedPackage`.
    //
    // This test uses the leading-slash path form when creating the streams. Some real-world OLE
    // writers (and/or older `cfb` versions) have been observed to require the leading slash when
    // opening these streams, so `formula-io` should try both `name` and `/{name}` when reading.
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
        pkg.write_all(&(xlsx_bytes.len() as u64).to_le_bytes())
            .expect("write EncryptedPackage size prefix");
        pkg.write_all(&xlsx_bytes)
            .expect("write EncryptedPackage payload");
    }

    let ole_bytes = ole.into_inner().into_inner();

    let tmp = tempfile::tempdir().expect("tempdir");
    let path = tmp.path().join("encrypted.xlsx");
    std::fs::write(&path, &ole_bytes).expect("write OLE container");

    let wb = open_workbook_with_password(&path, Some("password")).expect("open via password API");
    let Workbook::Xlsx(pkg) = wb else {
        panic!("expected Workbook::Xlsx from decrypted EncryptedPackage");
    };

    assert!(
        pkg.part("xl/workbook.xml").is_some(),
        "expected extracted package to be a valid XLSX zip"
    );
}
