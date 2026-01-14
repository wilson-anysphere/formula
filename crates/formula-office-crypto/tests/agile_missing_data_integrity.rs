use formula_office_crypto::{
    decrypt_encrypted_package_ole, encrypt_package_to_ole, EncryptOptions, EncryptionScheme,
    HashAlgorithm, OfficeCryptoError,
};
use std::io::{Cursor, Read, Write};
use zip::write::FileOptions;

fn build_minimal_zip_bytes() -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    // `formula-office-crypto` keeps its dev-dependency on `zip` minimal (no `deflate` feature),
    // so use a stored (uncompressed) entry for this synthetic fixture.
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Stored);
    zip.start_file("hello.txt", options)
        .expect("start zip entry");
    zip.write_all(b"hello world").expect("write zip entry");
    zip.finish().expect("finish zip").into_inner()
}

fn strip_data_integrity_from_encryption_info(encryption_info: &[u8]) -> Vec<u8> {
    assert!(
        encryption_info.len() >= 8,
        "EncryptionInfo must include 8-byte header"
    );
    let header = &encryption_info[..8];
    let xml_bytes = &encryption_info[8..];
    let xml = std::str::from_utf8(xml_bytes).expect("EncryptionInfo XML should be UTF-8");
    let mut stripped = xml.to_string();

    // Remove the first `<dataIntegrity .../>` (self-closing) or `<dataIntegrity>...</dataIntegrity>`
    // element when present. The parser is namespace-tolerant; this is purely to synthesize a fixture
    // missing the element.
    if let Some(start) = stripped.find("<dataIntegrity") {
        if let Some(end_rel) = stripped[start..].find("/>") {
            stripped.replace_range(start..start + end_rel + 2, "");
        } else if let Some(end_rel) = stripped[start..].find("</dataIntegrity>") {
            stripped.replace_range(
                start..start + end_rel + "</dataIntegrity>".len(),
                "",
            );
        }
    }

    let mut out = Vec::new();
    out.extend_from_slice(header);
    out.extend_from_slice(stripped.as_bytes());
    out
}

#[test]
fn decrypt_agile_allows_missing_data_integrity_element() {
    let zip_bytes = build_minimal_zip_bytes();
    let password = "password";

    // Keep the test fast: use a small spinCount and SHA-256.
    let opts = EncryptOptions {
        scheme: EncryptionScheme::Agile,
        key_bits: 128,
        hash_algorithm: HashAlgorithm::Sha256,
        spin_count: 1_000,
    };
    let ole_bytes = encrypt_package_to_ole(&zip_bytes, password, opts).expect("encrypt");

    // Extract streams from the OLE container.
    let mut ole = cfb::CompoundFile::open(Cursor::new(&ole_bytes)).expect("open cfb");
    let mut encryption_info = Vec::new();
    ole.open_stream("EncryptionInfo")
        .expect("open EncryptionInfo stream")
        .read_to_end(&mut encryption_info)
        .expect("read EncryptionInfo");
    let mut encrypted_package = Vec::new();
    ole.open_stream("EncryptedPackage")
        .expect("open EncryptedPackage stream")
        .read_to_end(&mut encrypted_package)
        .expect("read EncryptedPackage");

    let encryption_info = strip_data_integrity_from_encryption_info(&encryption_info);
    assert!(
        !std::str::from_utf8(&encryption_info[8..])
            .expect("utf-8")
            .contains("dataIntegrity"),
        "expected fixture to omit <dataIntegrity>"
    );

    // Rebuild the OLE container with the modified EncryptionInfo stream, keeping the same
    // EncryptedPackage ciphertext.
    let cursor = Cursor::new(Vec::new());
    let mut out = cfb::CompoundFile::create(cursor).expect("create cfb");
    out.create_stream("EncryptionInfo")
        .expect("create EncryptionInfo stream")
        .write_all(&encryption_info)
        .expect("write EncryptionInfo");
    out.create_stream("EncryptedPackage")
        .expect("create EncryptedPackage stream")
        .write_all(&encrypted_package)
        .expect("write EncryptedPackage");
    let modified_ole = out.into_inner().into_inner();

    let decrypted = decrypt_encrypted_package_ole(&modified_ole, password).expect("decrypt");
    assert_eq!(decrypted, zip_bytes);

    let err = decrypt_encrypted_package_ole(&modified_ole, "wrong-password").unwrap_err();
    assert!(
        matches!(err, OfficeCryptoError::InvalidPassword),
        "expected InvalidPassword, got: {err:?}"
    );
}
