use std::io::{Cursor, Read, Write};

use formula_office_crypto::{
    decrypt_encrypted_package_ole, encrypt_package_to_ole, EncryptOptions, EncryptionScheme,
    HashAlgorithm,
};

fn minimal_zip_bytes() -> Vec<u8> {
    use zip::write::SimpleFileOptions;

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);

    // Avoid optional compression backends; Stored always works.
    let options = SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    zip.start_file("hello.txt", options)
        .expect("start zip file");
    zip.write_all(b"hello world").expect("write zip file");

    zip.finish().expect("finish zip").into_inner()
}

fn extract_streams_from_ole(ole_bytes: &[u8]) -> (Vec<u8>, Vec<u8>) {
    let cursor = Cursor::new(ole_bytes);
    let mut ole = cfb::CompoundFile::open(cursor).expect("open cfb");

    let mut encryption_info = Vec::new();
    ole.open_stream("EncryptionInfo")
        .expect("open EncryptionInfo")
        .read_to_end(&mut encryption_info)
        .expect("read EncryptionInfo");

    let mut encrypted_package = Vec::new();
    ole.open_stream("EncryptedPackage")
        .expect("open EncryptedPackage")
        .read_to_end(&mut encrypted_package)
        .expect("read EncryptedPackage");

    (encryption_info, encrypted_package)
}

fn build_ole(encryption_info: &[u8], encrypted_package: &[u8]) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    ole.create_stream("EncryptionInfo")
        .expect("create EncryptionInfo")
        .write_all(encryption_info)
        .expect("write EncryptionInfo");
    ole.create_stream("EncryptedPackage")
        .expect("create EncryptedPackage")
        .write_all(encrypted_package)
        .expect("write EncryptedPackage");
    ole.into_inner().into_inner()
}

#[test]
fn decrypt_agile_without_data_integrity_element() {
    // Some real-world producers omit `<dataIntegrity>` entirely. We should still be able to
    // decrypt the package (but without integrity verification).
    let plaintext = minimal_zip_bytes();
    let password = "Password";

    // Use a small spinCount for test speed.
    let opts = EncryptOptions {
        scheme: EncryptionScheme::Agile,
        key_bits: 256,
        hash_algorithm: HashAlgorithm::Sha512,
        spin_count: 512,
    };

    let baseline_ole = encrypt_package_to_ole(&plaintext, password, opts).expect("encrypt");
    let (baseline_info, baseline_package) = extract_streams_from_ole(&baseline_ole);

    // Extract the UTF-8 XML bytes by scanning for the first `<` after the version header.
    let payload = baseline_info
        .get(8..)
        .expect("baseline EncryptionInfo must include version header");
    let xml_start = payload.iter().position(|&b| b == b'<').expect("xml start");
    let mut xml_bytes = &payload[xml_start..];
    while xml_bytes.last() == Some(&0) {
        xml_bytes = &xml_bytes[..xml_bytes.len() - 1];
    }
    let xml_str = std::str::from_utf8(xml_bytes).expect("baseline xml utf8");

    // Remove the `<dataIntegrity .../>` element (self-closing in our writer).
    let di_start = xml_str
        .find("<dataIntegrity")
        .expect("expected baseline XML to include <dataIntegrity>");
    let di_end_rel = xml_str[di_start..]
        .find("/>")
        .expect("expected <dataIntegrity/> to be self-closing");
    let mut di_end = di_start + di_end_rel + 2;
    while matches!(xml_str.as_bytes().get(di_end), Some(b'\n' | b'\r')) {
        di_end += 1;
    }
    let xml_no_di = format!("{}{}", &xml_str[..di_start], &xml_str[di_end..]);

    // Build a new EncryptionInfo stream with just the 8-byte header + modified XML (no length
    // prefix, which is also a known producer variant).
    let mut info_no_di = Vec::new();
    info_no_di.extend_from_slice(&baseline_info[..8]);
    info_no_di.extend_from_slice(xml_no_di.as_bytes());

    let ole_no_di = build_ole(&info_no_di, &baseline_package);
    let decrypted =
        decrypt_encrypted_package_ole(&ole_no_di, password).expect("decrypt no dataIntegrity");
    assert_eq!(decrypted, plaintext);
}

#[test]
fn decrypt_agile_encryption_info_real_world_variants() {
    // Keep the plaintext small to keep the test fast, but ensure it is a valid ZIP archive (the
    // decrypter performs lightweight ZIP validation).
    let plaintext = minimal_zip_bytes();
    let password = "Password";

    // Use a small spinCount for test speed.
    let opts = EncryptOptions {
        scheme: EncryptionScheme::Agile,
        key_bits: 256,
        hash_algorithm: HashAlgorithm::Sha512,
        spin_count: 512,
    };

    let baseline_ole = encrypt_package_to_ole(&plaintext, password, opts).expect("encrypt");
    let (baseline_info, baseline_package) = extract_streams_from_ole(&baseline_ole);

    // Baseline encryption info is: 8-byte version header followed by some producer-specific
    // wrapping/encoding of the XML descriptor. Extract the UTF-8 XML bytes by scanning for the
    // first `<` after the version header.
    let payload = baseline_info
        .get(8..)
        .expect("baseline EncryptionInfo must include version header");
    let xml_start = payload.iter().position(|&b| b == b'<').expect("xml start");
    let mut xml_bytes = &payload[xml_start..];
    while xml_bytes.last() == Some(&0) {
        xml_bytes = &xml_bytes[..xml_bytes.len() - 1];
    }
    let xml_str = std::str::from_utf8(xml_bytes).expect("baseline xml utf8");

    // --- Variant 1: UTF-8 with BOM and trailing NUL padding (still length-prefixed). ---
    let mut xml_bom_nul = Vec::new();
    xml_bom_nul.extend_from_slice(&[0xEF, 0xBB, 0xBF]); // UTF-8 BOM
    xml_bom_nul.extend_from_slice(xml_bytes);
    xml_bom_nul.extend_from_slice(&[0, 0, 0, 0, 0]); // padding

    let mut info_bom_nul = Vec::new();
    info_bom_nul.extend_from_slice(&baseline_info[..8]);
    info_bom_nul.extend_from_slice(&(xml_bom_nul.len() as u32).to_le_bytes());
    info_bom_nul.extend_from_slice(&xml_bom_nul);

    let ole_bom_nul = build_ole(&info_bom_nul, &baseline_package);
    let decrypted = decrypt_encrypted_package_ole(&ole_bom_nul, password).expect("decrypt bom+nul");
    assert_eq!(decrypted, plaintext);

    // --- Variant 2: UTF-16LE-encoded XML (length-prefixed). ---
    let mut xml_utf16le = Vec::new();
    xml_utf16le.extend_from_slice(&[0xFF, 0xFE]); // UTF-16LE BOM
    for cu in xml_str.encode_utf16() {
        xml_utf16le.extend_from_slice(&cu.to_le_bytes());
    }
    xml_utf16le.extend_from_slice(&[0x00, 0x00]); // terminating NUL

    let mut info_utf16le = Vec::new();
    info_utf16le.extend_from_slice(&baseline_info[..8]);
    info_utf16le.extend_from_slice(&(xml_utf16le.len() as u32).to_le_bytes());
    info_utf16le.extend_from_slice(&xml_utf16le);

    let ole_utf16le = build_ole(&info_utf16le, &baseline_package);
    let decrypted = decrypt_encrypted_package_ole(&ole_utf16le, password).expect("decrypt utf16le");
    assert_eq!(decrypted, plaintext);

    // --- Variant 3: no outer headerSize field (XML begins immediately after 8-byte header). ---
    let mut info_no_len = Vec::new();
    info_no_len.extend_from_slice(&baseline_info[..8]);
    info_no_len.extend_from_slice(xml_bytes);

    let ole_no_len = build_ole(&info_no_len, &baseline_package);
    let decrypted =
        decrypt_encrypted_package_ole(&ole_no_len, password).expect("decrypt no-len-prefix");
    assert_eq!(decrypted, plaintext);
}
