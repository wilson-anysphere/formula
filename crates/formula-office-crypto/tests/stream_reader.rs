use std::io::{Read, Seek, SeekFrom};

use formula_office_crypto::{
    decrypt_encrypted_package_ole, decrypt_encrypted_package_ole_to_reader, encrypt_package_to_ole,
    EncryptOptions,
};

fn make_plaintext(len: usize) -> Vec<u8> {
    let mut out = Vec::new();
    let _ = out.try_reserve_exact(len);
    out.extend_from_slice(b"PK"); // satisfy ZIP signature check
    for i in 2..len {
        out.push((i % 251) as u8);
    }
    out
}

#[test]
fn stream_reader_round_trips_and_supports_seek() {
    // Ensure we cross 4096-byte segment boundaries.
    let plain = make_plaintext(10_000);
    let password = "password";

    let ole = encrypt_package_to_ole(&plain, password, EncryptOptions::default()).expect("encrypt");
    let mut reader = decrypt_encrypted_package_ole_to_reader(&ole, password).expect("decrypt reader");

    // Read a small prefix.
    let mut prefix = vec![0u8; 32];
    reader.read_exact(&mut prefix).expect("read prefix");
    assert_eq!(&prefix, &plain[..32]);

    // Seek across a segment boundary and read.
    reader.seek(SeekFrom::Start(4090)).expect("seek");
    let mut buf = vec![0u8; 32];
    reader.read_exact(&mut buf).expect("read boundary span");
    assert_eq!(&buf, &plain[4090..4090 + 32]);

    // Random seek/read.
    for &offset in &[0u64, 123, 4096, 8191, 9000] {
        reader.seek(SeekFrom::Start(offset)).expect("seek");
        let mut tmp = vec![0u8; 17];
        reader.read_exact(&mut tmp).expect("read");
        assert_eq!(&tmp, &plain[offset as usize..offset as usize + 17]);
    }

    // Read to end matches original.
    reader.seek(SeekFrom::Start(0)).expect("rewind");
    let mut all = Vec::new();
    reader.read_to_end(&mut all).expect("read_to_end");
    assert_eq!(all, plain);
}

#[test]
fn stream_reader_decrypts_standard_fixture() {
    let path = std::path::Path::new(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/encrypted/ooxml/standard.xlsx"
    ));
    if !path.exists() {
        // Allow this test to land before fixtures are present in all environments.
        return;
    }

    let ole = std::fs::read(path).expect("read standard fixture");

    // Sanity: the Vec<u8> decrypt path should succeed for the real-world fixture.
    let decrypted = decrypt_encrypted_package_ole(&ole, "password").expect("decrypt standard bytes");
    assert!(decrypted.starts_with(b"PK"), "expected decrypted bytes to be a ZIP");

    let mut reader =
        decrypt_encrypted_package_ole_to_reader(&ole, "password").expect("decrypt standard reader");

    let mut header = [0u8; 2];
    reader.read_exact(&mut header).expect("read header");
    assert_eq!(&header, b"PK");
}

#[test]
fn stream_reader_decrypts_standard_rc4_fixture() {
    let path = std::path::Path::new(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/encrypted/ooxml/standard-rc4.xlsx"
    ));
    if !path.exists() {
        // Allow this test to land before fixtures are present in all environments.
        return;
    }

    let ole = std::fs::read(path).expect("read standard rc4 fixture");

    // Sanity: the Vec<u8> decrypt path should succeed for the real-world fixture.
    let decrypted =
        decrypt_encrypted_package_ole(&ole, "password").expect("decrypt standard rc4 bytes");
    assert!(decrypted.starts_with(b"PK"), "expected decrypted bytes to be a ZIP");

    let mut reader = decrypt_encrypted_package_ole_to_reader(&ole, "password")
        .expect("decrypt standard rc4 reader");

    let mut header = [0u8; 2];
    reader.read_exact(&mut header).expect("read header");
    assert_eq!(&header, b"PK");

    // Seek/read should match the Vec-based decryptor.
    let offset = 123u64.min(decrypted.len().saturating_sub(32) as u64);
    reader.seek(SeekFrom::Start(offset)).expect("seek");
    let mut buf = [0u8; 32];
    reader.read_exact(&mut buf).expect("read span");
    assert_eq!(&buf, &decrypted[offset as usize..offset as usize + 32]);

    // Full read matches Vec-based decryptor.
    reader.seek(SeekFrom::Start(0)).expect("rewind");
    let mut all = Vec::new();
    reader.read_to_end(&mut all).expect("read_to_end");
    assert_eq!(all, decrypted);
}
