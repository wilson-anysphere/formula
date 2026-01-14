#![cfg(not(target_arch = "wasm32"))]

use std::io::{Cursor, Read, Seek, Write};

use aes::Aes128;
use base64::engine::general_purpose::STANDARD as BASE64;
use base64::Engine as _;
use cbc::cipher::block_padding::NoPadding;
use cbc::cipher::{BlockEncryptMut, KeyIvInit};
use formula_xlsx::offcrypto::decrypt_agile_encrypted_package_stream;
use formula_xlsx::offcrypto::{
    derive_iv, derive_key, hash_password, HashAlgorithm, KEY_VALUE_BLOCK, VERIFIER_HASH_INPUT_BLOCK,
    VERIFIER_HASH_VALUE_BLOCK,
};
use formula_xlsx::OffCryptoError;
use ms_offcrypto_writer::Ecma376AgileWriter;
use rand::{rngs::StdRng, SeedableRng as _};
use sha1::Digest as _;
use zip::write::FileOptions;

fn make_zip_bytes(payload_len: usize) -> Vec<u8> {
    let payload: Vec<u8> = (0..payload_len).map(|i| (i % 251) as u8).collect();

    let mut cursor = Cursor::new(Vec::new());
    {
        let mut zip = zip::ZipWriter::new(&mut cursor);
        let opts = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Stored);
        zip.start_file("data.bin", opts).expect("start zip entry");
        zip.write_all(&payload).expect("write payload");
        zip.finish().expect("finish zip");
    }
    cursor.into_inner()
}

fn open_stream<R: Read + Seek + Write>(
    ole: &mut cfb::CompoundFile<R>,
    name: &str,
) -> cfb::Stream<R> {
    ole.open_stream(name)
        .or_else(|_| ole.open_stream(&format!("/{name}")))
        .expect("open OLE stream")
}

#[test]
fn decrypts_agile_encrypted_package_streaming() {
    let password = "correct horse battery staple";
    // Ensure we span multiple 4096-byte chunks and require truncation (not a multiple of 16).
    let plaintext = make_zip_bytes(12_345);

    let mut rng = StdRng::seed_from_u64(0xD15EA5E_u64);
    let cursor = Cursor::new(Vec::new());
    let mut writer =
        Ecma376AgileWriter::create(&mut rng, password, cursor).expect("create agile writer");
    writer
        .write_all(&plaintext)
        .expect("write plaintext package bytes");
    let cursor = writer.into_inner().expect("finalize agile writer");
    let encrypted_ole_bytes = cursor.into_inner();

    let mut ole = cfb::CompoundFile::open(Cursor::new(encrypted_ole_bytes)).expect("open cfb");

    let mut encryption_info_stream = open_stream(&mut ole, "EncryptionInfo");
    let mut encryption_info = Vec::new();
    encryption_info_stream
        .read_to_end(&mut encryption_info)
        .expect("read EncryptionInfo");

    let mut encrypted_package_stream = open_stream(&mut ole, "EncryptedPackage");
    let mut out = Vec::new();
    let declared_len = decrypt_agile_encrypted_package_stream(
        &encryption_info,
        &mut encrypted_package_stream,
        password,
        &mut out,
    )
    .expect("decrypt agile encrypted package");

    assert_eq!(declared_len as usize, plaintext.len());
    assert_eq!(out, plaintext);
}

#[test]
fn decrypts_agile_encrypted_package_streaming_tampered_size_header_fails_integrity() {
    let password = "correct horse battery staple";
    let plaintext = make_zip_bytes(12_345);

    let mut rng = StdRng::seed_from_u64(0xD15EA5E_u64);
    let cursor = Cursor::new(Vec::new());
    let mut writer =
        Ecma376AgileWriter::create(&mut rng, password, cursor).expect("create agile writer");
    writer
        .write_all(&plaintext)
        .expect("write plaintext package bytes");
    let cursor = writer.into_inner().expect("finalize agile writer");
    let encrypted_ole_bytes = cursor.into_inner();

    let mut ole = cfb::CompoundFile::open(Cursor::new(encrypted_ole_bytes)).expect("open cfb");

    let mut encryption_info_stream = open_stream(&mut ole, "EncryptionInfo");
    let mut encryption_info = Vec::new();
    encryption_info_stream
        .read_to_end(&mut encryption_info)
        .expect("read EncryptionInfo");

    let mut encrypted_package_stream = open_stream(&mut ole, "EncryptedPackage");
    let mut encrypted_package = Vec::new();
    encrypted_package_stream
        .read_to_end(&mut encrypted_package)
        .expect("read EncryptedPackage");

    assert!(encrypted_package.len() >= 8, "EncryptedPackage too short");
    let original_size = u64::from_le_bytes(
        encrypted_package[..8]
            .try_into()
            .expect("EncryptedPackage header is 8 bytes"),
    );
    assert!(original_size > 0, "unexpected empty EncryptedPackage payload");
    let tampered_size = original_size - 1;
    encrypted_package[..8].copy_from_slice(&tampered_size.to_le_bytes());

    let mut cursor = Cursor::new(encrypted_package);
    let mut out = Vec::new();
    let err = decrypt_agile_encrypted_package_stream(
        &encryption_info,
        &mut cursor,
        password,
        &mut out,
    )
    .expect_err("expected integrity failure");
    assert!(
        matches!(err, OffCryptoError::IntegrityMismatch),
        "expected IntegrityMismatch, got {err:?}"
    );
}

#[test]
fn decrypts_agile_encrypted_package_streaming_appended_ciphertext_fails_integrity() {
    let password = "correct horse battery staple";
    let plaintext = make_zip_bytes(12_345);

    let mut rng = StdRng::seed_from_u64(0xD15EA5E_u64);
    let cursor = Cursor::new(Vec::new());
    let mut writer =
        Ecma376AgileWriter::create(&mut rng, password, cursor).expect("create agile writer");
    writer
        .write_all(&plaintext)
        .expect("write plaintext package bytes");
    let cursor = writer.into_inner().expect("finalize agile writer");
    let encrypted_ole_bytes = cursor.into_inner();

    let mut ole = cfb::CompoundFile::open(Cursor::new(encrypted_ole_bytes)).expect("open cfb");

    let mut encryption_info_stream = open_stream(&mut ole, "EncryptionInfo");
    let mut encryption_info = Vec::new();
    encryption_info_stream
        .read_to_end(&mut encryption_info)
        .expect("read EncryptionInfo");

    let mut encrypted_package_stream = open_stream(&mut ole, "EncryptedPackage");
    let mut encrypted_package = Vec::new();
    encrypted_package_stream
        .read_to_end(&mut encrypted_package)
        .expect("read EncryptedPackage");

    // Append an extra AES block to simulate trailing bytes stored in the stream. `dataIntegrity`
    // authenticates the entire EncryptedPackage stream bytes, so this should be detected.
    encrypted_package.extend_from_slice(&[0xA5u8; 16]);

    let mut cursor = Cursor::new(encrypted_package);
    let mut out = Vec::new();
    let err = decrypt_agile_encrypted_package_stream(
        &encryption_info,
        &mut cursor,
        password,
        &mut out,
    )
    .expect_err("expected integrity failure");
    assert!(
        matches!(err, OffCryptoError::IntegrityMismatch),
        "expected IntegrityMismatch, got {err:?}"
    );
}

fn encrypt_aes128_cbc_no_padding(key: &[u8], iv: &[u8], plaintext: &[u8]) -> Vec<u8> {
    assert_eq!(key.len(), 16);
    assert_eq!(iv.len(), 16);
    assert!(
        plaintext.len() % 16 == 0,
        "plaintext must be AES-block aligned"
    );

    let mut buf = plaintext.to_vec();
    let len = buf.len();
    cbc::Encryptor::<Aes128>::new_from_slices(key, iv)
        .expect("valid key/iv")
        .encrypt_padded_mut::<NoPadding>(&mut buf, len)
        .expect("encrypt");
    buf
}

fn zero_pad_to_block(mut bytes: Vec<u8>, block_size: usize) -> Vec<u8> {
    let rem = bytes.len() % block_size;
    if rem != 0 {
        bytes.extend(std::iter::repeat(0u8).take(block_size - rem));
    }
    bytes
}

#[test]
fn decrypts_agile_encrypted_package_streaming_with_derived_password_key_iv() {
    // Build a synthetic Agile descriptor where the password key-encryptor blobs
    // (`encryptedVerifierHashInput`, `encryptedVerifierHashValue`, `encryptedKeyValue`) are encrypted
    // using per-blob derived IVs (Hash(saltValue || blockKey)[:blockSize]) instead of Excel's typical
    // `IV = saltValue[..blockSize]`.
    let password = "pw";
    let plaintext = make_zip_bytes(12_345);

    let hash_alg = HashAlgorithm::Sha1;
    let hash_size = 20usize;
    let block_size = 16usize;
    let key_encrypt_key_len = 16usize;

    let key_data_salt: Vec<u8> = (0u8..=15).collect();
    let password_salt: Vec<u8> = (16u8..=31).collect();
    let spin_count = 10u32;

    let package_key: Vec<u8> = (32u8..=47).collect(); // AES-128 package key

    // --- Build EncryptedPackage stream (segmented) ---------------------------------------------
    let mut encrypted_package = Vec::new();
    encrypted_package.extend_from_slice(&(plaintext.len() as u64).to_le_bytes());
    for (segment_index, segment) in plaintext.chunks(0x1000).enumerate() {
        let padded = zero_pad_to_block(segment.to_vec(), block_size);
        let iv = derive_iv(
            &key_data_salt,
            &(segment_index as u32).to_le_bytes(),
            block_size,
            hash_alg,
        )
        .unwrap();
        let ct = encrypt_aes128_cbc_no_padding(&package_key, &iv, &padded);
        encrypted_package.extend_from_slice(&ct);
    }

    // --- Build password key encryptor fields ---------------------------------------------------
    let password_hash = hash_password(password, &password_salt, spin_count, hash_alg).unwrap();
    let salt_iv = &password_salt[..block_size];
    let derived_iv = derive_iv(
        &password_salt,
        &VERIFIER_HASH_INPUT_BLOCK,
        block_size,
        hash_alg,
    )
    .unwrap();
    assert_ne!(
        derived_iv.as_slice(),
        salt_iv,
        "derived-IV scheme should not accidentally match Excel's saltValue IV"
    );

    let verifier_input: Vec<u8> = b"abcdefghijklmnop".to_vec();
    let verifier_hash: Vec<u8> = sha1::Sha1::digest(&verifier_input).to_vec();

    // Make verifierHashValue plaintext block-aligned by appending non-zero garbage after the digest.
    let mut verifier_hash_value_plain = verifier_hash.clone();
    verifier_hash_value_plain.extend_from_slice(&[0xA5u8; 12]);
    assert_eq!(verifier_hash_value_plain.len(), 32);

    let encrypt_pw_blob = |block_key: &[u8], plaintext: &[u8]| -> Vec<u8> {
        let k = derive_key(&password_hash, block_key, key_encrypt_key_len, hash_alg).unwrap();
        let iv = derive_iv(&password_salt, block_key, block_size, hash_alg).unwrap();
        encrypt_aes128_cbc_no_padding(&k, &iv, plaintext)
    };

    let encrypted_verifier_hash_input = encrypt_pw_blob(&VERIFIER_HASH_INPUT_BLOCK, &verifier_input);
    let encrypted_verifier_hash_value =
        encrypt_pw_blob(&VERIFIER_HASH_VALUE_BLOCK, &verifier_hash_value_plain);
    let encrypted_key_value = encrypt_pw_blob(&KEY_VALUE_BLOCK, &package_key);

    // --- Build EncryptionInfo stream (no dataIntegrity) ----------------------------------------
    let xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
            xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
  <keyData saltSize="16" blockSize="{block_size}" keyBits="128" hashSize="{hash_size}"
           cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA1"
           saltValue="{key_data_salt_b64}"/>
  <keyEncryptors>
    <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
      <p:encryptedKey saltSize="16" blockSize="{block_size}" keyBits="128" hashSize="{hash_size}"
                      spinCount="{spin_count}" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA1"
                      saltValue="{password_salt_b64}"
                      encryptedVerifierHashInput="{evhi_b64}"
                      encryptedVerifierHashValue="{evhv_b64}"
                      encryptedKeyValue="{ekv_b64}"/>
    </keyEncryptor>
  </keyEncryptors>
</encryption>"#,
        key_data_salt_b64 = BASE64.encode(&key_data_salt),
        password_salt_b64 = BASE64.encode(&password_salt),
        evhi_b64 = BASE64.encode(&encrypted_verifier_hash_input),
        evhv_b64 = BASE64.encode(&encrypted_verifier_hash_value),
        ekv_b64 = BASE64.encode(&encrypted_key_value),
    );

    let mut encryption_info = Vec::new();
    encryption_info.extend_from_slice(&4u16.to_le_bytes()); // major
    encryption_info.extend_from_slice(&4u16.to_le_bytes()); // minor
    encryption_info.extend_from_slice(&0u32.to_le_bytes()); // flags
    encryption_info.extend_from_slice(xml.as_bytes());

    let mut encrypted_package_stream = Cursor::new(encrypted_package);
    let mut out = Vec::new();
    let declared_len = decrypt_agile_encrypted_package_stream(
        &encryption_info,
        &mut encrypted_package_stream,
        password,
        &mut out,
    )
    .expect("decrypt derived-IV Agile stream");

    assert_eq!(declared_len as usize, plaintext.len());
    assert_eq!(out, plaintext);

    let mut encrypted_package_stream = Cursor::new(encrypted_package_stream.into_inner());
    let mut out = Vec::new();
    let err = decrypt_agile_encrypted_package_stream(
        &encryption_info,
        &mut encrypted_package_stream,
        "wrong-password",
        &mut out,
    )
    .expect_err("wrong password should fail");
    assert!(
        matches!(err, formula_xlsx::offcrypto::OffCryptoError::WrongPassword),
        "expected WrongPassword, got {err:?}"
    );
}

#[test]
fn decrypts_agile_encrypted_package_streaming_without_data_integrity() {
    let password = "correct horse battery staple";
    let plaintext = make_zip_bytes(12_345);

    let mut rng = StdRng::seed_from_u64(0xD15EA5E_u64);
    let cursor = Cursor::new(Vec::new());
    let mut writer =
        Ecma376AgileWriter::create(&mut rng, password, cursor).expect("create agile writer");
    writer
        .write_all(&plaintext)
        .expect("write plaintext package bytes");
    let cursor = writer.into_inner().expect("finalize agile writer");
    let encrypted_ole_bytes = cursor.into_inner();

    let mut ole = cfb::CompoundFile::open(Cursor::new(encrypted_ole_bytes)).expect("open cfb");

    let mut encryption_info_stream = open_stream(&mut ole, "EncryptionInfo");
    let mut encryption_info = Vec::new();
    encryption_info_stream
        .read_to_end(&mut encryption_info)
        .expect("read EncryptionInfo");

    // Patch the EncryptionInfo XML to remove `<dataIntegrity .../>`.
    let xml_start = encryption_info
        .iter()
        .position(|b| *b == b'<')
        .expect("EncryptionInfo must contain XML");
    let header = encryption_info[..xml_start].to_vec();
    let xml =
        std::str::from_utf8(&encryption_info[xml_start..]).expect("EncryptionInfo XML is UTF-8");
    let start = xml
        .find("<dataIntegrity")
        .expect("expected <dataIntegrity> element");
    let end = if let Some(end_rel) = xml[start..].find("/>") {
        start + end_rel + 2
    } else if let Some(end_rel) = xml[start..].find("</dataIntegrity>") {
        start + end_rel + "</dataIntegrity>".len()
    } else {
        panic!("expected </dataIntegrity> or />");
    };
    let mut patched_xml = String::new();
    patched_xml.push_str(&xml[..start]);
    patched_xml.push_str(&xml[end..]);
    encryption_info = header.into_iter().chain(patched_xml.into_bytes()).collect();

    let mut encrypted_package_stream = open_stream(&mut ole, "EncryptedPackage");
    let mut out = Vec::new();
    let declared_len = decrypt_agile_encrypted_package_stream(
        &encryption_info,
        &mut encrypted_package_stream,
        password,
        &mut out,
    )
    .expect("decrypt agile encrypted package without integrity");

    assert_eq!(declared_len as usize, plaintext.len());
    assert_eq!(out, plaintext);
}

#[test]
fn decrypts_agile_encrypted_package_streaming_when_size_header_high_dword_is_reserved() {
    let password = "correct horse battery staple";
    let plaintext = make_zip_bytes(12_345);

    let mut rng = StdRng::seed_from_u64(0xD15EA5E_u64);
    let cursor = Cursor::new(Vec::new());
    let mut writer =
        Ecma376AgileWriter::create(&mut rng, password, cursor).expect("create agile writer");
    writer
        .write_all(&plaintext)
        .expect("write plaintext package bytes");
    let cursor = writer.into_inner().expect("finalize agile writer");
    let encrypted_ole_bytes = cursor.into_inner();

    let mut ole = cfb::CompoundFile::open(Cursor::new(encrypted_ole_bytes)).expect("open cfb");

    let mut encryption_info_stream = open_stream(&mut ole, "EncryptionInfo");
    let mut encryption_info = Vec::new();
    encryption_info_stream
        .read_to_end(&mut encryption_info)
        .expect("read EncryptionInfo");

    // Patch the EncryptionInfo XML to remove `<dataIntegrity .../>` so we don't have to recompute
    // the HMAC after mutating the EncryptedPackage header bytes.
    let xml_start = encryption_info
        .iter()
        .position(|b| *b == b'<')
        .expect("EncryptionInfo must contain XML");
    let header = encryption_info[..xml_start].to_vec();
    let xml =
        std::str::from_utf8(&encryption_info[xml_start..]).expect("EncryptionInfo XML is UTF-8");
    let start = xml
        .find("<dataIntegrity")
        .expect("expected <dataIntegrity> element");
    let end = if let Some(end_rel) = xml[start..].find("/>") {
        start + end_rel + 2
    } else if let Some(end_rel) = xml[start..].find("</dataIntegrity>") {
        start + end_rel + "</dataIntegrity>".len()
    } else {
        panic!("expected </dataIntegrity> or />");
    };
    let mut patched_xml = String::new();
    patched_xml.push_str(&xml[..start]);
    patched_xml.push_str(&xml[end..]);
    encryption_info = header.into_iter().chain(patched_xml.into_bytes()).collect();

    // Mutate the EncryptedPackage stream header: set a non-zero high DWORD that should be treated
    // as "reserved" by compatibility parsers.
    let mut encrypted_package_stream = open_stream(&mut ole, "EncryptedPackage");
    let mut header_bytes = [0u8; 8];
    encrypted_package_stream
        .read_exact(&mut header_bytes)
        .expect("read EncryptedPackage header");
    // Preserve the low DWORD (actual size) and set the high DWORD to a non-zero value.
    header_bytes[4..8].copy_from_slice(&1u32.to_le_bytes());
    encrypted_package_stream
        .seek(std::io::SeekFrom::Start(0))
        .expect("seek EncryptedPackage to start");
    encrypted_package_stream
        .write_all(&header_bytes)
        .expect("write modified EncryptedPackage header");
    encrypted_package_stream
        .seek(std::io::SeekFrom::Start(0))
        .expect("seek EncryptedPackage to start");

    let mut out = Vec::new();
    let declared_len = decrypt_agile_encrypted_package_stream(
        &encryption_info,
        &mut encrypted_package_stream,
        password,
        &mut out,
    )
    .expect("decrypt agile encrypted package with reserved header hi dword");

    assert_eq!(declared_len as usize, plaintext.len());
    assert_eq!(out, plaintext);
}
