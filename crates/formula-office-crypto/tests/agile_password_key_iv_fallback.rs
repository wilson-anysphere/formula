use std::io::{Cursor, Read, Write};

use aes::{Aes128, Aes192, Aes256};
use base64::engine::general_purpose::STANDARD as BASE64;
use base64::Engine as _;
use cbc::{Decryptor, Encryptor};
use cipher::block_padding::NoPadding;
use cipher::{BlockDecryptMut, BlockEncryptMut, KeyIvInit};
use formula_office_crypto::{
    decrypt_encrypted_package_ole, encrypt_package_to_ole, EncryptOptions, EncryptionScheme,
    HashAlgorithm, OfficeCryptoError,
};
use quick_xml::events::Event;
use quick_xml::Reader;
use sha2::Digest as _;

const BLOCK_KEY_VERIFIER_HASH_INPUT: &[u8; 8] = b"\xFE\xA7\xD2\x76\x3B\x4B\x9E\x79";
const BLOCK_KEY_VERIFIER_HASH_VALUE: &[u8; 8] = b"\xD7\xAA\x0F\x6D\x30\x61\x34\x4E";
const BLOCK_KEY_ENCRYPTED_KEY_VALUE: &[u8; 8] = b"\x14\x6E\x0B\xE7\xAB\xAC\xD0\xD6";

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

fn local_name(name: &[u8]) -> &[u8] {
    match name.iter().position(|&b| b == b':') {
        Some(idx) => &name[idx + 1..],
        None => name,
    }
}

fn parse_hash_algorithm(name: &str) -> HashAlgorithm {
    match name {
        "MD5" | "MD-5" => HashAlgorithm::Md5,
        "SHA1" | "SHA-1" => HashAlgorithm::Sha1,
        "SHA256" | "SHA-256" => HashAlgorithm::Sha256,
        "SHA384" | "SHA-384" => HashAlgorithm::Sha384,
        "SHA512" | "SHA-512" => HashAlgorithm::Sha512,
        other => panic!("unsupported hash algorithm {other}"),
    }
}

fn password_to_utf16le(password: &str) -> Vec<u8> {
    let mut out = Vec::new();
    let _ = out.try_reserve(password.len().saturating_mul(2));
    for cu in password.encode_utf16() {
        out.extend_from_slice(&cu.to_le_bytes());
    }
    out
}

fn digest(hash_alg: HashAlgorithm, data: &[u8]) -> Vec<u8> {
    match hash_alg {
        HashAlgorithm::Md5 => md5::Md5::digest(data).to_vec(),
        HashAlgorithm::Sha1 => sha1::Sha1::digest(data).to_vec(),
        HashAlgorithm::Sha256 => sha2::Sha256::digest(data).to_vec(),
        HashAlgorithm::Sha384 => sha2::Sha384::digest(data).to_vec(),
        HashAlgorithm::Sha512 => sha2::Sha512::digest(data).to_vec(),
    }
}

fn truncate_hash(mut bytes: Vec<u8>, out_len: usize) -> Vec<u8> {
    if bytes.len() >= out_len {
        bytes.truncate(out_len);
        return bytes;
    }
    // MS-OFFCRYPTO TruncateHash padding behavior: extend with 0x36.
    bytes.resize(out_len, 0x36);
    bytes
}

fn hash_password(
    hash_alg: HashAlgorithm,
    salt: &[u8],
    password_utf16le: &[u8],
    spin_count: u32,
) -> Vec<u8> {
    let mut initial = Vec::new();
    let _ = initial.try_reserve_exact(salt.len().saturating_add(password_utf16le.len()));
    initial.extend_from_slice(salt);
    initial.extend_from_slice(password_utf16le);
    let mut h = digest(hash_alg, &initial);
    for i in 0..spin_count {
        let mut tmp = Vec::new();
        let _ = tmp.try_reserve_exact(4usize.saturating_add(h.len()));
        tmp.extend_from_slice(&i.to_le_bytes());
        tmp.extend_from_slice(&h);
        h = digest(hash_alg, &tmp);
    }
    h
}

fn derive_agile_key(
    hash_alg: HashAlgorithm,
    salt: &[u8],
    password_utf16le: &[u8],
    spin_count: u32,
    key_bytes: usize,
    block_key: &[u8],
) -> Vec<u8> {
    let h = hash_password(hash_alg, salt, password_utf16le, spin_count);
    let mut tmp = Vec::new();
    let _ = tmp.try_reserve_exact(h.len().saturating_add(block_key.len()));
    tmp.extend_from_slice(&h);
    tmp.extend_from_slice(block_key);
    truncate_hash(digest(hash_alg, &tmp), key_bytes)
}

fn derive_iv(hash_alg: HashAlgorithm, salt: &[u8], block_key: &[u8], iv_len: usize) -> Vec<u8> {
    let mut tmp = Vec::new();
    let _ = tmp.try_reserve_exact(salt.len().saturating_add(block_key.len()));
    tmp.extend_from_slice(salt);
    tmp.extend_from_slice(block_key);
    truncate_hash(digest(hash_alg, &tmp), iv_len)
}

fn aes_cbc_decrypt_no_padding(key: &[u8], iv: &[u8], ciphertext: &[u8]) -> Vec<u8> {
    assert_eq!(iv.len(), 16);
    assert_eq!(ciphertext.len() % 16, 0);
    let mut buf = ciphertext.to_vec();
    match key.len() {
        16 => {
            let dec = Decryptor::<Aes128>::new_from_slices(key, iv).expect("key/iv");
            dec.decrypt_padded_mut::<NoPadding>(&mut buf)
                .expect("decrypt");
        }
        24 => {
            let dec = Decryptor::<Aes192>::new_from_slices(key, iv).expect("key/iv");
            dec.decrypt_padded_mut::<NoPadding>(&mut buf)
                .expect("decrypt");
        }
        32 => {
            let dec = Decryptor::<Aes256>::new_from_slices(key, iv).expect("key/iv");
            dec.decrypt_padded_mut::<NoPadding>(&mut buf)
                .expect("decrypt");
        }
        other => panic!("unsupported AES key length {other}"),
    }
    buf
}

fn aes_cbc_encrypt_no_padding(key: &[u8], iv: &[u8], plaintext: &[u8]) -> Vec<u8> {
    assert_eq!(iv.len(), 16);
    assert_eq!(plaintext.len() % 16, 0);
    let mut buf = plaintext.to_vec();
    match key.len() {
        16 => {
            let enc = Encryptor::<Aes128>::new_from_slices(key, iv).expect("key/iv");
            enc.encrypt_padded_mut::<NoPadding>(&mut buf, plaintext.len())
                .expect("encrypt");
        }
        24 => {
            let enc = Encryptor::<Aes192>::new_from_slices(key, iv).expect("key/iv");
            enc.encrypt_padded_mut::<NoPadding>(&mut buf, plaintext.len())
                .expect("encrypt");
        }
        32 => {
            let enc = Encryptor::<Aes256>::new_from_slices(key, iv).expect("key/iv");
            enc.encrypt_padded_mut::<NoPadding>(&mut buf, plaintext.len())
                .expect("encrypt");
        }
        other => panic!("unsupported AES key length {other}"),
    }
    buf
}

#[derive(Debug)]
struct EncryptedKeyBlobs {
    salt_value: Vec<u8>,
    block_size: usize,
    key_bits: usize,
    spin_count: u32,
    hash_algorithm: HashAlgorithm,
    encrypted_verifier_hash_input: Vec<u8>,
    encrypted_verifier_hash_value: Vec<u8>,
    encrypted_key_value: Vec<u8>,
}

fn parse_password_encrypted_key(xml: &str) -> EncryptedKeyBlobs {
    let mut reader = Reader::from_str(xml);
    reader.trim_text(true);
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) | Ok(Event::Start(e)) => {
                let qname = e.name();
                let name = local_name(qname.as_ref());
                if name != b"encryptedKey" {
                    continue;
                }

                let mut salt_value = None;
                let mut block_size = None;
                let mut key_bits = None;
                let mut spin_count = None;
                let mut hash_algorithm = None;
                let mut enc_vhi = None;
                let mut enc_vhv = None;
                let mut enc_kv = None;

                for attr in e.attributes() {
                    let attr = attr.expect("attr");
                    let key = local_name(attr.key.as_ref());
                    let value = attr
                        .decode_and_unescape_value(&reader)
                        .expect("attr decode");
                    match key {
                        b"saltValue" => {
                            salt_value = Some(BASE64.decode(value.as_ref()).expect("saltValue b64"))
                        }
                        b"blockSize" => {
                            block_size = Some(value.parse::<usize>().expect("blockSize"))
                        }
                        b"keyBits" => key_bits = Some(value.parse::<usize>().expect("keyBits")),
                        b"spinCount" => spin_count = Some(value.parse::<u32>().expect("spinCount")),
                        b"hashAlgorithm" => {
                            hash_algorithm = Some(parse_hash_algorithm(value.as_ref()))
                        }
                        b"encryptedVerifierHashInput" => {
                            enc_vhi = Some(
                                BASE64
                                    .decode(value.as_ref())
                                    .expect("encryptedVerifierHashInput b64"),
                            )
                        }
                        b"encryptedVerifierHashValue" => {
                            enc_vhv = Some(
                                BASE64
                                    .decode(value.as_ref())
                                    .expect("encryptedVerifierHashValue b64"),
                            )
                        }
                        b"encryptedKeyValue" => {
                            enc_kv = Some(
                                BASE64
                                    .decode(value.as_ref())
                                    .expect("encryptedKeyValue b64"),
                            )
                        }
                        _ => {}
                    }
                }

                return EncryptedKeyBlobs {
                    salt_value: salt_value.expect("saltValue"),
                    block_size: block_size.expect("blockSize"),
                    key_bits: key_bits.expect("keyBits"),
                    spin_count: spin_count.expect("spinCount"),
                    hash_algorithm: hash_algorithm.expect("hashAlgorithm"),
                    encrypted_verifier_hash_input: enc_vhi.expect("encryptedVerifierHashInput"),
                    encrypted_verifier_hash_value: enc_vhv.expect("encryptedVerifierHashValue"),
                    encrypted_key_value: enc_kv.expect("encryptedKeyValue"),
                };
            }
            Ok(Event::Eof) => break,
            Err(e) => panic!("xml parse failed: {e}"),
            _ => {}
        }
        buf.clear();
    }

    panic!("missing encryptedKey element");
}

fn replace_xml_attr(xml: &str, attr: &str, new_value: &str) -> String {
    let needle = format!(r#"{attr}=""#);
    let start = xml
        .find(&needle)
        .unwrap_or_else(|| panic!("missing attribute {attr}"));
    let value_start = start + needle.len();
    let end_rel = xml[value_start..]
        .find('"')
        .unwrap_or_else(|| panic!("unterminated attribute {attr}"));
    let value_end = value_start + end_rel;

    let mut out = String::new();
    let _ = out.try_reserve(xml.len().saturating_sub(value_end - value_start).saturating_add(new_value.len()));
    out.push_str(&xml[..value_start]);
    out.push_str(new_value);
    out.push_str(&xml[value_end..]);
    out
}

#[test]
fn decrypt_agile_falls_back_to_derived_iv_for_password_key_encryptor_blobs() {
    // This builds a valid Agile-encrypted OLE wrapper using the crate's writer, then patches the
    // three password key-encryptor ciphertext blobs to use a per-blob derived IV
    // (Hash(saltValue || blockKey)[:16]) instead of the Excel scheme (IV=saltValue[:16]).
    //
    // The decryptor should fall back to the derived-IV variant only when the verifier check fails.
    let plaintext = minimal_zip_bytes();
    let password = "Password";

    // Keep spinCount small for test speed.
    let opts = EncryptOptions {
        scheme: EncryptionScheme::Agile,
        key_bits: 256,
        hash_algorithm: HashAlgorithm::Sha512,
        spin_count: 512,
    };
    let baseline_ole = encrypt_package_to_ole(&plaintext, password, opts).expect("encrypt");
    let (encryption_info, encrypted_package) = extract_streams_from_ole(&baseline_ole);

    let header = encryption_info
        .get(..8)
        .expect("EncryptionInfo must include version header");
    let xml = std::str::from_utf8(&encryption_info[8..]).expect("EncryptionInfo XML must be UTF-8");
    let encrypted_key = parse_password_encrypted_key(xml);

    assert_eq!(encrypted_key.block_size, 16);
    assert_eq!(encrypted_key.key_bits, 256);
    assert_eq!(encrypted_key.hash_algorithm, HashAlgorithm::Sha512);

    let pw_utf16 = password_to_utf16le(password);
    let key_bytes = encrypted_key.key_bits / 8;
    let verifier_iv_salt = &encrypted_key.salt_value[..encrypted_key.block_size];

    // Decrypt the existing (salt-IV) ciphertext blobs to recover the plaintexts.
    let key_vhi = derive_agile_key(
        encrypted_key.hash_algorithm,
        &encrypted_key.salt_value,
        &pw_utf16,
        encrypted_key.spin_count,
        key_bytes,
        BLOCK_KEY_VERIFIER_HASH_INPUT,
    );
    let plain_vhi = aes_cbc_decrypt_no_padding(
        &key_vhi,
        verifier_iv_salt,
        &encrypted_key.encrypted_verifier_hash_input,
    );

    let key_vhv = derive_agile_key(
        encrypted_key.hash_algorithm,
        &encrypted_key.salt_value,
        &pw_utf16,
        encrypted_key.spin_count,
        key_bytes,
        BLOCK_KEY_VERIFIER_HASH_VALUE,
    );
    let plain_vhv = aes_cbc_decrypt_no_padding(
        &key_vhv,
        verifier_iv_salt,
        &encrypted_key.encrypted_verifier_hash_value,
    );

    let key_kv = derive_agile_key(
        encrypted_key.hash_algorithm,
        &encrypted_key.salt_value,
        &pw_utf16,
        encrypted_key.spin_count,
        key_bytes,
        BLOCK_KEY_ENCRYPTED_KEY_VALUE,
    );
    let plain_kv = aes_cbc_decrypt_no_padding(
        &key_kv,
        verifier_iv_salt,
        &encrypted_key.encrypted_key_value,
    );

    // Re-encrypt using derived IVs.
    let iv_vhi_derived = derive_iv(
        encrypted_key.hash_algorithm,
        &encrypted_key.salt_value,
        BLOCK_KEY_VERIFIER_HASH_INPUT,
        encrypted_key.block_size,
    );
    let enc_vhi_derived = aes_cbc_encrypt_no_padding(&key_vhi, &iv_vhi_derived, &plain_vhi);

    let iv_vhv_derived = derive_iv(
        encrypted_key.hash_algorithm,
        &encrypted_key.salt_value,
        BLOCK_KEY_VERIFIER_HASH_VALUE,
        encrypted_key.block_size,
    );
    let enc_vhv_derived = aes_cbc_encrypt_no_padding(&key_vhv, &iv_vhv_derived, &plain_vhv);

    let iv_kv_derived = derive_iv(
        encrypted_key.hash_algorithm,
        &encrypted_key.salt_value,
        BLOCK_KEY_ENCRYPTED_KEY_VALUE,
        encrypted_key.block_size,
    );
    let enc_kv_derived = aes_cbc_encrypt_no_padding(&key_kv, &iv_kv_derived, &plain_kv);

    // Patch EncryptionInfo XML to use the derived-IV ciphertext blobs.
    let xml = replace_xml_attr(
        xml,
        "encryptedVerifierHashInput",
        &BASE64.encode(enc_vhi_derived),
    );
    let xml = replace_xml_attr(
        &xml,
        "encryptedVerifierHashValue",
        &BASE64.encode(enc_vhv_derived),
    );
    let xml = replace_xml_attr(&xml, "encryptedKeyValue", &BASE64.encode(enc_kv_derived));

    let mut patched_info = Vec::new();
    patched_info.extend_from_slice(header);
    patched_info.extend_from_slice(xml.as_bytes());

    let ole_bytes = build_ole(&patched_info, &encrypted_package);

    // Correct password should decrypt successfully via derived-IV fallback.
    let decrypted =
        decrypt_encrypted_package_ole(&ole_bytes, password).expect("decrypt derived-IV");
    assert_eq!(decrypted, plaintext);

    // Wrong password should still report InvalidPassword (not an integrity or format error).
    let err = decrypt_encrypted_package_ole(&ole_bytes, "wrong").expect_err("wrong password");
    assert!(
        matches!(err, OfficeCryptoError::InvalidPassword),
        "unexpected error: {err:?}"
    );
}
