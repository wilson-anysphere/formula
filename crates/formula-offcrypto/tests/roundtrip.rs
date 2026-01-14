use std::io::{Cursor, Read, Write};

use aes::{Aes128, Aes192, Aes256};
use cbc::Decryptor;
use cipher::block_padding::NoPadding;
use cipher::{BlockDecryptMut, KeyIvInit};
use sha1::{Digest as _, Sha1};

use formula_offcrypto::{
    agile_decrypt_package, agile_secret_key, agile_verify_password, decrypt_encrypted_package,
    parse_encryption_info, standard_derive_key, standard_verify_key, EncryptionInfo, HashAlgorithm,
    OffcryptoError, StandardEncryptionInfo,
};

mod support;

fn build_test_zip() -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("[Content_Types].xml", options)
        .expect("start [Content_Types].xml");
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>"#)
        .expect("write [Content_Types].xml");

    zip.start_file("xl/workbook.xml", options)
        .expect("start xl/workbook.xml");
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#)
        .expect("write xl/workbook.xml");

    zip.finish().expect("finish zip").into_inner()
}

fn assert_zip_contains_workbook_xml(bytes: &[u8]) {
    let cursor = Cursor::new(bytes);
    let zip = zip::ZipArchive::new(cursor).expect("zip archive");
    let mut found = false;
    for name in zip.file_names() {
        if name.eq_ignore_ascii_case("xl/workbook.xml") {
            found = true;
            break;
        }
    }
    assert!(found, "zip should contain xl/workbook.xml");
}

fn read_ole_stream(ole_bytes: &[u8], name: &str) -> Vec<u8> {
    let cursor = Cursor::new(ole_bytes);
    let mut ole = cfb::CompoundFile::open(cursor).expect("open ole");
    let mut out = Vec::new();
    ole.open_stream(name)
        .expect("open stream")
        .read_to_end(&mut out)
        .expect("read stream");
    out
}

fn hash_digest(algo: HashAlgorithm, data: &[u8]) -> Vec<u8> {
    match algo {
        HashAlgorithm::Sha1 => Sha1::digest(data).to_vec(),
        HashAlgorithm::Sha256 => sha2::Sha256::digest(data).to_vec(),
        HashAlgorithm::Sha384 => sha2::Sha384::digest(data).to_vec(),
        HashAlgorithm::Sha512 => sha2::Sha512::digest(data).to_vec(),
    }
}

fn derive_iv(algo: HashAlgorithm, salt: &[u8], block_index: u32) -> [u8; 16] {
    let mut buf = Vec::with_capacity(salt.len() + 4);
    buf.extend_from_slice(salt);
    buf.extend_from_slice(&block_index.to_le_bytes());
    let digest = hash_digest(algo, &buf);
    let mut iv = [0u8; 16];
    iv.copy_from_slice(&digest[..16]);
    iv
}

fn aes_cbc_decrypt_in_place(key: &[u8], iv: &[u8; 16], buf: &mut [u8]) -> Result<(), OffcryptoError> {
    if buf.len() % 16 != 0 {
        return Err(OffcryptoError::InvalidCiphertextLength { len: buf.len() });
    }

    match key.len() {
        16 => {
            let decryptor =
                Decryptor::<Aes128>::new_from_slices(key, iv).map_err(|_| OffcryptoError::InvalidKeyLength {
                    len: key.len(),
                })?;
            decryptor
                .decrypt_padded_mut::<NoPadding>(buf)
                .map_err(|_| OffcryptoError::InvalidEncryptionInfo {
                    context: "AES-CBC decrypt failed",
                })?;
        }
        24 => {
            let decryptor =
                Decryptor::<Aes192>::new_from_slices(key, iv).map_err(|_| OffcryptoError::InvalidKeyLength {
                    len: key.len(),
                })?;
            decryptor
                .decrypt_padded_mut::<NoPadding>(buf)
                .map_err(|_| OffcryptoError::InvalidEncryptionInfo {
                    context: "AES-CBC decrypt failed",
                })?;
        }
        32 => {
            let decryptor =
                Decryptor::<Aes256>::new_from_slices(key, iv).map_err(|_| OffcryptoError::InvalidKeyLength {
                    len: key.len(),
                })?;
            decryptor
                .decrypt_padded_mut::<NoPadding>(buf)
                .map_err(|_| OffcryptoError::InvalidEncryptionInfo {
                    context: "AES-CBC decrypt failed",
                })?;
        }
        _ => return Err(OffcryptoError::InvalidKeyLength { len: key.len() }),
    }

    Ok(())
}

#[test]
fn roundtrip_standard_encryption() {
    let password = "Password";
    let plaintext = build_test_zip();

    let (encryption_info, encrypted_package) = support::encrypt_standard(&plaintext, password);
    let ole_bytes = support::wrap_in_ole_cfb(&encryption_info, &encrypted_package);

    let encryption_info = read_ole_stream(&ole_bytes, "EncryptionInfo");
    let encrypted_package = read_ole_stream(&ole_bytes, "EncryptedPackage");

    let parsed = parse_encryption_info(&encryption_info).expect("parse EncryptionInfo");
    let EncryptionInfo::Standard { header, verifier, .. } = parsed else {
        panic!("expected Standard EncryptionInfo");
    };
    let info = StandardEncryptionInfo { header, verifier };

    let key = standard_derive_key(&info, password).expect("derive key");
    standard_verify_key(&info, &key).expect("verify key");

    let salt = info.verifier.salt.clone();
    let decrypted = decrypt_encrypted_package(&encrypted_package, |block, ct, pt| {
        pt.copy_from_slice(ct);
        let iv = derive_iv(HashAlgorithm::Sha1, &salt, block);
        aes_cbc_decrypt_in_place(&key, &iv, pt)?;
        Ok(())
    })
    .expect("decrypt EncryptedPackage");

    assert_eq!(decrypted, plaintext);
    assert_zip_contains_workbook_xml(&decrypted);
}

#[test]
fn roundtrip_agile_encryption() {
    let password = "Password";
    let plaintext = build_test_zip();

    let (encryption_info, encrypted_package) = support::encrypt_agile(&plaintext, password);
    let ole_bytes = support::wrap_in_ole_cfb(&encryption_info, &encrypted_package);

    let encryption_info = read_ole_stream(&ole_bytes, "EncryptionInfo");
    let encrypted_package = read_ole_stream(&ole_bytes, "EncryptedPackage");

    let parsed = parse_encryption_info(&encryption_info).expect("parse EncryptionInfo");
    let EncryptionInfo::Agile { info, .. } = parsed else {
        panic!("expected Agile EncryptionInfo");
    };

    agile_verify_password(&info, password).expect("verify password");
    let secret_key = agile_secret_key(&info, password).expect("derive agile secret key");

    assert_eq!(
        info.key_data_block_size, 16,
        "expected test helper to use 16-byte block size"
    );

    let decrypted = agile_decrypt_package(&info, &secret_key, &encrypted_package)
        .expect("decrypt EncryptedPackage");

    assert_eq!(decrypted, plaintext);
    assert_zip_contains_workbook_xml(&decrypted);
}
