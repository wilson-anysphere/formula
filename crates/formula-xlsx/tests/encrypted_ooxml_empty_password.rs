use std::io::Read as _;
use std::path::{Path, PathBuf};

use base64::Engine as _;
use formula_xlsx::offcrypto::{
    derive_iv, derive_key, hash_password, HashAlgorithm, HMAC_KEY_BLOCK, HMAC_VALUE_BLOCK,
    KEY_VALUE_BLOCK, VERIFIER_HASH_INPUT_BLOCK, VERIFIER_HASH_VALUE_BLOCK,
};

#[derive(Debug, PartialEq, Eq)]
enum DecryptError {
    PasswordRequired,
    WrongPassword,
    InvalidFixture(&'static str),
}

fn fixture_path(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/encrypted/ooxml")
        .join(name)
}

fn open_cfb_stream<R: std::io::Read + std::io::Write + std::io::Seek>(
    ole: &mut cfb::CompoundFile<R>,
    name: &str,
) -> cfb::Stream<R> {
    ole.open_stream(name).unwrap_or_else(|_| {
        let with_slash = format!("/{name}");
        ole.open_stream(&with_slash).expect("open stream")
    })
}

fn aes_cbc_decrypt_no_pad(
    key: &[u8],
    iv: &[u8],
    ciphertext: &[u8],
) -> Result<Vec<u8>, DecryptError> {
    use openssl::symm::{Cipher, Crypter, Mode};

    let cipher = match key.len() {
        16 => Cipher::aes_128_cbc(),
        24 => Cipher::aes_192_cbc(),
        32 => Cipher::aes_256_cbc(),
        _ => return Err(DecryptError::InvalidFixture("unsupported AES key length")),
    };

    if iv.len() != cipher.iv_len().unwrap_or(16) {
        return Err(DecryptError::InvalidFixture("invalid IV length"));
    }

    let mut c = Crypter::new(cipher, Mode::Decrypt, key, Some(iv))
        .map_err(|_| DecryptError::InvalidFixture("failed to initialize AES decryptor"))?;
    c.pad(false);

    let mut out = vec![0u8; ciphertext.len() + cipher.block_size()];
    let mut count = c
        .update(ciphertext, &mut out)
        .map_err(|_| DecryptError::InvalidFixture("AES decrypt update failed"))?;
    count += c
        .finalize(&mut out[count..])
        .map_err(|_| DecryptError::InvalidFixture("AES decrypt finalize failed"))?;
    out.truncate(count);
    Ok(out)
}

fn digest(hash_alg: HashAlgorithm, data: &[u8]) -> Vec<u8> {
    use digest::Digest as _;

    match hash_alg {
        HashAlgorithm::Sha1 => {
            let mut h = sha1::Sha1::new();
            h.update(data);
            h.finalize().to_vec()
        }
        HashAlgorithm::Sha256 => {
            let mut h = sha2::Sha256::new();
            h.update(data);
            h.finalize().to_vec()
        }
        HashAlgorithm::Sha384 => {
            let mut h = sha2::Sha384::new();
            h.update(data);
            h.finalize().to_vec()
        }
        HashAlgorithm::Sha512 => {
            let mut h = sha2::Sha512::new();
            h.update(data);
            h.finalize().to_vec()
        }
    }
}

fn hmac(hash_alg: HashAlgorithm, key: &[u8], data: &[u8]) -> Result<Vec<u8>, DecryptError> {
    use openssl::hash::MessageDigest;
    use openssl::pkey::PKey;
    use openssl::sign::Signer;

    let md = match hash_alg {
        HashAlgorithm::Sha1 => MessageDigest::sha1(),
        HashAlgorithm::Sha256 => MessageDigest::sha256(),
        HashAlgorithm::Sha384 => MessageDigest::sha384(),
        HashAlgorithm::Sha512 => MessageDigest::sha512(),
    };

    let pkey =
        PKey::hmac(key).map_err(|_| DecryptError::InvalidFixture("failed to init HMAC key"))?;
    let mut signer =
        Signer::new(md, &pkey).map_err(|_| DecryptError::InvalidFixture("failed to init HMAC"))?;
    signer
        .update(data)
        .map_err(|_| DecryptError::InvalidFixture("failed to compute HMAC"))?;
    signer
        .sign_to_vec()
        .map_err(|_| DecryptError::InvalidFixture("failed to compute HMAC"))
}

fn ct_eq(a: &[u8], b: &[u8]) -> bool {
    let mut diff = 0u8;
    let max_len = a.len().max(b.len());
    for idx in 0..max_len {
        let av = a.get(idx).copied().unwrap_or(0);
        let bv = b.get(idx).copied().unwrap_or(0);
        diff |= av ^ bv;
    }
    diff == 0 && a.len() == b.len()
}

#[derive(Debug)]
struct AgileEncryptionInfo {
    key_data_salt: Vec<u8>,
    key_data_block_size: usize,
    key_data_key_bits: usize,
    key_data_hash_size: usize,
    key_data_hash_alg: HashAlgorithm,

    encrypted_hmac_key: Vec<u8>,
    encrypted_hmac_value: Vec<u8>,

    spin_count: u32,
    encrypted_key_salt: Vec<u8>,
    encrypted_key_hash_alg: HashAlgorithm,

    encrypted_verifier_hash_input: Vec<u8>,
    encrypted_verifier_hash_value: Vec<u8>,
    encrypted_key_value: Vec<u8>,
}

fn parse_agile_encryption_info(bytes: &[u8]) -> Result<AgileEncryptionInfo, DecryptError> {
    if bytes.len() < 8 {
        return Err(DecryptError::InvalidFixture("EncryptionInfo too short"));
    }
    let major = u16::from_le_bytes([bytes[0], bytes[1]]);
    let minor = u16::from_le_bytes([bytes[2], bytes[3]]);
    if (major, minor) != (4, 4) {
        return Err(DecryptError::InvalidFixture(
            "expected Agile EncryptionInfo version 4.4",
        ));
    }

    let xml = std::str::from_utf8(&bytes[8..])
        .map_err(|_| DecryptError::InvalidFixture("EncryptionInfo XML is not UTF-8"))?;
    let doc =
        roxmltree::Document::parse(xml).map_err(|_| DecryptError::InvalidFixture("invalid XML"))?;

    const ENC_NS: &str = "http://schemas.microsoft.com/office/2006/encryption";
    const PW_NS: &str = "http://schemas.microsoft.com/office/2006/keyEncryptor/password";
    let b64 = base64::engine::general_purpose::STANDARD;

    let key_data = doc
        .descendants()
        .find(|n| {
            n.is_element()
                && n.tag_name().name() == "keyData"
                && n.tag_name().namespace() == Some(ENC_NS)
        })
        .ok_or(DecryptError::InvalidFixture("missing keyData element"))?;

    let key_data_salt = b64
        .decode(
            key_data
                .attribute("saltValue")
                .ok_or(DecryptError::InvalidFixture("missing keyData@saltValue"))?,
        )
        .map_err(|_| DecryptError::InvalidFixture("invalid base64 in keyData@saltValue"))?;

    let key_data_block_size: usize = key_data
        .attribute("blockSize")
        .ok_or(DecryptError::InvalidFixture("missing keyData@blockSize"))?
        .parse()
        .map_err(|_| DecryptError::InvalidFixture("invalid keyData@blockSize"))?;
    let key_data_key_bits: usize = key_data
        .attribute("keyBits")
        .ok_or(DecryptError::InvalidFixture("missing keyData@keyBits"))?
        .parse()
        .map_err(|_| DecryptError::InvalidFixture("invalid keyData@keyBits"))?;
    let key_data_hash_size: usize = key_data
        .attribute("hashSize")
        .ok_or(DecryptError::InvalidFixture("missing keyData@hashSize"))?
        .parse()
        .map_err(|_| DecryptError::InvalidFixture("invalid keyData@hashSize"))?;
    let key_data_hash_alg =
        HashAlgorithm::parse_offcrypto_name(key_data.attribute("hashAlgorithm").ok_or(
            DecryptError::InvalidFixture("missing keyData@hashAlgorithm"),
        )?)
        .map_err(|_| DecryptError::InvalidFixture("unsupported keyData@hashAlgorithm"))?;

    let data_integrity = doc
        .descendants()
        .find(|n| {
            n.is_element()
                && n.tag_name().name() == "dataIntegrity"
                && n.tag_name().namespace() == Some(ENC_NS)
        })
        .ok_or(DecryptError::InvalidFixture(
            "missing dataIntegrity element",
        ))?;
    let encrypted_hmac_key = b64
        .decode(data_integrity.attribute("encryptedHmacKey").ok_or(
            DecryptError::InvalidFixture("missing dataIntegrity@encryptedHmacKey"),
        )?)
        .map_err(|_| DecryptError::InvalidFixture("invalid base64 in encryptedHmacKey"))?;
    let encrypted_hmac_value = b64
        .decode(data_integrity.attribute("encryptedHmacValue").ok_or(
            DecryptError::InvalidFixture("missing dataIntegrity@encryptedHmacValue"),
        )?)
        .map_err(|_| DecryptError::InvalidFixture("invalid base64 in encryptedHmacValue"))?;

    let encrypted_key = doc
        .descendants()
        .find(|n| {
            n.is_element()
                && n.tag_name().name() == "encryptedKey"
                && n.tag_name().namespace() == Some(PW_NS)
        })
        .ok_or(DecryptError::InvalidFixture(
            "missing p:encryptedKey element",
        ))?;
    let spin_count: u32 = encrypted_key
        .attribute("spinCount")
        .ok_or(DecryptError::InvalidFixture(
            "missing encryptedKey@spinCount",
        ))?
        .parse()
        .map_err(|_| DecryptError::InvalidFixture("invalid encryptedKey@spinCount"))?;
    let encrypted_key_salt = b64
        .decode(
            encrypted_key
                .attribute("saltValue")
                .ok_or(DecryptError::InvalidFixture(
                    "missing encryptedKey@saltValue",
                ))?,
        )
        .map_err(|_| DecryptError::InvalidFixture("invalid base64 in encryptedKey@saltValue"))?;
    let encrypted_key_hash_alg =
        HashAlgorithm::parse_offcrypto_name(encrypted_key.attribute("hashAlgorithm").ok_or(
            DecryptError::InvalidFixture("missing encryptedKey@hashAlgorithm"),
        )?)
        .map_err(|_| DecryptError::InvalidFixture("unsupported encryptedKey@hashAlgorithm"))?;

    let encrypted_verifier_hash_input = b64
        .decode(
            encrypted_key
                .attribute("encryptedVerifierHashInput")
                .ok_or(DecryptError::InvalidFixture(
                    "missing encryptedKey@encryptedVerifierHashInput",
                ))?,
        )
        .map_err(|_| DecryptError::InvalidFixture("invalid base64 verifierHashInput"))?;
    let encrypted_verifier_hash_value = b64
        .decode(
            encrypted_key
                .attribute("encryptedVerifierHashValue")
                .ok_or(DecryptError::InvalidFixture(
                    "missing encryptedKey@encryptedVerifierHashValue",
                ))?,
        )
        .map_err(|_| DecryptError::InvalidFixture("invalid base64 verifierHashValue"))?;
    let encrypted_key_value = b64
        .decode(encrypted_key.attribute("encryptedKeyValue").ok_or(
            DecryptError::InvalidFixture("missing encryptedKey@encryptedKeyValue"),
        )?)
        .map_err(|_| DecryptError::InvalidFixture("invalid base64 encryptedKeyValue"))?;

    Ok(AgileEncryptionInfo {
        key_data_salt,
        key_data_block_size,
        key_data_key_bits,
        key_data_hash_size,
        key_data_hash_alg,
        encrypted_hmac_key,
        encrypted_hmac_value,
        spin_count,
        encrypted_key_salt,
        encrypted_key_hash_alg,
        encrypted_verifier_hash_input,
        encrypted_verifier_hash_value,
        encrypted_key_value,
    })
}

fn decrypt_agile_ooxml(path: &Path, password: Option<&str>) -> Result<Vec<u8>, DecryptError> {
    let password = password.ok_or(DecryptError::PasswordRequired)?;

    let file = std::fs::File::open(path).expect("open encrypted fixture");
    let mut ole = cfb::CompoundFile::open(file).expect("open CFB");

    let mut encryption_info = Vec::new();
    open_cfb_stream(&mut ole, "EncryptionInfo")
        .read_to_end(&mut encryption_info)
        .expect("read EncryptionInfo");
    let info = parse_agile_encryption_info(&encryption_info)?;

    if info.key_data_hash_alg != info.encrypted_key_hash_alg {
        return Err(DecryptError::InvalidFixture(
            "fixture uses different hashAlgorithm for keyData vs encryptedKey",
        ));
    }
    let hash_alg = info.key_data_hash_alg;

    let mut encrypted_package = Vec::new();
    open_cfb_stream(&mut ole, "EncryptedPackage")
        .read_to_end(&mut encrypted_package)
        .expect("read EncryptedPackage");
    if encrypted_package.len() < 8 {
        return Err(DecryptError::InvalidFixture(
            "EncryptedPackage stream is too short",
        ));
    }

    let total_size =
        u64::from_le_bytes(encrypted_package[0..8].try_into().expect("slice length")) as usize;
    let ciphertext = &encrypted_package[8..];

    // 1) Derive password hash.
    let pw_hash = hash_password(
        password,
        &info.encrypted_key_salt,
        info.spin_count,
        hash_alg,
    )
    .map_err(|_| DecryptError::InvalidFixture("failed to hash password"))?;

    let key_len = info.key_data_key_bits / 8;

    // 2) Decrypt verifierHashInput.
    let verifier_key = derive_key(&pw_hash, &VERIFIER_HASH_INPUT_BLOCK, key_len, hash_alg)
        .map_err(|_| DecryptError::InvalidFixture("failed to derive verifierHashInput key"))?;
    let verifier_iv = derive_iv(
        &info.encrypted_key_salt,
        &VERIFIER_HASH_INPUT_BLOCK,
        info.key_data_block_size,
        hash_alg,
    )
    .map_err(|_| DecryptError::InvalidFixture("failed to derive verifierHashInput iv"))?;
    let verifier_hash_input = aes_cbc_decrypt_no_pad(
        &verifier_key,
        &verifier_iv,
        &info.encrypted_verifier_hash_input,
    )?;

    // 3) Decrypt verifierHashValue and verify password.
    let verifier_value_key = derive_key(&pw_hash, &VERIFIER_HASH_VALUE_BLOCK, key_len, hash_alg)
        .map_err(|_| DecryptError::InvalidFixture("failed to derive verifierHashValue key"))?;
    let verifier_value_iv = derive_iv(
        &info.encrypted_key_salt,
        &VERIFIER_HASH_VALUE_BLOCK,
        info.key_data_block_size,
        hash_alg,
    )
    .map_err(|_| DecryptError::InvalidFixture("failed to derive verifierHashValue iv"))?;
    let verifier_hash_value = aes_cbc_decrypt_no_pad(
        &verifier_value_key,
        &verifier_value_iv,
        &info.encrypted_verifier_hash_value,
    )?;

    let expected = digest(hash_alg, &verifier_hash_input);
    let expected = &expected[..info.key_data_hash_size.min(expected.len())];
    let actual = &verifier_hash_value[..info.key_data_hash_size.min(verifier_hash_value.len())];
    if !ct_eq(expected, actual) {
        return Err(DecryptError::WrongPassword);
    }

    // 4) Decrypt keyValue (package key).
    let key_value_key = derive_key(&pw_hash, &KEY_VALUE_BLOCK, key_len, hash_alg)
        .map_err(|_| DecryptError::InvalidFixture("failed to derive keyValue key"))?;
    let key_value_iv = derive_iv(
        &info.encrypted_key_salt,
        &KEY_VALUE_BLOCK,
        info.key_data_block_size,
        hash_alg,
    )
    .map_err(|_| DecryptError::InvalidFixture("failed to derive keyValue iv"))?;
    let key_value_raw =
        aes_cbc_decrypt_no_pad(&key_value_key, &key_value_iv, &info.encrypted_key_value)?;
    let key_value = key_value_raw
        .get(..key_len)
        .ok_or(DecryptError::InvalidFixture("keyValue out of bounds"))?;

    // 5) Decrypt HMAC key + value.
    let hmac_key_iv = derive_iv(
        &info.key_data_salt,
        &HMAC_KEY_BLOCK,
        info.key_data_block_size,
        hash_alg,
    )
    .map_err(|_| DecryptError::InvalidFixture("failed to derive hmacKey iv"))?;
    let hmac_key_raw = aes_cbc_decrypt_no_pad(key_value, &hmac_key_iv, &info.encrypted_hmac_key)?;
    let hmac_key = hmac_key_raw
        .get(..info.key_data_hash_size)
        .ok_or(DecryptError::InvalidFixture("hmacKey out of bounds"))?;

    let hmac_value_iv = derive_iv(
        &info.key_data_salt,
        &HMAC_VALUE_BLOCK,
        info.key_data_block_size,
        hash_alg,
    )
    .map_err(|_| DecryptError::InvalidFixture("failed to derive hmacValue iv"))?;
    let hmac_value_raw =
        aes_cbc_decrypt_no_pad(key_value, &hmac_value_iv, &info.encrypted_hmac_value)?;
    let hmac_value = hmac_value_raw
        .get(..info.key_data_hash_size)
        .ok_or(DecryptError::InvalidFixture("hmacValue out of bounds"))?;

    // 6) Decrypt package segments (4096-byte segments with per-segment IV).
    const SEGMENT_SIZE: usize = 4096;
    if ciphertext.len() % SEGMENT_SIZE != 0 {
        return Err(DecryptError::InvalidFixture(
            "ciphertext length is not a multiple of 4096",
        ));
    }

    let mut decrypted = Vec::with_capacity(ciphertext.len());
    for (i, segment) in ciphertext.chunks(SEGMENT_SIZE).enumerate() {
        let iv = derive_iv(
            &info.key_data_salt,
            &(i as u32).to_le_bytes(),
            info.key_data_block_size,
            hash_alg,
        )
        .map_err(|_| DecryptError::InvalidFixture("failed to derive segment IV"))?;
        let pt = aes_cbc_decrypt_no_pad(key_value, &iv, segment)?;
        decrypted.extend_from_slice(&pt);
    }
    if total_size > decrypted.len() {
        return Err(DecryptError::InvalidFixture(
            "EncryptedPackage total size exceeds decrypted length",
        ));
    }
    decrypted.truncate(total_size);

    // 7) Verify HMAC (computed over decrypted plaintext bytes).
    let expected_hmac = hmac(hash_alg, hmac_key, &decrypted)?;
    let expected_hmac = &expected_hmac[..info.key_data_hash_size.min(expected_hmac.len())];
    if !ct_eq(expected_hmac, hmac_value) {
        return Err(DecryptError::InvalidFixture("HMAC mismatch"));
    }

    Ok(decrypted)
}

#[test]
fn decrypts_agile_fixture_with_empty_password() {
    let encrypted_path = fixture_path("agile-empty-password.xlsx");
    let plaintext_path = fixture_path("plaintext.xlsx");

    let expected = std::fs::read(&plaintext_path).expect("read plaintext fixture");
    let decrypted = decrypt_agile_ooxml(&encrypted_path, Some("")).expect("decrypt");

    assert_eq!(decrypted, expected);
}

#[test]
fn empty_password_is_distinct_from_missing_password() {
    let encrypted_path = fixture_path("agile-empty-password.xlsx");

    let err = decrypt_agile_ooxml(&encrypted_path, None).unwrap_err();
    assert_eq!(err, DecryptError::PasswordRequired);
}

#[test]
fn wrong_password_fails() {
    let encrypted_path = fixture_path("agile-empty-password.xlsx");

    let err = decrypt_agile_ooxml(&encrypted_path, Some("wrong-password")).unwrap_err();
    assert_eq!(err, DecryptError::WrongPassword);
}
