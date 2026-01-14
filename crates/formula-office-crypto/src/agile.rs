use base64::engine::general_purpose::{STANDARD as BASE64_STANDARD, STANDARD_NO_PAD as BASE64_STANDARD_NO_PAD};
use base64::Engine;
use hmac::{Hmac, Mac};
use quick_xml::events::Event;
use quick_xml::Reader;
use rand::rngs::OsRng;
use rand::RngCore;
use sha1::Sha1;
use sha2::{Sha256, Sha384, Sha512};

use crate::crypto::{
    aes_cbc_decrypt, aes_cbc_encrypt, derive_agile_key, derive_iv, password_to_utf16le,
    HashAlgorithm,
};
use crate::error::OfficeCryptoError;
use crate::util::{checked_vec_len, ct_eq, read_u64_le, EncryptionInfoHeader};
use zeroize::Zeroizing;

const BLOCK_KEY_VERIFIER_HASH_INPUT: &[u8; 8] = b"\xFE\xA7\xD2\x76\x3B\x4B\x9E\x79";
const BLOCK_KEY_VERIFIER_HASH_VALUE: &[u8; 8] = b"\xD7\xAA\x0F\x6D\x30\x61\x34\x4E";
const BLOCK_KEY_ENCRYPTED_KEY_VALUE: &[u8; 8] = b"\x14\x6E\x0B\xE7\xAB\xAC\xD0\xD6";
const BLOCK_KEY_INTEGRITY_HMAC_KEY: &[u8; 8] = b"\x5F\xB2\xAD\x01\x0C\xB9\xE1\xF6";
const BLOCK_KEY_INTEGRITY_HMAC_VALUE: &[u8; 8] = b"\xA0\x67\x7F\x02\xB2\x2C\x84\x33";

#[derive(Debug, Clone)]
pub(crate) struct AgileEncryptionInfo {
    #[allow(dead_code)]
    pub(crate) version_major: u16,
    #[allow(dead_code)]
    pub(crate) version_minor: u16,
    #[allow(dead_code)]
    pub(crate) flags: u32,
    pub(crate) key_data: AgileKeyData,
    #[allow(dead_code)]
    pub(crate) data_integrity: AgileDataIntegrity,
    pub(crate) password_key_encryptor: AgilePasswordKeyEncryptor,
}

#[derive(Debug, Clone)]
pub(crate) struct AgileKeyData {
    pub(crate) salt: Vec<u8>,
    pub(crate) block_size: usize,
    pub(crate) key_bits: usize,
    pub(crate) hash_algorithm: HashAlgorithm,
    pub(crate) cipher_algorithm: String,
    pub(crate) cipher_chaining: String,
}

#[derive(Debug, Clone)]
pub(crate) struct AgileDataIntegrity {
    #[allow(dead_code)]
    pub(crate) encrypted_hmac_key: Vec<u8>,
    #[allow(dead_code)]
    pub(crate) encrypted_hmac_value: Vec<u8>,
}

#[derive(Debug, Clone)]
pub(crate) struct AgilePasswordKeyEncryptor {
    pub(crate) salt: Vec<u8>,
    pub(crate) block_size: usize,
    pub(crate) key_bits: usize,
    pub(crate) spin_count: u32,
    pub(crate) hash_algorithm: HashAlgorithm,
    pub(crate) cipher_algorithm: String,
    pub(crate) cipher_chaining: String,
    pub(crate) encrypted_verifier_hash_input: Vec<u8>,
    pub(crate) encrypted_verifier_hash_value: Vec<u8>,
    pub(crate) encrypted_key_value: Vec<u8>,
}

pub(crate) fn parse_agile_encryption_info(
    bytes: &[u8],
    header: &EncryptionInfoHeader,
) -> Result<AgileEncryptionInfo, OfficeCryptoError> {
    let start = header.header_offset;
    let xml_len = header.header_size as usize;
    let xml_bytes = bytes.get(start..start + xml_len).ok_or_else(|| {
        OfficeCryptoError::InvalidFormat("EncryptionInfo XML size out of range".to_string())
    })?;
    let xml_str = std::str::from_utf8(xml_bytes).map_err(|_| {
        OfficeCryptoError::InvalidFormat("EncryptionInfo XML is not valid UTF-8".to_string())
    })?;

    let descriptor = parse_agile_descriptor(xml_str)?;

    Ok(AgileEncryptionInfo {
        version_major: header.version_major,
        version_minor: header.version_minor,
        flags: header.flags,
        key_data: descriptor.key_data,
        data_integrity: descriptor.data_integrity,
        password_key_encryptor: descriptor.password_key_encryptor,
    })
}

pub(crate) fn decrypt_agile_encrypted_package(
    info: &AgileEncryptionInfo,
    encrypted_package: &[u8],
    password: &str,
) -> Result<Vec<u8>, OfficeCryptoError> {
    if encrypted_package.len() < 8 {
        return Err(OfficeCryptoError::InvalidFormat(
            "EncryptedPackage stream too short".to_string(),
        ));
    }
    let total_size = read_u64_le(encrypted_package, 0)?;
    let expected_len = checked_vec_len(total_size)?;
    let ciphertext = &encrypted_package[8..];

    if info.key_data.cipher_algorithm != "AES" {
        return Err(OfficeCryptoError::UnsupportedEncryption(format!(
            "unsupported cipherAlgorithm {}",
            info.key_data.cipher_algorithm
        )));
    }
    if info.key_data.cipher_chaining != "ChainingModeCBC" {
        return Err(OfficeCryptoError::UnsupportedEncryption(format!(
            "unsupported cipherChaining {}",
            info.key_data.cipher_chaining
        )));
    }

    if info.password_key_encryptor.cipher_algorithm != "AES" {
        return Err(OfficeCryptoError::UnsupportedEncryption(format!(
            "unsupported password cipherAlgorithm {}",
            info.password_key_encryptor.cipher_algorithm
        )));
    }
    if info.password_key_encryptor.cipher_chaining != "ChainingModeCBC" {
        return Err(OfficeCryptoError::UnsupportedEncryption(format!(
            "unsupported password cipherChaining {}",
            info.password_key_encryptor.cipher_chaining
        )));
    }

    let pw_utf16 = password_to_utf16le(password);

    // In MS-OFFCRYPTO Agile encryption, the password key encryptor uses `saltValue` directly as the
    // AES-CBC IV when decrypting verifier and key blobs. (The per-purpose block keys are only for
    // key derivation.)
    let password_block_size = info.password_key_encryptor.block_size;
    if password_block_size != 16 {
        return Err(OfficeCryptoError::UnsupportedEncryption(format!(
            "unsupported password blockSize {password_block_size}"
        )));
    }
    let password_iv = info
        .password_key_encryptor
        .salt
        .get(..password_block_size)
        .ok_or_else(|| {
            OfficeCryptoError::InvalidFormat(format!(
                "password saltValue shorter than blockSize ({password_block_size})"
            ))
        })?;

    // Password verification.
    let verifier_input_key = derive_agile_key(
        info.password_key_encryptor.hash_algorithm,
        &info.password_key_encryptor.salt,
        &pw_utf16,
        info.password_key_encryptor.spin_count,
        info.password_key_encryptor.key_bits / 8,
        BLOCK_KEY_VERIFIER_HASH_INPUT,
    );
    let verifier_hash_input_plain: Zeroizing<Vec<u8>> = Zeroizing::new(aes_cbc_decrypt(
        &verifier_input_key,
        password_iv,
        &info.password_key_encryptor.encrypted_verifier_hash_input,
    )?);
    let verifier_hash_input_slice = verifier_hash_input_plain
        .get(..password_block_size)
        .ok_or_else(|| {
            OfficeCryptoError::InvalidFormat(
                "decrypted verifierHashInput shorter than 16 bytes".to_string(),
            )
        })?;

    let verifier_hash: Zeroizing<Vec<u8>> = Zeroizing::new(
        info.password_key_encryptor
            .hash_algorithm
            .digest(verifier_hash_input_slice),
    );

    let verifier_value_key = derive_agile_key(
        info.password_key_encryptor.hash_algorithm,
        &info.password_key_encryptor.salt,
        &pw_utf16,
        info.password_key_encryptor.spin_count,
        info.password_key_encryptor.key_bits / 8,
        BLOCK_KEY_VERIFIER_HASH_VALUE,
    );
    let verifier_hash_value_plain: Zeroizing<Vec<u8>> = Zeroizing::new(aes_cbc_decrypt(
        &verifier_value_key,
        password_iv,
        &info.password_key_encryptor.encrypted_verifier_hash_value,
    )?);
    let expected_hash_slice = verifier_hash_value_plain
        .get(..verifier_hash.len())
        .ok_or_else(|| {
            OfficeCryptoError::InvalidFormat(
                "decrypted verifierHashValue shorter than hash".to_string(),
            )
        })?;

    if !ct_eq(expected_hash_slice, verifier_hash.as_slice()) {
        return Err(OfficeCryptoError::InvalidPassword);
    }

    // Decrypt the package key.
    let key_value_key = derive_agile_key(
        info.password_key_encryptor.hash_algorithm,
        &info.password_key_encryptor.salt,
        &pw_utf16,
        info.password_key_encryptor.spin_count,
        info.password_key_encryptor.key_bits / 8,
        BLOCK_KEY_ENCRYPTED_KEY_VALUE,
    );
    let key_value_plain: Zeroizing<Vec<u8>> = Zeroizing::new(aes_cbc_decrypt(
        &key_value_key,
        password_iv,
        &info.password_key_encryptor.encrypted_key_value,
    )?);
    let key_len = info.key_data.key_bits / 8;
    let package_key_bytes = key_value_plain.get(..key_len).ok_or_else(|| {
        OfficeCryptoError::InvalidFormat("decrypted keyValue shorter than keyBytes".to_string())
    })?;
    let package_key: Zeroizing<Vec<u8>> = Zeroizing::new(package_key_bytes.to_vec());

    // Validate data integrity (HMAC over the entire EncryptedPackage stream).
    //
    // The HMAC key/value are encrypted using the package key, with IVs derived from the keyData
    // salt and fixed block keys.
    let digest_len = info.key_data.hash_algorithm.digest_len();
    let iv_hmac_key = derive_iv(
        info.key_data.hash_algorithm,
        &info.key_data.salt,
        BLOCK_KEY_INTEGRITY_HMAC_KEY,
        info.key_data.block_size,
    );
    let hmac_key_plain: Zeroizing<Vec<u8>> = Zeroizing::new(aes_cbc_decrypt(
        &package_key,
        &iv_hmac_key,
        &info.data_integrity.encrypted_hmac_key,
    )?);
    let hmac_key_plain = hmac_key_plain.get(..digest_len).ok_or_else(|| {
        OfficeCryptoError::InvalidFormat(
            "decrypted encryptedHmacKey shorter than hash output".to_string(),
        )
    })?;

    let iv_hmac_val = derive_iv(
        info.key_data.hash_algorithm,
        &info.key_data.salt,
        BLOCK_KEY_INTEGRITY_HMAC_VALUE,
        info.key_data.block_size,
    );
    let hmac_value_plain: Zeroizing<Vec<u8>> = Zeroizing::new(aes_cbc_decrypt(
        &package_key,
        &iv_hmac_val,
        &info.data_integrity.encrypted_hmac_value,
    )?);
    let expected_hmac = hmac_value_plain.get(..digest_len).ok_or_else(|| {
        OfficeCryptoError::InvalidFormat(
            "decrypted encryptedHmacValue shorter than hash output".to_string(),
        )
    })?;

    let computed_hmac = compute_hmac(info.key_data.hash_algorithm, hmac_key_plain, encrypted_package);
    if !ct_eq(expected_hmac, &computed_hmac) {
        return Err(OfficeCryptoError::IntegrityCheckFailed);
    }

    // Decrypt the package data in 4096-byte segments.
    const SEGMENT_LEN: usize = 4096;
    let mut out = Vec::new();
    out.try_reserve_exact(ciphertext.len()).map_err(|source| {
        OfficeCryptoError::EncryptedPackageAllocationFailed { total_size, source }
    })?;
    let mut offset = 0usize;
    let mut block_index = 0u32;
    while offset < ciphertext.len() {
        let seg_len = (ciphertext.len() - offset).min(SEGMENT_LEN);
        let seg = &ciphertext[offset..offset + seg_len];
        let iv = derive_iv(
            info.key_data.hash_algorithm,
            &info.key_data.salt,
            &block_index.to_le_bytes(),
            info.key_data.block_size,
        );
        let mut plain = aes_cbc_decrypt(&package_key, &iv, seg)?;
        out.append(&mut plain);
        offset += seg_len;
        block_index = block_index.checked_add(1).ok_or_else(|| {
            OfficeCryptoError::InvalidFormat("segment counter overflow".to_string())
        })?;
    }
    if expected_len > out.len() {
        return Err(OfficeCryptoError::InvalidFormat(format!(
            "decrypted package length {} shorter than expected {}",
            out.len(),
            expected_len
        )));
    }
    out.truncate(expected_len);
    Ok(out)
}

pub(crate) fn encrypt_agile_encrypted_package(
    zip_bytes: &[u8],
    password: &str,
    opts: &crate::EncryptOptions,
) -> Result<(Vec<u8>, Vec<u8>), OfficeCryptoError> {
    if opts.key_bits % 8 != 0 {
        return Err(OfficeCryptoError::InvalidOptions(
            "key_bits must be divisible by 8".to_string(),
        ));
    }
    if opts.key_bits != 128 && opts.key_bits != 256 {
        return Err(OfficeCryptoError::InvalidOptions(format!(
            "unsupported key_bits {} (expected 128 or 256)",
            opts.key_bits
        )));
    }

    let key_bytes = opts.key_bits / 8;
    let block_size = 16usize;
    let hash_alg = opts.hash_algorithm;

    let pw_utf16 = password_to_utf16le(password);

    // Random salts and keys.
    let mut salt_key_encryptor = vec![0u8; 16];
    let mut salt_key_data = vec![0u8; 16];
    OsRng.fill_bytes(&mut salt_key_encryptor);
    OsRng.fill_bytes(&mut salt_key_data);

    let mut package_key_plain = vec![0u8; key_bytes];
    OsRng.fill_bytes(&mut package_key_plain);
    let package_key_plain: Zeroizing<Vec<u8>> = Zeroizing::new(package_key_plain);

    let mut verifier_hash_input_plain = [0u8; 16];
    OsRng.fill_bytes(&mut verifier_hash_input_plain);
    let verifier_hash_value_plain = hash_alg.digest(&verifier_hash_input_plain);
    let verifier_hash_value_plain = pad_zero(&verifier_hash_value_plain, block_size);

    // See `decrypt_agile_encrypted_package`: password-key-encryptor fields use `saltValue`
    // as the IV (truncated to blockSize).
    let verifier_iv = salt_key_encryptor
        .get(..block_size)
        .ok_or_else(|| OfficeCryptoError::InvalidFormat("saltValue shorter than blockSize".to_string()))?;

    // Encrypt verifierHashInput.
    let key_vhi = derive_agile_key(
        hash_alg,
        &salt_key_encryptor,
        &pw_utf16,
        opts.spin_count,
        key_bytes,
        BLOCK_KEY_VERIFIER_HASH_INPUT,
    );
    let enc_vhi = aes_cbc_encrypt(&key_vhi, verifier_iv, &verifier_hash_input_plain)?;

    // Encrypt verifierHashValue.
    let key_vhv = derive_agile_key(
        hash_alg,
        &salt_key_encryptor,
        &pw_utf16,
        opts.spin_count,
        key_bytes,
        BLOCK_KEY_VERIFIER_HASH_VALUE,
    );
    let enc_vhv = aes_cbc_encrypt(&key_vhv, verifier_iv, &verifier_hash_value_plain)?;

    // Encrypt package key (encryptedKeyValue).
    let key_kv = derive_agile_key(
        hash_alg,
        &salt_key_encryptor,
        &pw_utf16,
        opts.spin_count,
        key_bytes,
        BLOCK_KEY_ENCRYPTED_KEY_VALUE,
    );
    let enc_kv = aes_cbc_encrypt(&key_kv, verifier_iv, &package_key_plain)?;

    // Encrypt package bytes.
    let encrypted_package = encrypt_encrypted_package_stream(
        zip_bytes,
        &package_key_plain,
        hash_alg,
        &salt_key_data,
        block_size,
    )?;

    // Integrity (HMAC over the EncryptedPackage stream).
    let mut hmac_key_plain = vec![0u8; hash_alg.digest_len()];
    OsRng.fill_bytes(&mut hmac_key_plain);
    let hmac_key_plain: Zeroizing<Vec<u8>> = Zeroizing::new(hmac_key_plain);
    let hmac_value_plain = compute_hmac(hash_alg, &hmac_key_plain, &encrypted_package);
    let hmac_value_plain = pad_zero(&hmac_value_plain, block_size);

    let iv_hmac_key = derive_iv(
        hash_alg,
        &salt_key_data,
        BLOCK_KEY_INTEGRITY_HMAC_KEY,
        block_size,
    );
    let encrypted_hmac_key = aes_cbc_encrypt(
        &package_key_plain,
        &iv_hmac_key,
        &pad_zero(&hmac_key_plain, block_size),
    )?;
    let iv_hmac_val = derive_iv(
        hash_alg,
        &salt_key_data,
        BLOCK_KEY_INTEGRITY_HMAC_VALUE,
        block_size,
    );
    let encrypted_hmac_value =
        aes_cbc_encrypt(&package_key_plain, &iv_hmac_val, &hmac_value_plain)?;

    // Build EncryptionInfo XML.
    let b64 = base64::engine::general_purpose::STANDARD;
    let salt_key_encryptor_b64 = b64.encode(&salt_key_encryptor);
    let salt_key_data_b64 = b64.encode(&salt_key_data);
    let enc_vhi_b64 = b64.encode(enc_vhi);
    let enc_vhv_b64 = b64.encode(enc_vhv);
    let enc_kv_b64 = b64.encode(enc_kv);
    let enc_hmac_key_b64 = b64.encode(encrypted_hmac_key);
    let enc_hmac_value_b64 = b64.encode(encrypted_hmac_value);

    let xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption">
  <keyData saltSize="16" blockSize="16" keyBits="{key_bits}" hashAlgorithm="{hash_alg_name}" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" saltValue="{salt_key_data_b64}"/>
  <dataIntegrity encryptedHmacKey="{enc_hmac_key_b64}" encryptedHmacValue="{enc_hmac_value_b64}"/>
  <keyEncryptors>
    <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
      <p:encryptedKey xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password"
        saltSize="16" blockSize="16" keyBits="{key_bits}" spinCount="{spin_count}" hashAlgorithm="{hash_alg_name}" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" saltValue="{salt_key_encryptor_b64}">
        <p:encryptedVerifierHashInput>{enc_vhi_b64}</p:encryptedVerifierHashInput>
        <p:encryptedVerifierHashValue>{enc_vhv_b64}</p:encryptedVerifierHashValue>
        <p:encryptedKeyValue>{enc_kv_b64}</p:encryptedKeyValue>
      </p:encryptedKey>
    </keyEncryptor>
  </keyEncryptors>
</encryption>"#,
        key_bits = opts.key_bits,
        spin_count = opts.spin_count,
        hash_alg_name = hash_alg.as_ooxml_name(),
        salt_key_data_b64 = salt_key_data_b64,
        salt_key_encryptor_b64 = salt_key_encryptor_b64,
        enc_vhi_b64 = enc_vhi_b64,
        enc_vhv_b64 = enc_vhv_b64,
        enc_kv_b64 = enc_kv_b64,
        enc_hmac_key_b64 = enc_hmac_key_b64,
        enc_hmac_value_b64 = enc_hmac_value_b64,
    );

    // Build EncryptionInfo stream: version header + xml length + xml bytes.
    let flags: u32 = 0x0000_0040;
    let xml_len = xml.as_bytes().len() as u32;
    let mut encryption_info = Vec::new();
    encryption_info.extend_from_slice(&4u16.to_le_bytes());
    encryption_info.extend_from_slice(&4u16.to_le_bytes());
    encryption_info.extend_from_slice(&flags.to_le_bytes());
    encryption_info.extend_from_slice(&xml_len.to_le_bytes());
    encryption_info.extend_from_slice(xml.as_bytes());

    Ok((encryption_info, encrypted_package))
}

fn encrypt_encrypted_package_stream(
    zip_bytes: &[u8],
    package_key: &[u8],
    hash_alg: HashAlgorithm,
    salt: &[u8],
    block_size: usize,
) -> Result<Vec<u8>, OfficeCryptoError> {
    const SEGMENT_LEN: usize = 4096;
    let original_size = zip_bytes.len() as u64;
    let mut out = Vec::with_capacity(8 + zip_bytes.len());
    out.extend_from_slice(&original_size.to_le_bytes());

    let mut block_index = 0u32;
    for chunk in zip_bytes.chunks(SEGMENT_LEN) {
        let iv = derive_iv(hash_alg, salt, &block_index.to_le_bytes(), block_size);
        let plain = pad_zero(chunk, block_size);
        let enc = aes_cbc_encrypt(package_key, &iv, &plain)?;
        out.extend_from_slice(&enc);
        block_index = block_index.checked_add(1).ok_or_else(|| {
            OfficeCryptoError::InvalidFormat("segment counter overflow".to_string())
        })?;
    }

    Ok(out)
}

fn pad_zero(data: &[u8], block_size: usize) -> Vec<u8> {
    if data.len() % block_size == 0 {
        return data.to_vec();
    }
    let mut out = data.to_vec();
    let pad = block_size - (out.len() % block_size);
    out.extend(std::iter::repeat(0u8).take(pad));
    out
}

fn compute_hmac(hash_alg: HashAlgorithm, key: &[u8], data: &[u8]) -> Vec<u8> {
    match hash_alg {
        HashAlgorithm::Md5 => {
            let mut mac: Hmac<md5::Md5> = Hmac::new_from_slice(key).expect("HMAC accepts any key size");
            mac.update(data);
            mac.finalize().into_bytes().to_vec()
        }
        HashAlgorithm::Sha1 => {
            let mut mac: Hmac<Sha1> = Hmac::new_from_slice(key).expect("HMAC accepts any key size");
            mac.update(data);
            mac.finalize().into_bytes().to_vec()
        }
        HashAlgorithm::Sha256 => {
            let mut mac: Hmac<Sha256> =
                Hmac::new_from_slice(key).expect("HMAC accepts any key size");
            mac.update(data);
            mac.finalize().into_bytes().to_vec()
        }
        HashAlgorithm::Sha384 => {
            let mut mac: Hmac<Sha384> =
                Hmac::new_from_slice(key).expect("HMAC accepts any key size");
            mac.update(data);
            mac.finalize().into_bytes().to_vec()
        }
        HashAlgorithm::Sha512 => {
            let mut mac: Hmac<Sha512> =
                Hmac::new_from_slice(key).expect("HMAC accepts any key size");
            mac.update(data);
            mac.finalize().into_bytes().to_vec()
        }
    }
}

struct AgileDescriptor {
    key_data: AgileKeyData,
    data_integrity: AgileDataIntegrity,
    password_key_encryptor: AgilePasswordKeyEncryptor,
}

fn parse_agile_descriptor(xml: &str) -> Result<AgileDescriptor, OfficeCryptoError> {
    let mut reader = Reader::from_str(xml);
    reader.trim_text(true);
    let mut buf = Vec::new();

    let mut key_data: Option<AgileKeyData> = None;
    let mut data_integrity: Option<AgileDataIntegrity> = None;
    let mut password_key_encryptor: Option<AgilePasswordKeyEncryptor> = None;

    let mut in_password_key_encryptor = false;
    let mut in_encrypted_key = false;
    let mut capture: Option<CaptureKind> = None;

    let mut tmp_encrypted_verifier_hash_input: Option<Vec<u8>> = None;
    let mut tmp_encrypted_verifier_hash_value: Option<Vec<u8>> = None;
    let mut tmp_encrypted_key_value: Option<Vec<u8>> = None;

    let mut tmp_password_attrs: Option<AgilePasswordAttrs> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let qname = e.name();
                let name = local_name(qname.as_ref());
                match name {
                    b"keyEncryptor" => {
                        // The `<keyEncryptor>` element indicates how the package key is protected
                        // (password vs certificate). We only support password-based decryption.
                        in_password_key_encryptor = is_password_key_encryptor(&e, &reader)?;
                    }
                    b"keyData" => {
                        let kd = parse_key_data_attrs(&e, &reader)?;
                        key_data = Some(kd);
                    }
                    b"dataIntegrity" => {
                        let di = parse_data_integrity_attrs(&e, &reader)?;
                        data_integrity = Some(di);
                    }
                    b"encryptedKey" if in_password_key_encryptor => {
                        in_encrypted_key = true;
                        tmp_password_attrs = Some(parse_password_key_encryptor_attrs(&e, &reader)?);

                        // Some producers (e.g. `ms_offcrypto_writer`) encode the verifier/key
                        // blobs as base64 attributes on the `<encryptedKey/>` element instead of
                        // child elements. Accept either form.
                        let (vhi, vhv, kv) = parse_encrypted_key_value_attrs(&e, &reader)?;
                        if vhi.is_some() {
                            tmp_encrypted_verifier_hash_input = vhi;
                        }
                        if vhv.is_some() {
                            tmp_encrypted_verifier_hash_value = vhv;
                        }
                        if kv.is_some() {
                            tmp_encrypted_key_value = kv;
                        }
                    }
                    b"encryptedVerifierHashInput" if in_encrypted_key => {
                        capture = Some(CaptureKind::VerifierHashInput);
                    }
                    b"encryptedVerifierHashValue" if in_encrypted_key => {
                        capture = Some(CaptureKind::VerifierHashValue);
                    }
                    b"encryptedKeyValue" if in_encrypted_key => {
                        capture = Some(CaptureKind::KeyValue);
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(e)) => {
                let qname = e.name();
                let name = local_name(qname.as_ref());
                match name {
                    b"keyData" => {
                        let kd = parse_key_data_attrs(&e, &reader)?;
                        key_data = Some(kd);
                    }
                    b"dataIntegrity" => {
                        let di = parse_data_integrity_attrs(&e, &reader)?;
                        data_integrity = Some(di);
                    }
                    b"encryptedKey" if in_password_key_encryptor => {
                        let attrs = parse_password_key_encryptor_attrs(&e, &reader)?;
                        let (vhi, vhv, kv) = parse_encrypted_key_value_attrs(&e, &reader)?;
                        password_key_encryptor = Some(AgilePasswordKeyEncryptor {
                            salt: attrs.salt,
                            block_size: attrs.block_size,
                            key_bits: attrs.key_bits,
                            spin_count: attrs.spin_count,
                            hash_algorithm: attrs.hash_algorithm,
                            cipher_algorithm: attrs.cipher_algorithm,
                            cipher_chaining: attrs.cipher_chaining,
                            encrypted_verifier_hash_input: vhi.ok_or_else(|| {
                                OfficeCryptoError::InvalidFormat(
                                    "missing encryptedVerifierHashInput".to_string(),
                                )
                            })?,
                            encrypted_verifier_hash_value: vhv.ok_or_else(|| {
                                OfficeCryptoError::InvalidFormat(
                                    "missing encryptedVerifierHashValue".to_string(),
                                )
                            })?,
                            encrypted_key_value: kv.ok_or_else(|| {
                                OfficeCryptoError::InvalidFormat(
                                    "missing encryptedKeyValue".to_string(),
                                )
                            })?,
                        });
                    }
                    _ => {}
                }
            }
            Ok(Event::End(e)) => {
                let qname = e.name();
                let name = local_name(qname.as_ref());
                if name == b"keyEncryptor" {
                    in_password_key_encryptor = false;
                }
                if name == b"encryptedKey" && in_encrypted_key {
                    in_encrypted_key = false;
                    capture = None;
                    let attrs = tmp_password_attrs.take().ok_or_else(|| {
                        OfficeCryptoError::InvalidFormat(
                            "encryptedKey missing required attributes".to_string(),
                        )
                    })?;
                    let encrypted_verifier_hash_input =
                        tmp_encrypted_verifier_hash_input.take().ok_or_else(|| {
                            OfficeCryptoError::InvalidFormat(
                                "missing encryptedVerifierHashInput".to_string(),
                            )
                        })?;
                    let encrypted_verifier_hash_value =
                        tmp_encrypted_verifier_hash_value.take().ok_or_else(|| {
                            OfficeCryptoError::InvalidFormat(
                                "missing encryptedVerifierHashValue".to_string(),
                            )
                        })?;
                    let encrypted_key_value = tmp_encrypted_key_value.take().ok_or_else(|| {
                        OfficeCryptoError::InvalidFormat("missing encryptedKeyValue".to_string())
                    })?;
                    password_key_encryptor = Some(AgilePasswordKeyEncryptor {
                        salt: attrs.salt,
                        block_size: attrs.block_size,
                        key_bits: attrs.key_bits,
                        spin_count: attrs.spin_count,
                        hash_algorithm: attrs.hash_algorithm,
                        cipher_algorithm: attrs.cipher_algorithm,
                        cipher_chaining: attrs.cipher_chaining,
                        encrypted_verifier_hash_input,
                        encrypted_verifier_hash_value,
                        encrypted_key_value,
                    });
                }
                if matches!(
                    name,
                    b"encryptedVerifierHashInput"
                        | b"encryptedVerifierHashValue"
                        | b"encryptedKeyValue"
                ) {
                    capture = None;
                }
            }
            Ok(Event::Text(t)) => {
                if let Some(kind) = capture {
                    let text = t
                        .unescape()
                        .map_err(|_| {
                            OfficeCryptoError::InvalidFormat(
                                "invalid XML escape in base64 text".to_string(),
                            )
                        })?
                        .to_string();
                    let decoded = decode_b64_attr(&text).map_err(|_| {
                        OfficeCryptoError::InvalidFormat(
                            "invalid base64 in EncryptionInfo".to_string(),
                        )
                    })?;
                    match kind {
                        CaptureKind::VerifierHashInput => {
                            tmp_encrypted_verifier_hash_input = Some(decoded)
                        }
                        CaptureKind::VerifierHashValue => {
                            tmp_encrypted_verifier_hash_value = Some(decoded)
                        }
                        CaptureKind::KeyValue => tmp_encrypted_key_value = Some(decoded),
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(OfficeCryptoError::InvalidFormat(format!(
                    "failed to parse EncryptionInfo XML: {e}"
                )))
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(AgileDescriptor {
        key_data: key_data.ok_or_else(|| {
            OfficeCryptoError::InvalidFormat("missing keyData element".to_string())
        })?,
        data_integrity: data_integrity.ok_or_else(|| {
            OfficeCryptoError::InvalidFormat("missing dataIntegrity element".to_string())
        })?,
        password_key_encryptor: password_key_encryptor.ok_or_else(|| {
            OfficeCryptoError::InvalidFormat("missing password keyEncryptor".to_string())
        })?,
    })
}

#[derive(Clone, Copy)]
enum CaptureKind {
    VerifierHashInput,
    VerifierHashValue,
    KeyValue,
}

fn local_name(name: &[u8]) -> &[u8] {
    match name.iter().position(|&b| b == b':') {
        Some(idx) => &name[idx + 1..],
        None => name,
    }
}

fn is_password_key_encryptor(
    e: &quick_xml::events::BytesStart<'_>,
    reader: &Reader<impl std::io::BufRead>,
) -> Result<bool, OfficeCryptoError> {
    const PASSWORD_URI: &str = "http://schemas.microsoft.com/office/2006/keyEncryptor/password";
    for attr in e.attributes() {
        let attr = attr
            .map_err(|_| OfficeCryptoError::InvalidFormat("invalid XML attribute".to_string()))?;
        let key = local_name(attr.key.as_ref());
        if key != b"uri" {
            continue;
        }
        let value = attr.decode_and_unescape_value(reader).map_err(|_| {
            OfficeCryptoError::InvalidFormat("invalid XML attribute encoding".to_string())
        })?;
        return Ok(value.as_ref() == PASSWORD_URI);
    }
    Ok(false)
}

fn parse_key_data_attrs(
    e: &quick_xml::events::BytesStart<'_>,
    reader: &Reader<impl std::io::BufRead>,
) -> Result<AgileKeyData, OfficeCryptoError> {
    let mut salt_value: Option<Vec<u8>> = None;
    let mut block_size: Option<usize> = None;
    let mut key_bits: Option<usize> = None;
    let mut hash_algorithm: Option<HashAlgorithm> = None;
    let mut cipher_algorithm: Option<String> = None;
    let mut cipher_chaining: Option<String> = None;

    for attr in e.attributes() {
        let attr = attr
            .map_err(|_| OfficeCryptoError::InvalidFormat("invalid XML attribute".to_string()))?;
        let key = local_name(attr.key.as_ref());
        let value = attr.decode_and_unescape_value(reader).map_err(|_| {
            OfficeCryptoError::InvalidFormat("invalid XML attribute encoding".to_string())
        })?;
        match key {
            b"saltValue" => {
                salt_value = Some(decode_b64_attr(value.as_ref()).map_err(|_| {
                    OfficeCryptoError::InvalidFormat("invalid base64 saltValue".to_string())
                })?);
            }
            b"blockSize" => {
                block_size = Some(value.parse::<usize>().map_err(|_| {
                    OfficeCryptoError::InvalidFormat("invalid blockSize".to_string())
                })?);
            }
            b"keyBits" => {
                key_bits = Some(value.parse::<usize>().map_err(|_| {
                    OfficeCryptoError::InvalidFormat("invalid keyBits".to_string())
                })?);
            }
            b"hashAlgorithm" => {
                hash_algorithm = Some(HashAlgorithm::from_name(value.as_ref())?);
            }
            b"cipherAlgorithm" => {
                cipher_algorithm = Some(value.as_ref().to_string());
            }
            b"cipherChaining" => {
                cipher_chaining = Some(value.as_ref().to_string());
            }
            _ => {}
        }
    }

    Ok(AgileKeyData {
        salt: salt_value.ok_or_else(|| {
            OfficeCryptoError::InvalidFormat("keyData missing saltValue".to_string())
        })?,
        block_size: block_size.ok_or_else(|| {
            OfficeCryptoError::InvalidFormat("keyData missing blockSize".to_string())
        })?,
        key_bits: key_bits.ok_or_else(|| {
            OfficeCryptoError::InvalidFormat("keyData missing keyBits".to_string())
        })?,
        hash_algorithm: hash_algorithm.ok_or_else(|| {
            OfficeCryptoError::InvalidFormat("keyData missing hashAlgorithm".to_string())
        })?,
        cipher_algorithm: cipher_algorithm.ok_or_else(|| {
            OfficeCryptoError::InvalidFormat("keyData missing cipherAlgorithm".to_string())
        })?,
        cipher_chaining: cipher_chaining.ok_or_else(|| {
            OfficeCryptoError::InvalidFormat("keyData missing cipherChaining".to_string())
        })?,
    })
}

fn parse_data_integrity_attrs(
    e: &quick_xml::events::BytesStart<'_>,
    reader: &Reader<impl std::io::BufRead>,
) -> Result<AgileDataIntegrity, OfficeCryptoError> {
    let mut encrypted_hmac_key: Option<Vec<u8>> = None;
    let mut encrypted_hmac_value: Option<Vec<u8>> = None;
    for attr in e.attributes() {
        let attr = attr
            .map_err(|_| OfficeCryptoError::InvalidFormat("invalid XML attribute".to_string()))?;
        let key = local_name(attr.key.as_ref());
        let value = attr.decode_and_unescape_value(reader).map_err(|_| {
            OfficeCryptoError::InvalidFormat("invalid XML attribute encoding".to_string())
        })?;
        match key {
            b"encryptedHmacKey" => {
                encrypted_hmac_key = Some(decode_b64_attr(value.as_ref()).map_err(|_| {
                    OfficeCryptoError::InvalidFormat(
                        "invalid base64 encryptedHmacKey".to_string(),
                    )
                })?);
            }
            b"encryptedHmacValue" => {
                encrypted_hmac_value = Some(decode_b64_attr(value.as_ref()).map_err(|_| {
                    OfficeCryptoError::InvalidFormat(
                        "invalid base64 encryptedHmacValue".to_string(),
                    )
                })?);
            }
            _ => {}
        }
    }
    Ok(AgileDataIntegrity {
        encrypted_hmac_key: encrypted_hmac_key.ok_or_else(|| {
            OfficeCryptoError::InvalidFormat("dataIntegrity missing encryptedHmacKey".to_string())
        })?,
        encrypted_hmac_value: encrypted_hmac_value.ok_or_else(|| {
            OfficeCryptoError::InvalidFormat("dataIntegrity missing encryptedHmacValue".to_string())
        })?,
    })
}

fn parse_encrypted_key_value_attrs(
    e: &quick_xml::events::BytesStart<'_>,
    reader: &Reader<impl std::io::BufRead>,
) -> Result<(Option<Vec<u8>>, Option<Vec<u8>>, Option<Vec<u8>>), OfficeCryptoError> {
    let mut encrypted_verifier_hash_input: Option<Vec<u8>> = None;
    let mut encrypted_verifier_hash_value: Option<Vec<u8>> = None;
    let mut encrypted_key_value: Option<Vec<u8>> = None;

    for attr in e.attributes() {
        let attr = attr
            .map_err(|_| OfficeCryptoError::InvalidFormat("invalid XML attribute".to_string()))?;
        let key = local_name(attr.key.as_ref());
        let value = attr.decode_and_unescape_value(reader).map_err(|_| {
            OfficeCryptoError::InvalidFormat("invalid XML attribute encoding".to_string())
        })?;
        match key {
            b"encryptedVerifierHashInput" => {
                encrypted_verifier_hash_input = Some(decode_b64_attr(value.as_ref()).map_err(|_| {
                    OfficeCryptoError::InvalidFormat(
                        "invalid base64 encryptedVerifierHashInput".to_string(),
                    )
                })?);
            }
            b"encryptedVerifierHashValue" => {
                encrypted_verifier_hash_value = Some(decode_b64_attr(value.as_ref()).map_err(|_| {
                    OfficeCryptoError::InvalidFormat(
                        "invalid base64 encryptedVerifierHashValue".to_string(),
                    )
                })?);
            }
            b"encryptedKeyValue" => {
                encrypted_key_value = Some(decode_b64_attr(value.as_ref()).map_err(|_| {
                    OfficeCryptoError::InvalidFormat(
                        "invalid base64 encryptedKeyValue".to_string(),
                    )
                })?);
            }
            _ => {}
        }
    }

    Ok((
        encrypted_verifier_hash_input,
        encrypted_verifier_hash_value,
        encrypted_key_value,
    ))
}

#[derive(Debug)]
struct AgilePasswordAttrs {
    salt: Vec<u8>,
    block_size: usize,
    key_bits: usize,
    spin_count: u32,
    hash_algorithm: HashAlgorithm,
    cipher_algorithm: String,
    cipher_chaining: String,
}

fn parse_password_key_encryptor_attrs(
    e: &quick_xml::events::BytesStart<'_>,
    reader: &Reader<impl std::io::BufRead>,
) -> Result<AgilePasswordAttrs, OfficeCryptoError> {
    let mut salt_value: Option<Vec<u8>> = None;
    let mut block_size: Option<usize> = None;
    let mut key_bits: Option<usize> = None;
    let mut spin_count: Option<u32> = None;
    let mut hash_algorithm: Option<HashAlgorithm> = None;
    let mut cipher_algorithm: Option<String> = None;
    let mut cipher_chaining: Option<String> = None;

    for attr in e.attributes() {
        let attr = attr
            .map_err(|_| OfficeCryptoError::InvalidFormat("invalid XML attribute".to_string()))?;
        let key = local_name(attr.key.as_ref());
        let value = attr.decode_and_unescape_value(reader).map_err(|_| {
            OfficeCryptoError::InvalidFormat("invalid XML attribute encoding".to_string())
        })?;
        match key {
            b"saltValue" => {
                salt_value = Some(decode_b64_attr(value.as_ref()).map_err(|_| {
                    OfficeCryptoError::InvalidFormat("invalid base64 saltValue".to_string())
                })?);
            }
            b"blockSize" => {
                block_size = Some(value.parse::<usize>().map_err(|_| {
                    OfficeCryptoError::InvalidFormat("invalid blockSize".to_string())
                })?);
            }
            b"keyBits" => {
                key_bits = Some(value.parse::<usize>().map_err(|_| {
                    OfficeCryptoError::InvalidFormat("invalid keyBits".to_string())
                })?);
            }
            b"spinCount" => {
                spin_count = Some(value.parse::<u32>().map_err(|_| {
                    OfficeCryptoError::InvalidFormat("invalid spinCount".to_string())
                })?);
            }
            b"hashAlgorithm" => {
                hash_algorithm = Some(HashAlgorithm::from_name(value.as_ref())?);
            }
            b"cipherAlgorithm" => {
                cipher_algorithm = Some(value.as_ref().to_string());
            }
            b"cipherChaining" => {
                cipher_chaining = Some(value.as_ref().to_string());
            }
            _ => {}
        }
    }

    Ok(AgilePasswordAttrs {
        salt: salt_value.ok_or_else(|| {
            OfficeCryptoError::InvalidFormat("encryptedKey missing saltValue".to_string())
        })?,
        block_size: block_size.ok_or_else(|| {
            OfficeCryptoError::InvalidFormat("encryptedKey missing blockSize".to_string())
        })?,
        key_bits: key_bits.ok_or_else(|| {
            OfficeCryptoError::InvalidFormat("encryptedKey missing keyBits".to_string())
        })?,
        spin_count: spin_count.ok_or_else(|| {
            OfficeCryptoError::InvalidFormat("encryptedKey missing spinCount".to_string())
        })?,
        hash_algorithm: hash_algorithm.ok_or_else(|| {
            OfficeCryptoError::InvalidFormat("encryptedKey missing hashAlgorithm".to_string())
        })?,
        cipher_algorithm: cipher_algorithm.ok_or_else(|| {
            OfficeCryptoError::InvalidFormat("encryptedKey missing cipherAlgorithm".to_string())
        })?,
        cipher_chaining: cipher_chaining.ok_or_else(|| {
            OfficeCryptoError::InvalidFormat("encryptedKey missing cipherChaining".to_string())
        })?,
    })
}

fn decode_b64_attr(value: &str) -> Result<Vec<u8>, base64::DecodeError> {
    let bytes = value.as_bytes();

    // Avoid allocating for the common case where no whitespace is present.
    let mut cleaned: Option<Vec<u8>> = None;
    for (idx, &b) in bytes.iter().enumerate() {
        if matches!(b, b'\r' | b'\n' | b'\t' | b' ') {
            let mut out = Vec::with_capacity(bytes.len());
            out.extend_from_slice(&bytes[..idx]);
            for &b2 in &bytes[idx..] {
                if !matches!(b2, b'\r' | b'\n' | b'\t' | b' ') {
                    out.push(b2);
                }
            }
            cleaned = Some(out);
            break;
        }
    }

    let input = cleaned.as_deref().unwrap_or(bytes);
    BASE64_STANDARD
        .decode(input)
        .or_else(|_| BASE64_STANDARD_NO_PAD.decode(input))
}

#[cfg(test)]
pub(crate) mod tests {
    use super::*;
    use crate::crypto::{
        aes_cbc_encrypt, derive_agile_key, derive_iv, password_to_utf16le, HashAlgorithm,
    };
    use crate::util::parse_encryption_info_header;

    pub(crate) fn agile_encryption_info_fixture() -> Vec<u8> {
        // A small, deterministic Agile EncryptionInfo fixture for parsing tests.
        let xml = agile_descriptor_fixture_xml();
        let version_major = 4u16;
        let version_minor = 4u16;
        let flags = 0x0000_0040u32;
        let xml_len = xml.as_bytes().len() as u32;

        let mut out = Vec::new();
        out.extend_from_slice(&version_major.to_le_bytes());
        out.extend_from_slice(&version_minor.to_le_bytes());
        out.extend_from_slice(&flags.to_le_bytes());
        out.extend_from_slice(&xml_len.to_le_bytes());
        out.extend_from_slice(xml.as_bytes());

        let hdr = parse_encryption_info_header(&out).expect("header");
        assert_eq!(hdr.kind, crate::util::EncryptionInfoKind::Agile);
        out
    }

    pub(crate) fn agile_descriptor_fixture_xml() -> String {
        // Build a minimal-but-valid agile descriptor (values not meant to be secure).
        let password = "Password";
        let pw_utf16 = password_to_utf16le(password);
        let hash_alg = HashAlgorithm::Sha512;
        let spin_count = 100_000u32;
        let key_bits = 256usize;
        let block_size = 16usize;

        let salt_key_encryptor: [u8; 16] = [
            0xA0, 0xA1, 0xA2, 0xA3, 0xA4, 0xA5, 0xA6, 0xA7, 0xA8, 0xA9, 0xAA, 0xAB, 0xAC, 0xAD,
            0xAE, 0xAF,
        ];
        let salt_key_data: [u8; 16] = [
            0xB0, 0xB1, 0xB2, 0xB3, 0xB4, 0xB5, 0xB6, 0xB7, 0xB8, 0xB9, 0xBA, 0xBB, 0xBC, 0xBD,
            0xBE, 0xBF,
        ];

        let verifier_hash_input_plain: [u8; 16] = *b"formula-agl-test";
        let verifier_hash_value_plain = hash_alg.digest(&verifier_hash_input_plain);
        let package_key_plain: [u8; 32] = [0x11; 32];

        let key_vhi = derive_agile_key(
            hash_alg,
            &salt_key_encryptor,
            &pw_utf16,
            spin_count,
            key_bits / 8,
            BLOCK_KEY_VERIFIER_HASH_INPUT,
        );
        let iv_vhi = derive_iv(
            hash_alg,
            &salt_key_encryptor,
            BLOCK_KEY_VERIFIER_HASH_INPUT,
            block_size,
        );
        let enc_vhi =
            aes_cbc_encrypt(&key_vhi, &iv_vhi, &verifier_hash_input_plain).expect("enc vhi");

        let key_vhv = derive_agile_key(
            hash_alg,
            &salt_key_encryptor,
            &pw_utf16,
            spin_count,
            key_bits / 8,
            BLOCK_KEY_VERIFIER_HASH_VALUE,
        );
        let iv_vhv = derive_iv(
            hash_alg,
            &salt_key_encryptor,
            BLOCK_KEY_VERIFIER_HASH_VALUE,
            block_size,
        );
        let enc_vhv =
            aes_cbc_encrypt(&key_vhv, &iv_vhv, &verifier_hash_value_plain).expect("enc vhv");

        let key_kv = derive_agile_key(
            hash_alg,
            &salt_key_encryptor,
            &pw_utf16,
            spin_count,
            key_bits / 8,
            BLOCK_KEY_ENCRYPTED_KEY_VALUE,
        );
        let iv_kv = derive_iv(
            hash_alg,
            &salt_key_encryptor,
            BLOCK_KEY_ENCRYPTED_KEY_VALUE,
            block_size,
        );
        let enc_kv = aes_cbc_encrypt(&key_kv, &iv_kv, &package_key_plain).expect("enc key");

        let b64 = base64::engine::general_purpose::STANDARD;
        let salt_key_encryptor_b64 = b64.encode(salt_key_encryptor);
        let salt_key_data_b64 = b64.encode(salt_key_data);
        let enc_vhi_b64 = b64.encode(enc_vhi);
        let enc_vhv_b64 = b64.encode(enc_vhv);
        let enc_kv_b64 = b64.encode(enc_kv);

        // Dummy integrity fields.
        let dummy = b64.encode([0u8; 32]);

        format!(
            r#"<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption">
  <keyData saltSize="16" blockSize="16" keyBits="256" hashAlgorithm="SHA512" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" saltValue="{salt_key_data_b64}"/>
  <dataIntegrity encryptedHmacKey="{dummy}" encryptedHmacValue="{dummy}"/>
  <keyEncryptors>
    <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
      <p:encryptedKey xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password"
        saltSize="16" blockSize="16" keyBits="256" spinCount="100000" hashAlgorithm="SHA512" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" saltValue="{salt_key_encryptor_b64}">
        <p:encryptedVerifierHashInput>{enc_vhi_b64}</p:encryptedVerifierHashInput>
        <p:encryptedVerifierHashValue>{enc_vhv_b64}</p:encryptedVerifierHashValue>
        <p:encryptedKeyValue>{enc_kv_b64}</p:encryptedKeyValue>
      </p:encryptedKey>
    </keyEncryptor>
  </keyEncryptors>
</encryption>"#
        )
    }
}
