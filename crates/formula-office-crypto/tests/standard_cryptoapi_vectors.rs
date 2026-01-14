use aes::cipher::{generic_array::GenericArray, BlockEncrypt, KeyInit};
use aes::{Aes128, Aes192, Aes256};
use formula_office_crypto::{decrypt_standard_encrypted_package, OfficeCryptoError};
use sha1::{Digest as _, Sha1};

// Worked example from `docs/offcrypto-standard-cryptoapi.md`.
const SALT_00_TO_0F: [u8; 16] = [
    0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0A, 0x0B, 0x0C, 0x0D,
    0x0E, 0x0F,
];

const EXPECTED_KEY_AES256_BLOCK0: [u8; 32] = [
    0xDE, 0x54, 0x51, 0xB9, 0xDC, 0x3F, 0xCB, 0x38, 0x37, 0x92, 0xCB, 0xEE, 0xC8, 0x0B,
    0x6B, 0xC3, 0x07, 0x95, 0xC2, 0x70, 0x5E, 0x07, 0x50, 0x39, 0x40, 0x71, 0x99, 0xF7,
    0xD2, 0x99, 0xB6, 0xE4,
];

const EXPECTED_KEY_AES192_BLOCK0: [u8; 24] = [
    0xDE, 0x54, 0x51, 0xB9, 0xDC, 0x3F, 0xCB, 0x38, 0x37, 0x92, 0xCB, 0xEE, 0xC8, 0x0B,
    0x6B, 0xC3, 0x07, 0x95, 0xC2, 0x70, 0x5E, 0x07, 0x50, 0x39,
];

fn pad_zero(data: &[u8], block_size: usize) -> Vec<u8> {
    if data.len() % block_size == 0 {
        return data.to_vec();
    }
    let mut out = data.to_vec();
    let pad = block_size - (out.len() % block_size);
    out.extend(std::iter::repeat(0u8).take(pad));
    out
}

fn aes_ecb_encrypt(key: &[u8], plaintext: &[u8]) -> Vec<u8> {
    assert_eq!(plaintext.len() % 16, 0);
    let mut buf = plaintext.to_vec();

    fn encrypt_with<C>(key: &[u8], buf: &mut [u8])
    where
        C: BlockEncrypt + KeyInit,
    {
        let cipher = C::new_from_slice(key).expect("valid AES key");
        for block in buf.chunks_mut(16) {
            cipher.encrypt_block(GenericArray::from_mut_slice(block));
        }
    }

    match key.len() {
        16 => encrypt_with::<Aes128>(key, &mut buf),
        24 => encrypt_with::<Aes192>(key, &mut buf),
        32 => encrypt_with::<Aes256>(key, &mut buf),
        other => panic!("unsupported AES key length {other}"),
    }

    buf
}

fn build_standard_encryption_info_aes(
    alg_id: u32,
    key_bits: u32,
    salt: &[u8],
    encrypted_verifier: &[u8],
    encrypted_verifier_hash: &[u8],
) -> Vec<u8> {
    // MS-OFFCRYPTO Standard EncryptionInfo header.
    let version_major = 3u16;
    let version_minor = 2u16;
    let flags = 0x0000_0040u32;

    // EncryptionHeader (MS-OFFCRYPTO).
    let header_flags = 0x0000_0004u32 | 0x0000_0020u32; // fCryptoAPI | fAES
    let size_extra = 0u32;
    let alg_id_hash = 0x0000_8004u32; // CALG_SHA1
    let provider_type = 0x0000_0018u32; // PROV_RSA_AES (typical)
    let reserved1 = 0u32;
    let reserved2 = 0u32;
    let csp_name_utf16_nul = [0u8, 0u8];

    let mut header_bytes = Vec::new();
    header_bytes.extend_from_slice(&header_flags.to_le_bytes());
    header_bytes.extend_from_slice(&size_extra.to_le_bytes());
    header_bytes.extend_from_slice(&alg_id.to_le_bytes());
    header_bytes.extend_from_slice(&alg_id_hash.to_le_bytes());
    header_bytes.extend_from_slice(&key_bits.to_le_bytes());
    header_bytes.extend_from_slice(&provider_type.to_le_bytes());
    header_bytes.extend_from_slice(&reserved1.to_le_bytes());
    header_bytes.extend_from_slice(&reserved2.to_le_bytes());
    header_bytes.extend_from_slice(&csp_name_utf16_nul);

    let header_size = header_bytes.len() as u32;

    // EncryptionVerifier.
    let verifier_hash_size = 20u32; // SHA-1
    let mut verifier_bytes = Vec::new();
    verifier_bytes.extend_from_slice(&(salt.len() as u32).to_le_bytes());
    verifier_bytes.extend_from_slice(salt);
    verifier_bytes.extend_from_slice(encrypted_verifier);
    verifier_bytes.extend_from_slice(&verifier_hash_size.to_le_bytes());
    verifier_bytes.extend_from_slice(encrypted_verifier_hash);

    let mut out = Vec::new();
    out.extend_from_slice(&version_major.to_le_bytes());
    out.extend_from_slice(&version_minor.to_le_bytes());
    out.extend_from_slice(&flags.to_le_bytes());
    out.extend_from_slice(&header_size.to_le_bytes());
    out.extend_from_slice(&header_bytes);
    out.extend_from_slice(&verifier_bytes);
    out
}

fn build_encrypted_package_aes(key: &[u8], plaintext: &[u8]) -> Vec<u8> {
    let mut out = Vec::new();
    out.extend_from_slice(&(plaintext.len() as u64).to_le_bytes());
    let padded = pad_zero(plaintext, 16);
    let ciphertext = aes_ecb_encrypt(key, &padded);
    out.extend_from_slice(&ciphertext);
    out
}

#[test]
fn standard_cryptoapi_aes256_derivation_vector_decrypts_package() {
    // Valid OOXML/ZIP payload to satisfy the crate's zip sanity checks.
    let plaintext = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/xlsx/basic/basic.xlsx"
    ));

    // Deterministic verifier.
    let verifier_plain: [u8; 16] = *b"formula-std-test";
    let verifier_hash: [u8; 20] = Sha1::digest(&verifier_plain).into();
    let verifier_hash_padded = pad_zero(&verifier_hash, 16);

    let encrypted_verifier = aes_ecb_encrypt(&EXPECTED_KEY_AES256_BLOCK0, &verifier_plain);
    let encrypted_verifier_hash = aes_ecb_encrypt(&EXPECTED_KEY_AES256_BLOCK0, &verifier_hash_padded);

    let encryption_info = build_standard_encryption_info_aes(
        0x0000_6610, // CALG_AES_256
        256,
        &SALT_00_TO_0F,
        &encrypted_verifier,
        &encrypted_verifier_hash,
    );
    let encrypted_package = build_encrypted_package_aes(&EXPECTED_KEY_AES256_BLOCK0, plaintext);

    let decrypted =
        decrypt_standard_encrypted_package(&encryption_info, &encrypted_package, "password")
            .expect("decrypt");
    assert_eq!(decrypted, plaintext);

    let err = decrypt_standard_encrypted_package(&encryption_info, &encrypted_package, "wrong-password")
        .expect_err("wrong password");
    assert!(matches!(err, OfficeCryptoError::InvalidPassword));
}

#[test]
fn standard_cryptoapi_aes192_derivation_vector_decrypts_package() {
    let plaintext = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/xlsx/basic/basic.xlsx"
    ));

    let verifier_plain: [u8; 16] = *b"formula-std-test";
    let verifier_hash: [u8; 20] = Sha1::digest(&verifier_plain).into();
    let verifier_hash_padded = pad_zero(&verifier_hash, 16);

    let encrypted_verifier = aes_ecb_encrypt(&EXPECTED_KEY_AES192_BLOCK0, &verifier_plain);
    let encrypted_verifier_hash = aes_ecb_encrypt(&EXPECTED_KEY_AES192_BLOCK0, &verifier_hash_padded);

    let encryption_info = build_standard_encryption_info_aes(
        0x0000_660F, // CALG_AES_192
        192,
        &SALT_00_TO_0F,
        &encrypted_verifier,
        &encrypted_verifier_hash,
    );
    let encrypted_package = build_encrypted_package_aes(&EXPECTED_KEY_AES192_BLOCK0, plaintext);

    let decrypted =
        decrypt_standard_encrypted_package(&encryption_info, &encrypted_package, "password")
            .expect("decrypt");
    assert_eq!(decrypted, plaintext);
}
