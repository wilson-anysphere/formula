use aes::cipher::{generic_array::GenericArray, BlockEncrypt, KeyInit};
use aes::{Aes128, Aes192, Aes256};
use formula_office_crypto::{decrypt_standard_encrypted_package, OfficeCryptoError};
use md5::{Digest as _, Md5};
use sha1::Sha1;

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

fn hex_decode(mut s: &str) -> Vec<u8> {
    s = s.trim();
    let mut compact = String::new();
    let _ = compact.try_reserve(s.len());
    for ch in s.chars() {
        if ch.is_ascii_hexdigit() {
            compact.push(ch);
        }
    }
    assert_eq!(compact.len() % 2, 0, "hex string must have even length");
    let mut out = Vec::new();
    let _ = out.try_reserve_exact(compact.len() / 2);
    let bytes = compact.as_bytes();
    for i in (0..bytes.len()).step_by(2) {
        let hi = (bytes[i] as char).to_digit(16).unwrap();
        let lo = (bytes[i + 1] as char).to_digit(16).unwrap();
        out.push(((hi << 4) | lo) as u8);
    }
    out
}

fn password_to_utf16le(password: &str) -> Vec<u8> {
    let mut out = Vec::new();
    let _ = out.try_reserve(password.len().saturating_mul(2));
    for cu in password.encode_utf16() {
        out.extend_from_slice(&cu.to_le_bytes());
    }
    out
}

fn rc4_apply(key: &[u8], data: &mut [u8]) {
    assert!(!key.is_empty(), "RC4 key must be non-empty");
    let mut s = [0u8; 256];
    for (i, v) in s.iter_mut().enumerate() {
        *v = i as u8;
    }

    // KSA
    let mut j: u8 = 0;
    for i in 0..256u16 {
        let idx = i as usize;
        j = j.wrapping_add(s[idx]).wrapping_add(key[idx % key.len()]);
        s.swap(idx, j as usize);
    }

    // PRGA
    let mut i: u8 = 0;
    j = 0;
    for b in data {
        i = i.wrapping_add(1);
        j = j.wrapping_add(s[i as usize]);
        s.swap(i as usize, j as usize);
        let idx = s[i as usize].wrapping_add(s[j as usize]);
        *b ^= s[idx as usize];
    }
}

fn standard_rc4_spun_password_hash_sha1(password: &str, salt: &[u8]) -> [u8; 20] {
    let pw = password_to_utf16le(password);
    let mut hasher = Sha1::new();
    hasher.update(salt);
    hasher.update(&pw);
    let mut h: [u8; 20] = hasher.finalize().into();
    for i in 0..50_000u32 {
        let mut hasher = Sha1::new();
        hasher.update(i.to_le_bytes());
        hasher.update(h);
        h = hasher.finalize().into();
    }
    h
}

fn standard_rc4_derive_block_key_sha1(h: [u8; 20], block: u32, key_len: usize) -> Vec<u8> {
    let mut hasher = Sha1::new();
    hasher.update(h);
    hasher.update(block.to_le_bytes());
    let digest = hasher.finalize();
    digest[..key_len].to_vec()
}

fn standard_rc4_spun_password_hash_md5(password: &str, salt: &[u8]) -> [u8; 16] {
    let pw = password_to_utf16le(password);
    let mut hasher = Md5::new();
    hasher.update(salt);
    hasher.update(&pw);
    let mut h: [u8; 16] = hasher.finalize().into();
    for i in 0..50_000u32 {
        let mut hasher = Md5::new();
        hasher.update(i.to_le_bytes());
        hasher.update(h);
        h = hasher.finalize().into();
    }
    h
}

fn standard_rc4_derive_block_key_md5(h: [u8; 16], block: u32, key_len: usize) -> Vec<u8> {
    let mut hasher = Md5::new();
    hasher.update(h);
    hasher.update(block.to_le_bytes());
    let digest = hasher.finalize();
    digest[..key_len].to_vec()
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
fn standard_cryptoapi_rc4_sha1_vector_decrypts_package() {
    // Valid OOXML/ZIP payload to satisfy the crate's zip sanity checks.
    let plaintext = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/xlsx/basic/basic.xlsx"
    ));
    let password = "password";

    // Key derivation vector (docs/offcrypto-standard-cryptoapi-rc4.md).
    let h = standard_rc4_spun_password_hash_sha1(password, &SALT_00_TO_0F);
    assert_eq!(
        h.to_vec(),
        hex_decode("1b5972284eab6481eb6565a0985b334b3e65e041")
    );
    let key0 = standard_rc4_derive_block_key_sha1(h, 0, 16);
    assert_eq!(key0, hex_decode("6ad7dedf2da3514b1d85eabee069d47d"));

    // Build Standard/CryptoAPI RC4 EncryptionInfo.
    let version_major = 3u16;
    let version_minor = 2u16;
    let flags = 0x0000_0040u32;

    // EncryptionHeader (MS-OFFCRYPTO).
    let header_flags = 0x0000_0004u32; // fCryptoAPI (no fAES)
    let size_extra = 0u32;
    let alg_id = 0x0000_6801u32; // CALG_RC4
    let alg_id_hash = 0x0000_8004u32; // CALG_SHA1
    let key_bits = 128u32;
    let provider_type = 0x0000_0001u32; // PROV_RSA_FULL (typical for RC4)
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

    // EncryptionVerifier (RC4 uses unpadded encryptedVerifierHash length).
    let verifier_plain: [u8; 16] = *b"formula-rc4-test";
    let verifier_hash: [u8; 20] = Sha1::digest(&verifier_plain).into();
    let verifier_hash_size = verifier_hash.len() as u32;
    let mut verifier_buf = Vec::new();
    let _ = verifier_buf.try_reserve_exact(16usize.saturating_add(verifier_hash.len()));
    verifier_buf.extend_from_slice(&verifier_plain);
    verifier_buf.extend_from_slice(&verifier_hash);
    rc4_apply(&key0, &mut verifier_buf); // single continuous RC4 stream
    let encrypted_verifier = &verifier_buf[..16];
    let encrypted_verifier_hash = &verifier_buf[16..];

    let mut verifier_bytes = Vec::new();
    verifier_bytes.extend_from_slice(&(SALT_00_TO_0F.len() as u32).to_le_bytes());
    verifier_bytes.extend_from_slice(&SALT_00_TO_0F);
    verifier_bytes.extend_from_slice(encrypted_verifier);
    verifier_bytes.extend_from_slice(&verifier_hash_size.to_le_bytes());
    verifier_bytes.extend_from_slice(encrypted_verifier_hash);

    let mut encryption_info = Vec::new();
    encryption_info.extend_from_slice(&version_major.to_le_bytes());
    encryption_info.extend_from_slice(&version_minor.to_le_bytes());
    encryption_info.extend_from_slice(&flags.to_le_bytes());
    encryption_info.extend_from_slice(&header_size.to_le_bytes());
    encryption_info.extend_from_slice(&header_bytes);
    encryption_info.extend_from_slice(&verifier_bytes);

    // Build EncryptedPackage: 8-byte size prefix + RC4 ciphertext (0x200-byte blocks).
    let mut ciphertext = plaintext.to_vec();
    for (block_index, chunk) in ciphertext.chunks_mut(0x200).enumerate() {
        let key = standard_rc4_derive_block_key_sha1(h, block_index as u32, 16);
        rc4_apply(&key, chunk);
    }

    // Guard: ensure Standard/CryptoAPI uses 0x200-byte re-keying and not the BIFF8 0x400 interval.
    //
    // This avoids a potential "encrypt+decrypt with the same bug" false positive if this test
    // accidentally used a 0x400-byte interval while the decryptor also used 0x400.
    if plaintext.len() >= 0x400 && ciphertext.len() >= 0x400 {
        let key0 = standard_rc4_derive_block_key_sha1(h, 0, 16);
        let key1 = standard_rc4_derive_block_key_sha1(h, 1, 16);

        let mut block0_400 = plaintext[..0x400].to_vec();
        rc4_apply(&key0, &mut block0_400);

        let mut expected_block1 = plaintext[0x200..0x400].to_vec();
        rc4_apply(&key1, &mut expected_block1);

        assert_eq!(
            &ciphertext[0x200..0x400],
            expected_block1.as_slice(),
            "expected Standard RC4 to re-key at 0x200-byte boundary"
        );
        assert_ne!(
            &ciphertext[0x200..0x400],
            &block0_400[0x200..0x400],
            "ciphertext suggests incorrect 0x400-byte RC4 re-key interval"
        );
    }

    let mut encrypted_package = Vec::new();
    encrypted_package.extend_from_slice(&(plaintext.len() as u64).to_le_bytes());
    encrypted_package.extend_from_slice(&ciphertext);

    let decrypted =
        decrypt_standard_encrypted_package(&encryption_info, &encrypted_package, password)
            .expect("decrypt");
    assert_eq!(decrypted.as_slice(), plaintext);

    let err = decrypt_standard_encrypted_package(&encryption_info, &encrypted_package, "wrong-password")
        .expect_err("wrong password");
    assert!(matches!(err, OfficeCryptoError::InvalidPassword));
}

#[test]
fn standard_cryptoapi_rc4_md5_vector_decrypts_package() {
    // Valid OOXML/ZIP payload to satisfy the crate's zip sanity checks.
    let plaintext = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/xlsx/basic/basic.xlsx"
    ));
    let password = "password";

    // Key derivation vector (docs/offcrypto-standard-cryptoapi-rc4.md).
    let h = standard_rc4_spun_password_hash_md5(password, &SALT_00_TO_0F);
    assert_eq!(
        h.to_vec(),
        hex_decode("2079476089fda784c3a3cfeb98102c7e")
    );
    let key0 = standard_rc4_derive_block_key_md5(h, 0, 16);
    assert_eq!(key0, hex_decode("69badcae244868e209d4e053ccd2a3bc"));

    // Build Standard/CryptoAPI RC4 EncryptionInfo (MD5).
    let version_major = 3u16;
    let version_minor = 2u16;
    let flags = 0x0000_0040u32;

    let header_flags = 0x0000_0004u32; // fCryptoAPI (no fAES)
    let size_extra = 0u32;
    let alg_id = 0x0000_6801u32; // CALG_RC4
    let alg_id_hash = 0x0000_8003u32; // CALG_MD5
    let key_bits = 128u32;
    let provider_type = 0x0000_0001u32; // PROV_RSA_FULL
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

    // EncryptionVerifier (RC4 uses unpadded encryptedVerifierHash length; MD5 digest is 16 bytes).
    let verifier_plain: [u8; 16] = *b"formula-rc4-test";
    let verifier_hash: [u8; 16] = Md5::digest(&verifier_plain).into();
    let verifier_hash_size = verifier_hash.len() as u32;

    let mut verifier_buf = Vec::new();
    let _ = verifier_buf.try_reserve_exact(16usize.saturating_add(verifier_hash.len()));
    verifier_buf.extend_from_slice(&verifier_plain);
    verifier_buf.extend_from_slice(&verifier_hash);
    rc4_apply(&key0, &mut verifier_buf);
    let encrypted_verifier = &verifier_buf[..16];
    let encrypted_verifier_hash = &verifier_buf[16..];

    let mut verifier_bytes = Vec::new();
    verifier_bytes.extend_from_slice(&(SALT_00_TO_0F.len() as u32).to_le_bytes());
    verifier_bytes.extend_from_slice(&SALT_00_TO_0F);
    verifier_bytes.extend_from_slice(encrypted_verifier);
    verifier_bytes.extend_from_slice(&verifier_hash_size.to_le_bytes());
    verifier_bytes.extend_from_slice(encrypted_verifier_hash);

    let mut encryption_info = Vec::new();
    encryption_info.extend_from_slice(&version_major.to_le_bytes());
    encryption_info.extend_from_slice(&version_minor.to_le_bytes());
    encryption_info.extend_from_slice(&flags.to_le_bytes());
    encryption_info.extend_from_slice(&header_size.to_le_bytes());
    encryption_info.extend_from_slice(&header_bytes);
    encryption_info.extend_from_slice(&verifier_bytes);

    // Build EncryptedPackage: 8-byte size prefix + RC4 ciphertext (0x200-byte blocks).
    let mut ciphertext = plaintext.to_vec();
    for (block_index, chunk) in ciphertext.chunks_mut(0x200).enumerate() {
        let key = standard_rc4_derive_block_key_md5(h, block_index as u32, 16);
        rc4_apply(&key, chunk);
    }

    // Guard: ensure Standard/CryptoAPI uses 0x200-byte re-keying and not the BIFF8 0x400 interval.
    if plaintext.len() >= 0x400 && ciphertext.len() >= 0x400 {
        let key0 = standard_rc4_derive_block_key_md5(h, 0, 16);
        let key1 = standard_rc4_derive_block_key_md5(h, 1, 16);

        let mut block0_400 = plaintext[..0x400].to_vec();
        rc4_apply(&key0, &mut block0_400);

        let mut expected_block1 = plaintext[0x200..0x400].to_vec();
        rc4_apply(&key1, &mut expected_block1);

        assert_eq!(
            &ciphertext[0x200..0x400],
            expected_block1.as_slice(),
            "expected Standard RC4 to re-key at 0x200-byte boundary"
        );
        assert_ne!(
            &ciphertext[0x200..0x400],
            &block0_400[0x200..0x400],
            "ciphertext suggests incorrect 0x400-byte RC4 re-key interval"
        );
    }

    let mut encrypted_package = Vec::new();
    encrypted_package.extend_from_slice(&(plaintext.len() as u64).to_le_bytes());
    encrypted_package.extend_from_slice(&ciphertext);

    let decrypted =
        decrypt_standard_encrypted_package(&encryption_info, &encrypted_package, password)
            .expect("decrypt");
    assert_eq!(decrypted.as_slice(), plaintext);

    let err = decrypt_standard_encrypted_package(&encryption_info, &encrypted_package, "wrong-password")
        .expect_err("wrong password");
    assert!(matches!(err, OfficeCryptoError::InvalidPassword));
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

    // Regression guard: Standard AES fallback attempts an alternate key derivation; ensure a wrong
    // password still reports `InvalidPassword` (and not an "unsupported" error due to key length).
    let err =
        decrypt_standard_encrypted_package(&encryption_info, &encrypted_package, "wrong-password")
            .expect_err("wrong password");
    assert!(matches!(err, OfficeCryptoError::InvalidPassword));
}
