//! Deterministic test vectors for MS-OFFCRYPTO Standard (CryptoAPI) MD5 primitives.
//!
//! We already have end-to-end encrypted workbook fixtures, but MD5-based CryptoAPI files are less
//! common than SHA-1. These vectors ensure the MD5 KDF implementation stays correct across
//! platforms (and don't rely on Windows CryptoAPI availability / FIPS policy).

use formula_io::offcrypto::cryptoapi::{
    crypt_derive_key, final_hash, hash_password_fixed_spin, password_to_utf16le, HashAlg,
};

fn hex_bytes(s: &str) -> Vec<u8> {
    let s = s.trim();
    assert!(
        s.len() % 2 == 0,
        "hex string must have even length (got {})",
        s.len()
    );
    let mut out = Vec::with_capacity(s.len() / 2);
    let bytes = s.as_bytes();
    for i in (0..bytes.len()).step_by(2) {
        let hi = (bytes[i] as char).to_digit(16).expect("hex digit");
        let lo = (bytes[i + 1] as char).to_digit(16).expect("hex digit");
        out.push(((hi << 4) | lo) as u8);
    }
    out
}

#[test]
fn cryptoapi_md5_fixed_spin_and_derive_key_vectors() {
    let password = "password";
    let salt: Vec<u8> = (0u8..16).collect();

    let password_utf16le = password_to_utf16le(password);
    let h = hash_password_fixed_spin(&password_utf16le, &salt, HashAlg::Md5);
    assert_eq!(h, hex_bytes("2079476089fda784c3a3cfeb98102c7e"));

    let expected_hfinal = [
        "69badcae244868e209d4e053ccd2a3bc",
        "6f4d502ab37700ffdab5704160455b47",
        "ac69022e396c7750872133f37e2c7afc",
        "1b056e7118ab8d35e9d67adee8b11104",
    ];
    let expected_cryptderivekey = [
        "8d666ec55103fdbdc3281cc271f6cb7c",
        "892b60ddd451139fed758f20fe5d1be0",
        "d9034198455f9bd171ad16d04cea4c42",
        "06f5756e6e23c795cd6786f5dd565830",
    ];

    for (block, (expected_hfinal, expected_cryptderivekey)) in expected_hfinal
        .iter()
        .zip(expected_cryptderivekey.iter())
        .enumerate()
    {
        let hfinal = final_hash(&h, block as u32, HashAlg::Md5);
        assert_eq!(
            hfinal,
            hex_bytes(expected_hfinal),
            "MD5 final hash mismatch for block {block}"
        );

        // AES session keys derived via CryptoAPI's `CryptDeriveKey` ipad/opad expansion.
        let key = crypt_derive_key(&hfinal, 16, HashAlg::Md5);
        assert_eq!(
            key,
            hex_bytes(expected_cryptderivekey),
            "MD5 CryptDeriveKey mismatch for block {block}"
        );
    }
}

