//! MS-OFFCRYPTO Standard / CryptoAPI / RC4 test vectors.
//!
//! These tests exist primarily as **developer documentation**: they encode one deterministic
//! reference vector for the Standard RC4 key derivation described in MS-OFFCRYPTO.
//!
//! See `docs/offcrypto-standard-cryptoapi-rc4.md` for the full writeup, including why the RC4
//! re-key interval is 0x200 bytes (and not the BIFF8 0x400-byte interval).

use sha1::{Digest as _, Sha1};

fn hex_decode(mut s: &str) -> Vec<u8> {
    // Keep parsing permissive for readability in expected-value literals.
    s = s.trim();
    let mut compact = String::with_capacity(s.len());
    for ch in s.chars() {
        if ch.is_ascii_hexdigit() {
            compact.push(ch);
        }
    }
    assert!(
        compact.len() % 2 == 0,
        "hex string must have even length (got {})",
        compact.len()
    );

    let mut out = Vec::with_capacity(compact.len() / 2);
    let bytes = compact.as_bytes();
    for i in (0..bytes.len()).step_by(2) {
        let hi = (bytes[i] as char).to_digit(16).unwrap();
        let lo = (bytes[i + 1] as char).to_digit(16).unwrap();
        out.push(((hi << 4) | lo) as u8);
    }
    out
}

fn password_utf16le_bytes(password: &str) -> Vec<u8> {
    // UTF-16LE with no BOM and no terminator.
    let mut out = Vec::with_capacity(password.len().saturating_mul(2));
    for cu in password.encode_utf16() {
        out.extend_from_slice(&cu.to_le_bytes());
    }
    out
}

/// Standard CryptoAPI "spun password hash" helper.
///
/// Algorithm:
/// 1. `H = SHA1(salt || UTF-16LE(password))`
/// 2. For `i in 0..spinCount`: `H = SHA1(LE32(i) || H)`
fn standard_rc4_spun_password_hash(password: &str, salt: &[u8], spin_count: u32) -> [u8; 20] {
    let pw = password_utf16le_bytes(password);
    let mut h = Sha1::new();
    h.update(salt);
    h.update(&pw);
    let mut cur: [u8; 20] = h.finalize().into();

    for i in 0..spin_count {
        let mut h = Sha1::new();
        h.update(i.to_le_bytes());
        h.update(cur);
        cur = h.finalize().into();
    }

    cur
}

/// Standard CryptoAPI per-block key derivation helper.
///
/// Algorithm:
/// `rc4_key = SHA1(H || LE32(block_index))[0..key_len]`
fn standard_rc4_derive_block_key(h: [u8; 20], block_index: u32, key_len: usize) -> Vec<u8> {
    let mut hasher = Sha1::new();
    hasher.update(h);
    hasher.update(block_index.to_le_bytes());
    let digest = hasher.finalize();
    digest[..key_len].to_vec()
}

/// Minimal RC4 implementation (KSA + PRGA).
fn rc4_apply(key: &[u8], data: &[u8]) -> Vec<u8> {
    assert!(!key.is_empty(), "RC4 key must be non-empty");

    let mut s = [0u8; 256];
    for (i, b) in s.iter_mut().enumerate() {
        *b = i as u8;
    }

    let mut j: u8 = 0;
    for i in 0..256u16 {
        let i8 = i as u8;
        j = j
            .wrapping_add(s[i as usize])
            .wrapping_add(key[i as usize % key.len()]);
        s.swap(i as usize, j as usize);
        // `i` increments in u16 to avoid accidental overflow logic changes.
        let _ = i8;
    }

    let mut i: u8 = 0;
    j = 0;
    let mut out = Vec::with_capacity(data.len());
    for &b in data {
        i = i.wrapping_add(1);
        j = j.wrapping_add(s[i as usize]);
        s.swap(i as usize, j as usize);
        let k = s[(s[i as usize].wrapping_add(s[j as usize])) as usize];
        out.push(b ^ k);
    }
    out
}

#[test]
fn rc4_wikipedia_vector() {
    // Well-known RC4 test vector (from Wikipedia / many implementations):
    // key="Key", plaintext="Plaintext"
    let key = b"Key";
    let plaintext = b"Plaintext";
    let expected = hex_decode("BBF316E8D940AF0AD3");
    assert_eq!(rc4_apply(key, plaintext), expected);
}

#[test]
fn standard_cryptoapi_rc4_derivation_vector() {
    let password = "password";
    let salt: Vec<u8> = (0u8..=0x0F).collect();
    let spin_count = 50_000u32;
    let key_len = 16usize;

    let h = standard_rc4_spun_password_hash(password, &salt, spin_count);
    assert_eq!(
        h.to_vec(),
        hex_decode("1b5972284eab6481eb6565a0985b334b3e65e041")
    );

    let key0 = standard_rc4_derive_block_key(h, 0, key_len);
    let key1 = standard_rc4_derive_block_key(h, 1, key_len);
    assert_eq!(key0, hex_decode("6ad7dedf2da3514b1d85eabee069d47d"));
    assert_eq!(key1, hex_decode("2ed4e8825cd48aa4a47994cda7415b4a"));

    // Sanity: different block indexes must yield different keys.
    assert_ne!(key0, key1);

    let plaintext = b"Hello, RC4 CryptoAPI!";
    let ciphertext = rc4_apply(&key0, plaintext);
    assert_eq!(
        ciphertext,
        hex_decode("e7c9974140e69857dbdec656c7ccb4f9283d723236")
    );
    assert_eq!(rc4_apply(&key0, &ciphertext), plaintext);
}
