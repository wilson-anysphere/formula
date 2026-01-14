//! MS-OFFCRYPTO Standard / CryptoAPI / RC4 test vectors.
//!
//! These tests exist primarily as **developer documentation**: they encode one deterministic
//! reference vector for the Standard RC4 key derivation described in MS-OFFCRYPTO.
//!
//! See `docs/offcrypto-standard-cryptoapi-rc4.md` for the full writeup, including why the RC4
//! re-key interval is 0x200 bytes (and not the BIFF8 0x400-byte interval).

use md5::{Digest as _, Md5};
use sha1::Sha1;
use std::io::{Cursor, Read as _, Seek as _, SeekFrom};

use formula_io::offcrypto::cryptoapi::CALG_MD5;
use formula_io::{HashAlg, Rc4CryptoApiDecryptReader};

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

/// Standard CryptoAPI per-block key derivation helper (SHA-1).
///
/// Algorithm:
/// `h_block = SHA1(H || LE32(block_index))`, `key_material = h_block[0..key_len]` where
/// `key_len = keySize/8` (40→5 bytes, 56→7 bytes, 128→16 bytes).
///
/// Important: for **40-bit RC4** (`key_len == 5` / `keySize == 0`/`40`), CryptoAPI/Office do **not**
/// feed the raw 5-byte `key_material` into RC4. Instead they pad it to 16 bytes
/// (`key_material || 0x00 * 11`) before initializing RC4. That quirk is covered by
/// `standard_cryptoapi_rc4_40_bit_key_vector` below.
fn standard_rc4_derive_block_key(h: [u8; 20], block_index: u32, key_len: usize) -> Vec<u8> {
    let mut hasher = Sha1::new();
    hasher.update(h);
    hasher.update(block_index.to_le_bytes());
    let digest = hasher.finalize();
    digest[..key_len].to_vec()
}

/// Standard CryptoAPI "spun password hash" helper for MD5.
fn standard_rc4_spun_password_hash_md5(password: &str, salt: &[u8], spin_count: u32) -> [u8; 16] {
    let pw = password_utf16le_bytes(password);
    let mut h = Md5::new();
    h.update(salt);
    h.update(&pw);
    let mut cur: [u8; 16] = h.finalize().into();

    for i in 0..spin_count {
        let mut h = Md5::new();
        h.update(i.to_le_bytes());
        h.update(cur);
        cur = h.finalize().into();
    }

    cur
}

/// Standard CryptoAPI per-block key derivation helper (MD5).
fn standard_rc4_derive_block_key_md5(h: [u8; 16], block_index: u32, key_len: usize) -> Vec<u8> {
    let mut hasher = Md5::new();
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
    let key2 = standard_rc4_derive_block_key(h, 2, key_len);
    let key3 = standard_rc4_derive_block_key(h, 3, key_len);
    assert_eq!(key0, hex_decode("6ad7dedf2da3514b1d85eabee069d47d"));
    assert_eq!(key1, hex_decode("2ed4e8825cd48aa4a47994cda7415b4a"));
    assert_eq!(key2, hex_decode("9ce57d0699be3938951f47fa949361db"));
    assert_eq!(key3, hex_decode("e65b2643eaba3815a37a61159f137840"));

    // Key material for 40-bit and 56-bit keys is a raw truncation of `SHA1(H || LE32(block))`.
    // (CryptoAPI applies an additional padding quirk for 40-bit RC4 when initializing RC4.)
    let key0_40 = standard_rc4_derive_block_key(h, 0, 5);
    let key1_40 = standard_rc4_derive_block_key(h, 1, 5);
    assert_eq!(key0_40, hex_decode("6ad7dedf2d"));
    assert_eq!(key1_40, hex_decode("2ed4e8825c"));

    let key0_56 = standard_rc4_derive_block_key(h, 0, 7);
    let key1_56 = standard_rc4_derive_block_key(h, 1, 7);
    assert_eq!(key0_56, hex_decode("6ad7dedf2da351"));
    assert_eq!(key1_56, hex_decode("2ed4e8825cd48a"));

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

#[test]
fn standard_cryptoapi_rc4_derivation_md5_vector() {
    let password = "password";
    let salt: Vec<u8> = (0u8..=0x0F).collect();
    let spin_count = 50_000u32;
    let key_len = 16usize;

    let h = standard_rc4_spun_password_hash_md5(password, &salt, spin_count);
    assert_eq!(h.to_vec(), hex_decode("2079476089fda784c3a3cfeb98102c7e"));

    let key0 = standard_rc4_derive_block_key_md5(h, 0, key_len);
    let key1 = standard_rc4_derive_block_key_md5(h, 1, key_len);
    let key2 = standard_rc4_derive_block_key_md5(h, 2, key_len);
    let key3 = standard_rc4_derive_block_key_md5(h, 3, key_len);

    // Sanity: different block indexes must yield different keys.
    assert_ne!(key0, key1);

    assert_eq!(key0, hex_decode("69badcae244868e209d4e053ccd2a3bc"));
    assert_eq!(key1, hex_decode("6f4d502ab37700ffdab5704160455b47"));
    assert_eq!(key2, hex_decode("ac69022e396c7750872133f37e2c7afc"));
    assert_eq!(key3, hex_decode("1b056e7118ab8d35e9d67adee8b11104"));

    // Key material for 40-bit and 56-bit keys is a raw truncation of `MD5(H || LE32(block))`.
    // (CryptoAPI applies an additional padding quirk for 40-bit RC4 when initializing RC4.)
    let key0_40 = standard_rc4_derive_block_key_md5(h, 0, 5);
    let key1_40 = standard_rc4_derive_block_key_md5(h, 1, 5);
    assert_eq!(key0_40, hex_decode("69badcae24"));
    assert_eq!(key1_40, hex_decode("6f4d502ab3"));

    let key0_56 = standard_rc4_derive_block_key_md5(h, 0, 7);
    let key1_56 = standard_rc4_derive_block_key_md5(h, 1, 7);
    assert_eq!(key0_56, hex_decode("69badcae244868"));
    assert_eq!(key1_56, hex_decode("6f4d502ab37700"));

    let plaintext = b"Hello, RC4 CryptoAPI!";
    let ciphertext = rc4_apply(&key0, plaintext);
    assert_eq!(
        ciphertext,
        hex_decode("425dd9c8165e1216065e53eb586e897b5e85a07a6d")
    );
    assert_eq!(rc4_apply(&key0, &ciphertext), plaintext);
}

#[test]
fn standard_cryptoapi_rc4_40_bit_key_vector() {
    let password = "password";
    let salt: Vec<u8> = (0u8..=0x0F).collect();
    let spin_count = 50_000u32;
    let key_len = 5usize; // 40-bit

    let h = standard_rc4_spun_password_hash(password, &salt, spin_count);

    // Hb = SHA1(H || LE32(0))
    let mut hasher = Sha1::new();
    hasher.update(h);
    hasher.update(0u32.to_le_bytes());
    let hb: [u8; 20] = hasher.finalize().into();
    assert_eq!(
        hb.to_vec(),
        hex_decode("6ad7dedf2da3514b1d85eabee069d47dd058967f")
    );

    let key_material = hb[..5].to_vec();
    assert_eq!(key_material, hex_decode("6ad7dedf2d"));

    // CryptoAPI/Office represent a "40-bit" RC4 key as a 16-byte RC4 key where the high 88 bits are
    // zero. RC4's KSA depends on both the key bytes and the key length, so treating the 40-bit key
    // material as a raw 5-byte key produces a different keystream than CryptoAPI/Office.
    let mut padded_key = key_material.clone();
    padded_key.resize(16, 0);
    assert_eq!(padded_key, hex_decode("6ad7dedf2d0000000000000000000000"));

    let plaintext = b"Hello, RC4 CryptoAPI!";

    // CryptoAPI padded 16-byte key (correct behavior).
    let ciphertext_padded = rc4_apply(&padded_key, plaintext);
    assert_eq!(
        ciphertext_padded,
        hex_decode("7a8bd000713a6e30ba9916476d27b01d36707a6ef8")
    );

    // Regression guard: the unpadded 5-byte key produces a different keystream/ciphertext.
    let ciphertext_unpadded = rc4_apply(&key_material, plaintext);
    assert_eq!(
        ciphertext_unpadded,
        hex_decode("d1fa444913b4839b06eb4851750a07761005f025bf")
    );
    assert_ne!(ciphertext_padded, ciphertext_unpadded);

    // Ensure the production decrypt reader uses the CryptoAPI padded key form.
    let mut stream = Vec::new();
    stream.extend_from_slice(&(plaintext.len() as u64).to_le_bytes());
    stream.extend_from_slice(&ciphertext_padded);

    let mut cursor = Cursor::new(stream);
    cursor.seek(SeekFrom::Start(8)).unwrap();

    let mut reader =
        Rc4CryptoApiDecryptReader::new(cursor, plaintext.len() as u64, h.to_vec(), key_len).unwrap();
    let mut decrypted = Vec::new();
    reader.read_to_end(&mut decrypted).unwrap();
    assert_eq!(decrypted, plaintext);
}

#[test]
fn standard_cryptoapi_rc4_md5_40_bit_key_vector() {
    let password = "password";
    let salt: Vec<u8> = (0u8..=0x0F).collect();
    let spin_count = 50_000u32;
    let key_len = 5usize; // 40-bit (keySize/8)

    // H = MD5(salt || UTF16LE(password)), spun 50k times.
    let h = standard_rc4_spun_password_hash_md5(password, &salt, spin_count);

    // Hb0 = MD5(H || LE32(0)).
    let mut hasher = Md5::new();
    hasher.update(h);
    hasher.update(0u32.to_le_bytes());
    let hb0: [u8; 16] = hasher.finalize().into();
    assert_eq!(
        hb0.to_vec(),
        hex_decode("69badcae244868e209d4e053ccd2a3bc")
    );

    // The 40-bit RC4 key is the first 5 bytes of Hb0. (Padding this to 16 bytes is a common legacy
    // behavior that changes the RC4 KSA and produces a different keystream.)
    let key_material = hb0[..5].to_vec();
    assert_eq!(key_material, hex_decode("69badcae24"));

    let mut padded_key = key_material.clone();
    padded_key.resize(16, 0);
    assert_eq!(padded_key, hex_decode("69badcae240000000000000000000000"));

    let plaintext = b"Hello, RC4 CryptoAPI!";

    // Spec-correct 5-byte key.
    let ciphertext_unpadded = rc4_apply(&key_material, plaintext);
    assert_eq!(
        ciphertext_unpadded,
        hex_decode("db037cd60d38c882019b5f5d8c43382373f476da28")
    );

    // Demonstrate that (incorrect) zero-padding changes ciphertext.
    let ciphertext_padded = rc4_apply(&padded_key, plaintext);
    assert_eq!(
        ciphertext_padded,
        hex_decode("565016a3d8158632bb36ce1d76996628512061bfa3")
    );
    assert_ne!(ciphertext_unpadded, ciphertext_padded);

    // Ensure the production decrypt reader uses:
    // - MD5 for per-block key derivation
    // - the unpadded 5-byte RC4 key when keySize==0/40 (keyLen=5)
    let mut stream = Vec::new();
    stream.extend_from_slice(&(plaintext.len() as u64).to_le_bytes());
    stream.extend_from_slice(&ciphertext_unpadded);

    // Exercise the production framing + parameter-validation path.
    let cursor = Cursor::new(stream.clone());
    let mut reader =
        Rc4CryptoApiDecryptReader::from_encrypted_package_stream(cursor, h.to_vec(), 0, CALG_MD5)
            .unwrap();
    let mut decrypted = Vec::new();
    reader.read_to_end(&mut decrypted).unwrap();
    assert_eq!(decrypted, plaintext);

    // Also exercise the lower-level constructor that assumes the caller has already parsed the
    // `EncryptedPackage` size prefix.
    let mut cursor = Cursor::new(stream);
    cursor.seek(SeekFrom::Start(8)).unwrap();
    let mut reader = Rc4CryptoApiDecryptReader::new_with_hash_alg(
        cursor,
        plaintext.len() as u64,
        h.to_vec(),
        key_len,
        HashAlg::Md5,
    )
    .unwrap();
    let mut decrypted = Vec::new();
    reader.read_to_end(&mut decrypted).unwrap();
    assert_eq!(decrypted, plaintext);
}
