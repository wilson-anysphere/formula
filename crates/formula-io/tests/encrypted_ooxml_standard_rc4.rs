//! End-to-end fixture test for MS-OFFCRYPTO Standard / CryptoAPI / RC4 encrypted OOXML.
//!
//! This exercises the full parsing + verifier validation + `EncryptedPackage` decryption path
//! (including the **0x200-byte** per-block RC4 re-key interval).

use std::io::{Cursor, Read as _, Seek as _, SeekFrom};
use std::path::PathBuf;

use formula_io::Rc4CryptoApiDecryptReader;
use sha1::{Digest as _, Sha1};
use sha2::Sha256;

const CALG_RC4: u32 = 0x0000_6801;
const CALG_SHA1: u32 = 0x0000_8004;
const SPIN_COUNT: u32 = 50_000;

fn fixture_path(rel: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/encrypted/ooxml")
        .join(rel)
}

fn open_stream_case_tolerant<R: std::io::Seek>(
    ole: &mut cfb::CompoundFile<R>,
    name: &str,
) -> std::io::Result<cfb::Stream<R>> {
    ole.open_stream(name)
        .or_else(|_| ole.open_stream(format!("/{name}")))
}

#[derive(Debug)]
struct StandardRc4EncryptionInfo {
    alg_id: u32,
    alg_id_hash: u32,
    key_size_bits: u32,
    salt: [u8; 16],
    encrypted_verifier: [u8; 16],
    verifier_hash_size: u32,
    encrypted_verifier_hash: Vec<u8>,
}

fn read_u16_le(bytes: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes(bytes[offset..offset + 2].try_into().unwrap())
}

fn read_u32_le(bytes: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes(bytes[offset..offset + 4].try_into().unwrap())
}

fn parse_standard_rc4_encryption_info(bytes: &[u8]) -> StandardRc4EncryptionInfo {
    assert!(
        bytes.len() >= 8 + 4,
        "EncryptionInfo stream too short (len={})",
        bytes.len()
    );

    let major = read_u16_le(bytes, 0);
    let minor = read_u16_le(bytes, 2);
    assert!(
        minor == 2 && matches!(major, 2 | 3 | 4),
        "expected Standard EncryptionInfo version *.2 with major=2/3/4, got {major}.{minor}"
    );

    let header_size = read_u32_le(bytes, 8) as usize;
    let header_start = 12usize;
    let header_end = header_start + header_size;
    assert!(
        header_end <= bytes.len(),
        "EncryptionHeader out of range (header_size={header_size}, stream_len={})",
        bytes.len()
    );

    let header = &bytes[header_start..header_end];
    assert!(
        header.len() >= 32,
        "EncryptionHeader too short (len={})",
        header.len()
    );
    let alg_id = read_u32_le(header, 8);
    let alg_id_hash = read_u32_le(header, 12);
    let key_size_bits = read_u32_le(header, 16);

    // EncryptionVerifier starts immediately after the header.
    let mut pos = header_end;
    assert!(
        pos + 4 <= bytes.len(),
        "EncryptionVerifier truncated at saltSize"
    );
    let salt_size = read_u32_le(bytes, pos) as usize;
    pos += 4;
    assert_eq!(salt_size, 16, "expected 16-byte salt");
    assert!(pos + 16 <= bytes.len(), "EncryptionVerifier.salt truncated");
    let salt: [u8; 16] = bytes[pos..pos + 16].try_into().unwrap();
    pos += 16;

    assert!(
        pos + 16 <= bytes.len(),
        "EncryptionVerifier.encryptedVerifier truncated"
    );
    let encrypted_verifier: [u8; 16] = bytes[pos..pos + 16].try_into().unwrap();
    pos += 16;

    assert!(
        pos + 4 <= bytes.len(),
        "EncryptionVerifier.verifierHashSize truncated"
    );
    let verifier_hash_size = read_u32_le(bytes, pos);
    pos += 4;
    assert_eq!(
        verifier_hash_size, 20,
        "expected verifierHashSize=20 for SHA1"
    );

    let encrypted_verifier_hash = bytes[pos..].to_vec();
    assert!(
        encrypted_verifier_hash.len() >= verifier_hash_size as usize,
        "encryptedVerifierHash too short (len={})",
        encrypted_verifier_hash.len()
    );

    StandardRc4EncryptionInfo {
        alg_id,
        alg_id_hash,
        key_size_bits,
        salt,
        encrypted_verifier,
        verifier_hash_size,
        encrypted_verifier_hash,
    }
}

fn password_utf16le_bytes(password: &str) -> Vec<u8> {
    let mut out = Vec::with_capacity(password.len().saturating_mul(2));
    for cu in password.encode_utf16() {
        out.extend_from_slice(&cu.to_le_bytes());
    }
    out
}

fn spun_password_hash_sha1(password: &str, salt: &[u8]) -> [u8; 20] {
    let pw = password_utf16le_bytes(password);

    // h = SHA1(salt || pw)
    let mut hasher = Sha1::new();
    hasher.update(salt);
    hasher.update(&pw);
    let mut h: [u8; 20] = hasher.finalize().into();

    // for i in 0..SPIN_COUNT-1: h = SHA1(LE32(i) || h)
    let mut buf = [0u8; 4 + 20];
    for i in 0..SPIN_COUNT {
        buf[..4].copy_from_slice(&i.to_le_bytes());
        buf[4..].copy_from_slice(&h);
        h = Sha1::digest(buf).into();
    }

    h
}

fn derive_block_key_sha1(h: &[u8; 20], block_index: u32, key_len: usize) -> Vec<u8> {
    assert!(key_len <= 20);
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
        let si = s[i as usize];
        j = j.wrapping_add(si).wrapping_add(key[i as usize % key.len()]);
        s.swap(i as usize, j as usize);
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

fn verify_password(info: &StandardRc4EncryptionInfo, password: &str) -> Result<[u8; 20], String> {
    if info.alg_id != CALG_RC4 {
        return Err(format!("expected CALG_RC4, got 0x{:08X}", info.alg_id));
    }
    if info.alg_id_hash != CALG_SHA1 {
        return Err(format!(
            "expected CALG_SHA1, got 0x{:08X}",
            info.alg_id_hash
        ));
    }

    let key_len = usize::try_from(info.key_size_bits / 8).map_err(|_| "keySize overflow")?;
    if key_len == 0 || key_len > 20 || info.key_size_bits % 8 != 0 {
        return Err(format!("invalid keySizeBits={}", info.key_size_bits));
    }

    let h = spun_password_hash_sha1(password, &info.salt);
    let key0 = derive_block_key_sha1(&h, 0, key_len);

    // Decrypt verifier + verifierHash as a single RC4 stream.
    let mut ciphertext = Vec::new();
    ciphertext.extend_from_slice(&info.encrypted_verifier);
    ciphertext.extend_from_slice(&info.encrypted_verifier_hash);
    let plain = rc4_apply(&key0, &ciphertext);

    if plain.len() < 16 + info.verifier_hash_size as usize {
        return Err("decrypted verifier payload is too short".to_string());
    }

    let verifier = &plain[..16];
    let verifier_hash = &plain[16..16 + info.verifier_hash_size as usize];
    let expected = Sha1::digest(verifier);
    if verifier_hash != expected.as_slice() {
        return Err("verifier hash mismatch".to_string());
    }

    Ok(h)
}

#[test]
fn decrypts_standard_cryptoapi_rc4_fixture_and_rejects_wrong_password() {
    let encrypted_path = fixture_path("standard-rc4.xlsx");
    let plaintext_path = fixture_path("plaintext.xlsx");

    assert!(
        encrypted_path.exists(),
        "missing fixture {}",
        encrypted_path.display()
    );
    assert!(
        plaintext_path.exists(),
        "missing fixture {}",
        plaintext_path.display()
    );

    let file = std::fs::File::open(&encrypted_path).expect("open encrypted fixture");
    let mut ole = cfb::CompoundFile::open(file).expect("open OLE container");

    // Read EncryptionInfo and parse header/verifier fields.
    let mut encryption_info_bytes = Vec::new();
    open_stream_case_tolerant(&mut ole, "EncryptionInfo")
        .expect("open EncryptionInfo")
        .read_to_end(&mut encryption_info_bytes)
        .expect("read EncryptionInfo");
    let info = parse_standard_rc4_encryption_info(&encryption_info_bytes);
    assert_eq!(info.alg_id, CALG_RC4);
    assert_eq!(info.alg_id_hash, CALG_SHA1);

    // Wrong password must fail verifier check.
    assert!(
        verify_password(&info, "wrong-password").is_err(),
        "expected wrong password verifier to fail"
    );

    // Correct password.
    let h = verify_password(&info, "password").expect("password verifier should succeed");
    let key_len = (info.key_size_bits / 8) as usize;

    // Read EncryptedPackage stream.
    let mut encrypted_package_bytes = Vec::new();
    open_stream_case_tolerant(&mut ole, "EncryptedPackage")
        .expect("open EncryptedPackage")
        .read_to_end(&mut encrypted_package_bytes)
        .expect("read EncryptedPackage");
    assert!(
        encrypted_package_bytes.len() >= 8,
        "EncryptedPackage is too short"
    );

    let package_size = u64::from_le_bytes(encrypted_package_bytes[..8].try_into().unwrap());
    let mut cursor = Cursor::new(encrypted_package_bytes);
    cursor.seek(SeekFrom::Start(8)).unwrap();

    // Decrypt using the library's 0x200-block RC4 decrypt reader.
    let mut reader = Rc4CryptoApiDecryptReader::new(cursor, package_size, h.to_vec(), key_len)
        .expect("create RC4 decrypt reader");
    let mut decrypted = Vec::new();
    reader
        .read_to_end(&mut decrypted)
        .expect("read decrypted package");

    assert_eq!(
        decrypted.len() as u64,
        package_size,
        "decrypted size should match EncryptedPackage header"
    );
    assert!(
        decrypted.starts_with(b"PK\x03\x04"),
        "decrypted package should be a ZIP (missing PK\\x03\\x04 signature)"
    );

    let expected = std::fs::read(&plaintext_path).expect("read expected plaintext fixture");
    let decrypted_sha = Sha256::digest(&decrypted);
    let expected_sha = Sha256::digest(&expected);
    assert_eq!(
        decrypted_sha.as_slice(),
        expected_sha.as_slice(),
        "SHA256 mismatch (decrypted package bytes differ from plaintext.xlsx)"
    );
}
