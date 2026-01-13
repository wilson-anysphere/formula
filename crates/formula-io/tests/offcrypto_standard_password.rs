use std::io::Read;
use std::path::PathBuf;

use aes::cipher::{BlockDecrypt, KeyInit};
use sha1::{Digest, Sha1};

const FIXTURE_PASSWORD: &str = "password";

/// Fixture produced with Apache POI (standard encryption / CryptoAPI) and then patched to use
/// `EncryptionInfo` version `3.2` (Office 2007-style Standard encryption).
///
/// The workbook contains a single sheet ("Sheet1") with cell A1="hello".
fn fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures/offcrypto_standard_cryptoapi_password.xlsx")
}

#[derive(Debug)]
struct StandardEncryptionInfo {
    version_major: u16,
    version_minor: u16,
    flags: u32,
    header: StandardEncryptionHeader,
    verifier: StandardEncryptionVerifier,
}

#[derive(Debug)]
struct StandardEncryptionHeader {
    #[allow(dead_code)]
    flags: u32,
    alg_id: u32,
    alg_id_hash: u32,
    key_size_bits: u32,
}

#[derive(Debug)]
struct StandardEncryptionVerifier {
    salt: [u8; 16],
    encrypted_verifier: [u8; 16],
    verifier_hash_size: u32,
    encrypted_verifier_hash: Vec<u8>,
}

fn read_u16_le(bytes: &[u8], offset: &mut usize) -> Result<u16, String> {
    let b = bytes
        .get(*offset..*offset + 2)
        .ok_or_else(|| "unexpected EOF while reading u16".to_string())?;
    *offset += 2;
    Ok(u16::from_le_bytes([b[0], b[1]]))
}

fn read_u32_le(bytes: &[u8], offset: &mut usize) -> Result<u32, String> {
    let b = bytes
        .get(*offset..*offset + 4)
        .ok_or_else(|| "unexpected EOF while reading u32".to_string())?;
    *offset += 4;
    Ok(u32::from_le_bytes([b[0], b[1], b[2], b[3]]))
}

fn read_exact<const N: usize>(bytes: &[u8], offset: &mut usize) -> Result<[u8; N], String> {
    let b = bytes
        .get(*offset..*offset + N)
        .ok_or_else(|| format!("unexpected EOF while reading [{N}]"))?;
    *offset += N;
    let mut out = [0u8; N];
    out.copy_from_slice(b);
    Ok(out)
}

fn parse_standard_encryption_info(bytes: &[u8]) -> Result<StandardEncryptionInfo, String> {
    let mut off = 0usize;
    let version_major = read_u16_le(bytes, &mut off)?;
    let version_minor = read_u16_le(bytes, &mut off)?;
    let flags = read_u32_le(bytes, &mut off)?;

    // [MS-OFFCRYPTO] 2.3.4.5:
    //   - u32 headerSize
    //   - EncryptionHeader (headerSize bytes)
    //   - EncryptionVerifier
    let header_size = read_u32_le(bytes, &mut off)? as usize;
    let header_end = off
        .checked_add(header_size)
        .ok_or_else(|| "header size overflow".to_string())?;
    let header_bytes = bytes
        .get(off..header_end)
        .ok_or_else(|| "EncryptionHeader truncated".to_string())?;
    off = header_end;

    let mut h_off = 0usize;
    let header_flags = read_u32_le(header_bytes, &mut h_off)?;
    let _size_extra = read_u32_le(header_bytes, &mut h_off)?;
    let alg_id = read_u32_le(header_bytes, &mut h_off)?;
    let alg_id_hash = read_u32_le(header_bytes, &mut h_off)?;
    let key_size_bits = read_u32_le(header_bytes, &mut h_off)?;
    // Remaining header fields (providerType + reserved + CSPName) are not needed for password
    // verification.

    let header = StandardEncryptionHeader {
        flags: header_flags,
        alg_id,
        alg_id_hash,
        key_size_bits,
    };

    let salt_size = read_u32_le(bytes, &mut off)?;
    if salt_size != 16 {
        return Err(format!("unexpected salt size {salt_size} (expected 16)"));
    }
    let salt = read_exact::<16>(bytes, &mut off)?;
    let encrypted_verifier = read_exact::<16>(bytes, &mut off)?;
    let verifier_hash_size = read_u32_le(bytes, &mut off)?;
    if verifier_hash_size != 20 {
        // Standard encryption uses SHA-1; tolerate other sizes, but keep the fixture strict.
        return Err(format!(
            "unexpected verifierHashSize {verifier_hash_size} (expected 20 for SHA-1)"
        ));
    }

    // [MS-OFFCRYPTO] 2.3.4.9: encrypted verifier hash is always 32 bytes (AES block aligned),
    // even though the decrypted hash is 20 bytes.
    let encrypted_verifier_hash = bytes
        .get(off..off + 32)
        .ok_or_else(|| "encryptedVerifierHash truncated".to_string())?
        .to_vec();
    off += 32;

    if off != bytes.len() {
        return Err(format!(
            "unexpected trailing bytes in EncryptionInfo: {}",
            bytes.len() - off
        ));
    }

    let verifier = StandardEncryptionVerifier {
        salt,
        encrypted_verifier,
        verifier_hash_size,
        encrypted_verifier_hash,
    };

    Ok(StandardEncryptionInfo {
        version_major,
        version_minor,
        flags,
        header,
        verifier,
    })
}

fn sha1_bytes(data: &[u8]) -> [u8; 20] {
    let mut hasher = Sha1::new();
    hasher.update(data);
    let digest = hasher.finalize();
    let mut out = [0u8; 20];
    out.copy_from_slice(&digest);
    out
}

fn hash_password_sha1(password: &str, salt: &[u8; 16], spin_count: u32) -> [u8; 20] {
    // [MS-OFFCRYPTO] Standard encryption key derivation matches Apache POI's `hashPassword` helper:
    //
    //   H0 = SHA1(salt || password_utf16le)
    //   Hi = SHA1(i_le || H_{i-1}) for i in 0..spinCount-1
    let mut pw_utf16le = Vec::with_capacity(password.len() * 2);
    for ch in password.encode_utf16() {
        pw_utf16le.extend_from_slice(&ch.to_le_bytes());
    }

    let mut h0 = Sha1::new();
    h0.update(salt);
    h0.update(&pw_utf16le);
    let digest = h0.finalize();
    let mut hash = [0u8; 20];
    hash.copy_from_slice(&digest);

    for i in 0..spin_count {
        let mut h = Sha1::new();
        h.update(i.to_le_bytes());
        h.update(hash);
        let digest = h.finalize();
        hash.copy_from_slice(&digest);
    }

    hash
}

fn fill_and_xor_then_sha1(hash: &[u8; 20], fill_byte: u8) -> [u8; 20] {
    let mut buff = [fill_byte; 64];
    for (i, b) in hash.iter().enumerate() {
        buff[i] ^= b;
    }
    sha1_bytes(&buff)
}

fn derive_standard_key(
    password: &str,
    verifier: &StandardEncryptionVerifier,
    key_size_bytes: usize,
) -> Vec<u8> {
    // Standard encryption uses a fixed spin count of 50000 (see Apache POI StandardEncryptionVerifier).
    let pw_hash = hash_password_sha1(password, &verifier.salt, 50_000);

    // blockKey = 0 (little-endian u32)
    let mut h = Sha1::new();
    h.update(pw_hash);
    h.update(0u32.to_le_bytes());
    let digest = h.finalize();
    let mut final_hash = [0u8; 20];
    final_hash.copy_from_slice(&digest);

    // Derive key using HMAC-like construction described in MS-OFFCRYPTO (and used by Apache POI).
    let x1 = fill_and_xor_then_sha1(&final_hash, 0x36);
    let x2 = fill_and_xor_then_sha1(&final_hash, 0x5c);

    let mut x3 = Vec::with_capacity(40);
    x3.extend_from_slice(&x1);
    x3.extend_from_slice(&x2);
    x3.truncate(key_size_bytes);
    x3
}

fn aes_ecb_decrypt(key: &[u8], ciphertext: &[u8]) -> Result<Vec<u8>, String> {
    if ciphertext.len() % 16 != 0 {
        return Err("ciphertext length must be a multiple of 16 for AES-ECB".to_string());
    }

    match key.len() {
        16 => {
            let cipher = aes::Aes128::new_from_slice(key).map_err(|e| e.to_string())?;
            let mut out = ciphertext.to_vec();
            for chunk in out.chunks_mut(16) {
                let block = aes::cipher::generic_array::GenericArray::from_mut_slice(chunk);
                cipher.decrypt_block(block);
            }
            Ok(out)
        }
        24 => {
            let cipher = aes::Aes192::new_from_slice(key).map_err(|e| e.to_string())?;
            let mut out = ciphertext.to_vec();
            for chunk in out.chunks_mut(16) {
                let block = aes::cipher::generic_array::GenericArray::from_mut_slice(chunk);
                cipher.decrypt_block(block);
            }
            Ok(out)
        }
        32 => {
            let cipher = aes::Aes256::new_from_slice(key).map_err(|e| e.to_string())?;
            let mut out = ciphertext.to_vec();
            for chunk in out.chunks_mut(16) {
                let block = aes::cipher::generic_array::GenericArray::from_mut_slice(chunk);
                cipher.decrypt_block(block);
            }
            Ok(out)
        }
        other => Err(format!("unsupported AES key size {other} bytes")),
    }
}

fn verify_standard_password(info: &StandardEncryptionInfo, password: &str) -> Result<bool, String> {
    // CALG_AES_128 = 0x660E, CALG_AES_192 = 0x660F, CALG_AES_256 = 0x6610.
    // CALG_RC4     = 0x6801.
    const CALG_AES_128: u32 = 0x0000_660E;
    const CALG_AES_192: u32 = 0x0000_660F;
    const CALG_AES_256: u32 = 0x0000_6610;

    if !matches!(
        info.header.alg_id,
        CALG_AES_128 | CALG_AES_192 | CALG_AES_256
    ) {
        return Err(format!(
            "unexpected algId 0x{:08x} (fixture expected AES)",
            info.header.alg_id
        ));
    }

    // CALG_SHA1 = 0x8004.
    if info.header.alg_id_hash != 0x0000_8004 {
        return Err(format!(
            "unexpected algIdHash 0x{:08x} (expected SHA-1)",
            info.header.alg_id_hash
        ));
    }

    let key_size_bytes = (info.header.key_size_bits / 8) as usize;
    let key = derive_standard_key(password, &info.verifier, key_size_bytes);
    let verifier = aes_ecb_decrypt(&key, &info.verifier.encrypted_verifier)?;
    let decrypted_verifier_hash = aes_ecb_decrypt(&key, &info.verifier.encrypted_verifier_hash)?;

    let calc_hash = sha1_bytes(&verifier);
    let expected_len = info.verifier.verifier_hash_size as usize;
    let expected = decrypted_verifier_hash
        .get(..expected_len)
        .ok_or_else(|| "decrypted verifier hash truncated".to_string())?;
    if expected_len != calc_hash.len() {
        return Ok(false);
    }
    Ok(calc_hash.as_slice() == expected)
}

#[test]
fn parses_encryption_info_and_verifies_password() {
    let path = fixture_path();
    let file = std::fs::File::open(&path).expect("open encrypted fixture");
    let mut ole = cfb::CompoundFile::open(file).expect("parse OLE container");

    assert!(
        ole.exists("EncryptionInfo"),
        "fixture is expected to contain an EncryptionInfo stream"
    );
    assert!(
        ole.exists("EncryptedPackage"),
        "fixture is expected to contain an EncryptedPackage stream"
    );

    let mut stream = ole
        .open_stream("EncryptionInfo")
        .expect("open EncryptionInfo stream");
    let mut buf = Vec::new();
    stream.read_to_end(&mut buf).expect("read EncryptionInfo");

    let info = parse_standard_encryption_info(&buf).expect("parse Standard EncryptionInfo");

    assert_eq!(info.version_major, 3, "expected Standard encryption major version 3");
    assert_eq!(info.version_minor, 2, "expected Standard encryption minor version 2");

    // EncryptionInfo flags should include fCryptoAPI (0x04) and fAES (0x20).
    assert_eq!(
        info.flags, 0x24,
        "expected fCryptoAPI|fAES flags (0x24), got 0x{:08x}",
        info.flags
    );

    // CALG_AES_128 = 0x660E.
    assert_eq!(
        info.header.alg_id, 0x0000_660E,
        "expected algId=CALG_AES_128 (0x660E)"
    );
    assert_eq!(info.header.key_size_bits, 128);

    let ok = verify_standard_password(&info, FIXTURE_PASSWORD).expect("verify password");
    assert!(ok, "expected correct password to verify");

    let bad = verify_standard_password(&info, "not-the-password").expect("verify wrong password");
    assert!(!bad, "expected wrong password to fail verification");
}
