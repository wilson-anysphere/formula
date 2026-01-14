//! End-to-end fixture test for MS-OFFCRYPTO Standard / CryptoAPI / RC4 encrypted OOXML.
//!
//! This exercises the full parsing + verifier validation + `EncryptedPackage` decryption path
//! (including the **0x200-byte** per-block RC4 re-key interval).

use std::io::Read as _;
use std::path::PathBuf;

use formula_io::offcrypto::cryptoapi::{hash_password_fixed_spin, password_to_utf16le};
use formula_io::offcrypto::{parse_encryption_info_standard, verify_password_standard, CALG_RC4};
use formula_io::{HashAlg, Rc4CryptoApiDecryptReader};
use sha2::{Digest as _, Sha256};

fn fixture_path(rel: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/encrypted/ooxml")
        .join(rel)
}

fn open_stream_case_tolerant<R: std::io::Read + std::io::Write + std::io::Seek>(
    ole: &mut cfb::CompoundFile<R>,
    name: &str,
) -> std::io::Result<cfb::Stream<R>> {
    ole.open_stream(name)
        .or_else(|_| ole.open_stream(format!("/{name}")))
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

    // Read EncryptionInfo and parse the Standard/CryptoAPI header + verifier.
    let mut encryption_info_bytes = Vec::new();
    open_stream_case_tolerant(&mut ole, "EncryptionInfo")
        .expect("open EncryptionInfo")
        .read_to_end(&mut encryption_info_bytes)
        .expect("read EncryptionInfo");
    let info = parse_encryption_info_standard(&encryption_info_bytes).expect("parse EncryptionInfo");
    assert_eq!(info.header.alg_id, CALG_RC4);

    // Wrong password must fail verifier check.
    assert!(
        !verify_password_standard(&info, "wrong-password").expect("verify_password_standard"),
        "expected wrong password to fail verifier check"
    );

    // Correct password must pass verifier check.
    assert!(
        verify_password_standard(&info, "password").expect("verify_password_standard"),
        "expected correct password to pass verifier check"
    );

    // Derive the base hash `H` (fixed 50k spin) used for per-block RC4 keys:
    //   rc4_key_block = Hash(H || LE32(block_index))[..key_len]
    let hash_alg = HashAlg::from_calg_id(info.header.alg_id_hash).expect("HashAlg::from_calg_id");
    let password_utf16le = password_to_utf16le("password");
    let h = hash_password_fixed_spin(&password_utf16le, &info.verifier.salt, hash_alg);

    // Decrypt using the library's 0x200-block RC4 decrypt reader.
    let encrypted_package_stream =
        open_stream_case_tolerant(&mut ole, "EncryptedPackage").expect("open EncryptedPackage");
    let mut reader = Rc4CryptoApiDecryptReader::from_encrypted_package_stream(
        encrypted_package_stream,
        h,
        info.header.key_size,
        info.header.alg_id_hash,
    )
    .expect("create RC4 decrypt reader");
    let package_size = reader.package_size();
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
