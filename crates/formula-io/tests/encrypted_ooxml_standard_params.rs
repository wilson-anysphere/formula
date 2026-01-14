use std::io::Read as _;
use std::path::{Path, PathBuf};

use formula_offcrypto::{parse_encryption_info, EncryptionInfo, EncryptionVersionInfo};

/// Encrypted OOXML fixtures live at `fixtures/encrypted/ooxml/`.
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

fn read_stream(path: &Path, stream_name: &str) -> Vec<u8> {
    let file = std::fs::File::open(path)
        .unwrap_or_else(|err| panic!("open fixture file {} failed: {err}", path.display()));
    let mut ole = cfb::CompoundFile::open(file)
        .unwrap_or_else(|err| panic!("open cfb (OLE) container {} failed: {err}", path.display()));

    let mut stream = open_stream_case_tolerant(&mut ole, stream_name).unwrap_or_else(|err| {
        panic!(
            "open {stream_name} stream from {} failed: {err}",
            path.display()
        )
    });

    let mut buf = Vec::new();
    stream.read_to_end(&mut buf).unwrap_or_else(|err| {
        panic!(
            "read {stream_name} stream from {} failed: {err}",
            path.display()
        )
    });
    buf
}

#[derive(Debug, Clone, Copy)]
struct ExpectedStandardParams {
    alg_id: u32,
    alg_id_hash: u32,
    key_size_bits: u32,
    salt_size: usize,
    verifier_hash_size: u32,
    encrypted_verifier_hash_len: usize,
}

fn assert_standard_fixture_encryption_params(path: &Path, expected: ExpectedStandardParams) {
    let encryption_info = read_stream(path, "EncryptionInfo");

    // --- Version header pinning ---------------------------------------------------------------
    let version =
        EncryptionVersionInfo::parse(&encryption_info).expect("parse EncryptionVersionInfo");
    assert_eq!(
        version.minor,
        2,
        "expected Standard EncryptionInfo versionMinor==2, got {} for {}",
        version.minor,
        path.display()
    );
    assert!(
        matches!(version.major, 2 | 3 | 4),
        "expected Standard EncryptionInfo versionMajor in {{2,3,4}}, got {} for {}",
        version.major,
        path.display()
    );

    // --- Standard parameter pinning ----------------------------------------------------------
    let parsed = parse_encryption_info(&encryption_info)
        .unwrap_or_else(|err| panic!("parse EncryptionInfo from {} failed: {err}", path.display()));
    let (header, verifier) = match parsed {
        EncryptionInfo::Standard {
            header, verifier, ..
        } => (header, verifier),
        other => panic!(
            "expected Standard EncryptionInfo for {}, got {other:?}",
            path.display()
        ),
    };

    assert_eq!(
        header.alg_id,
        expected.alg_id,
        "cipher algId changed for {}",
        path.display()
    );
    assert_eq!(
        header.alg_id_hash,
        expected.alg_id_hash,
        "hash algIdHash changed for {}",
        path.display()
    );
    assert_eq!(
        header.key_size_bits,
        expected.key_size_bits,
        "keySize changed for {}",
        path.display()
    );
    assert_eq!(
        verifier.salt.len(),
        expected.salt_size,
        "saltSize changed for {}",
        path.display()
    );
    assert_eq!(
        verifier.verifier_hash_size,
        expected.verifier_hash_size,
        "verifierHashSize changed for {}",
        path.display()
    );
    assert_eq!(
        verifier.encrypted_verifier_hash.len(),
        expected.encrypted_verifier_hash_len,
        "encryptedVerifierHash length changed for {}",
        path.display()
    );
}

#[test]
fn standard_encryption_info_parameters_are_pinned() {
    // CryptoAPI ALG_ID constants used by the MS-OFFCRYPTO Standard encryption header.
    const CALG_AES_128: u32 = 0x0000_660E;
    const CALG_RC4: u32 = 0x0000_6801;
    const CALG_SHA1: u32 = 0x0000_8004;

    let expected_aes = ExpectedStandardParams {
        alg_id: CALG_AES_128,
        alg_id_hash: CALG_SHA1,
        key_size_bits: 128,
        salt_size: 16,
        verifier_hash_size: 20,
        // SHA1 (20 bytes) padded to AES block alignment (16) => 32 bytes.
        encrypted_verifier_hash_len: 32,
    };

    let expected_rc4 = ExpectedStandardParams {
        alg_id: CALG_RC4,
        alg_id_hash: CALG_SHA1,
        key_size_bits: 128,
        salt_size: 16,
        verifier_hash_size: 20,
        // RC4 encrypts the verifier hash without AES block padding.
        encrypted_verifier_hash_len: 20,
    };

    for (rel, expected) in [
        ("standard.xlsx", expected_aes),
        ("standard-large.xlsx", expected_aes),
        ("standard-basic.xlsm", expected_aes),
        ("standard-4.2.xlsx", expected_aes),
        ("standard-unicode.xlsx", expected_aes),
        ("standard-rc4.xlsx", expected_rc4),
    ] {
        let path = fixture_path(rel);
        assert_standard_fixture_encryption_params(&path, expected);
    }
}
