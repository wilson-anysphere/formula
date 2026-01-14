use std::io::Read;
use std::path::PathBuf;

use formula_io::offcrypto::{
    parse_encryption_info_standard, verify_password_standard, CALG_AES_128, CALG_AES_192,
    CALG_AES_256, CALG_RC4, CALG_SHA1,
};

const FIXTURE_PASSWORD: &str = "password";

/// Fixture produced with Apache POI (standard encryption / CryptoAPI).
///
/// Standard encryption is identified by `versionMinor == 2`, but `versionMajor` varies in the wild
/// (commonly 3.2 or 4.2). This fixture uses a Standard header and should remain valid across those
/// variants.
///
/// The workbook contains a single sheet ("Sheet1") with cell A1="hello".
fn fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures/offcrypto_standard_cryptoapi_password.xlsx")
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

    assert!(buf.len() >= 4, "EncryptionInfo stream is too short");
    let major = u16::from_le_bytes([buf[0], buf[1]]);
    let minor = u16::from_le_bytes([buf[2], buf[3]]);
    assert!(
        minor == 2 && matches!(major, 2 | 3 | 4),
        "expected Standard EncryptionInfo version *.2 with major=2/3/4, got {major}.{minor}"
    );

    let info = parse_encryption_info_standard(&buf).expect("parse Standard EncryptionInfo");

    assert!(
        info.header.flags.f_cryptoapi,
        "expected EncryptionHeader.flags.fCryptoAPI to be set"
    );
    assert!(
        matches!(info.header.alg_id, CALG_AES_128 | CALG_AES_192 | CALG_AES_256 | CALG_RC4),
        "expected Standard encryption algId to be AES or RC4, got 0x{:08x}",
        info.header.alg_id
    );
    assert_eq!(
        info.header.alg_id, CALG_AES_128,
        "expected fixture algId=CALG_AES_128 (0x660E)"
    );
    assert_eq!(info.header.alg_id_hash, CALG_SHA1);
    assert_eq!(info.header.key_size, 128);

    let ok = verify_password_standard(&info, FIXTURE_PASSWORD).expect("verify password");
    assert!(ok, "expected correct password to verify");

    let bad =
        verify_password_standard(&info, "not-the-password").expect("verify wrong password");
    assert!(!bad, "expected wrong password to fail verification");
}
