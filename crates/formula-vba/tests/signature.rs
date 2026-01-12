#![cfg(not(target_arch = "wasm32"))]

use std::io::{Cursor, Write};

use formula_vba::{
    extract_signer_certificate_info, verify_vba_digital_signature, VbaSignatureVerification,
};

const TEST_KEY_PEM: &str = r#"-----BEGIN PRIVATE KEY-----
MIIEvAIBADANBgkqhkiG9w0BAQEFAASCBKYwggSiAgEAAoIBAQC8kN1a0raWt6a7
MzszVTIVgdZHbie+mkVWDoMrgTQYX8tm/3yqTLQMXWhuV0hZtrUydWlsRB8k0aTS
aXFCzmmNgAqFh13uQ/rFW82zh5UCWXuaX43uc5JWebD4TzkN2b4vye3s/S3QCmZK
5kT6jWPDaRyngOvaHgcBB9meMS6QT9Efb2SdV/a6QkrGm0nhMfJyZEY00FKEhxJf
A4JlVDVhmmQdpCoXb++cqK/xo9DehmrivP1CL/dFPjy3wkbtHpb+uAatzBNtaqmE
bYwtaw0rqxlkbKZT6baayf9klTXFah4bEzRDSJQrzM6HjhNYDiCBM9omNSowkyb9
PVJqkRRvAgMBAAECggEAAalfIflAXaShpf2mFGY4SkM6IagBVqciXEdFdaEuVdam
QrKWpSOG5KMAFBTV0OCQyTCKrMcO5TKpuqbuNhH+kR4jOZj/RWW49HtCHUZhFEO4
mJwl8od3LybkXdPI42vbRq2HWLLEcBRfYWKVEgYj7mljNMCok2P3WNV6X+/8Ao6F
n9+NDnE/K4e3xr/7pS4hldm/b67KZh92Rhzfoezdpk+uVXtU6ccTeyO10YCnng2w
Qhls2Hkrx92GspAp8gdK5Hnrk2y/Lx8EmThUUSSP7h2uxvoUs3RNevZQYp2vv6fc
0ffD1M9fI9iz79UKkad+1VGwjO2SPPK28LZWUNgO1QKBgQDjX/4W7ZlbeYNN1sxa
pbdEB0eNFs5jk3B4JH986h0lTPZwdkcaEDwJ9sp/pxceYPFZ8ul/IGg9nlDNIa88
BccrH/o1gZjB68UdM2fu6jWYDC9dzscUYjFkFDndQoH8ACg1Nt27UZi5TKN1DTGM
dnPf4Tb7VDOsVpPiCxZDCsgFbQKBgQDUThqr43mgAaB8nlmCVRFIP4Wn3BQScXt/
J5xLKsI5AadvCbNUPw/gO518qhcsFNSKRUzBx1d10AR1h/NX820td/swjHkWm31V
PjbNl+5G8RwHib2miqdI4KRe5RdGeWfSW9wEU1epkeCGPBbApNfgEec0/PN0T+7q
xNBiaSaDywKBgGPMNT0hCkexHOWkWsuKota0Dz6o/OuNwjapZl+Qbjx5/Ey+TVTu
PTvuW1EOKMKHsEdXrA7FTZuGClcO6tgAfTu7bFnhyQeMkVbQwlSF7gIPjxawdIbI
1n7jtcYcs+rEsuEwdMAL/2mNbs0ofk/1icSBGF3VxlxlH8F+NkY0zDg9AoGAOQi6
dY7or6mAObo4haDgwa3+8/dVlRbTfHdhr3fPMY1WM6hBetJuK2kYh9MR4o++AV9Y
nX416rp1WDWrk+cbX2mqG4LBTOd8phfOlTDJnFlNlGDWiBUbl6JxxeR5ej9HOuXe
l3LkS/Oag7VEz3/5VoK4wC1sIcUPhBZXfPiOlj0CgYAVBqxAtjYV+Of4nzYXlvyD
nKgzkiBZCPvjLuINLxl02hMkl5L1rkYYFlBonRXkBZi/qi/sy5yWJFD4bNdXADjx
l6I38mljR1b525IXYYgxl70AE5/oiURtl3rzv4gzYvm7lhV7/c7ZTwY0X43vTO7d
0TiTGpZ2jyGWBsNrW2X+Rw==
-----END PRIVATE KEY-----"#;

const TEST_CERT_PEM: &str = r#"-----BEGIN CERTIFICATE-----
MIIDFzCCAf+gAwIBAgIUQZEa3yk9CWWcytfnuDxC4+5iaPUwDQYJKoZIhvcNAQEL
BQAwGzEZMBcGA1UEAwwQRm9ybXVsYSBWQkEgVGVzdDAeFw0yNjAxMTExMDM2NDBa
Fw0zNjAxMDkxMDM2NDBaMBsxGTAXBgNVBAMMEEZvcm11bGEgVkJBIFRlc3QwggEi
MA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQC8kN1a0raWt6a7MzszVTIVgdZH
bie+mkVWDoMrgTQYX8tm/3yqTLQMXWhuV0hZtrUydWlsRB8k0aTSaXFCzmmNgAqF
h13uQ/rFW82zh5UCWXuaX43uc5JWebD4TzkN2b4vye3s/S3QCmZK5kT6jWPDaRyn
gOvaHgcBB9meMS6QT9Efb2SdV/a6QkrGm0nhMfJyZEY00FKEhxJfA4JlVDVhmmQd
pCoXb++cqK/xo9DehmrivP1CL/dFPjy3wkbtHpb+uAatzBNtaqmEbYwtaw0rqxlk
bKZT6baayf9klTXFah4bEzRDSJQrzM6HjhNYDiCBM9omNSowkyb9PVJqkRRvAgMB
AAGjUzBRMB0GA1UdDgQWBBSyceRXYQd4wvXncCr1AcYneVlpWTAfBgNVHSMEGDAW
gBSyceRXYQd4wvXncCr1AcYneVlpWTAPBgNVHRMBAf8EBTADAQH/MA0GCSqGSIb3
DQEBCwUAA4IBAQBbcQVLwUMdKA5xj2woUkEe9kcTtS9YOMeCoBE48Fw8KfgkbKtK
lte7yIBdgHdjjAke88g9Dh64OlcRQigu0fS025bXcw1g7AKc0fkBDro8j8GHqdi6
APR5O9xnfdslBSX1cDN/530Q+vRpha/LxLfSG2UXovmb163110RD6ina9gTIvy9r
plrbDIYpuR+SiI0uaQtcwCdbXPtHLlEUUp0ZbnW3i+RHmt9DnwQM1B/hAv9zdg9m
ls5Xirz7pTI39gHpSd86SfJWBbPPcJHabdmgRTJW8AbxMjS2xBDU3pxzGw52MgfK
Kj4ozoiZRiNvvWvqUGOt1yKu7S7nbEPuW3rX
-----END CERTIFICATE-----"#;

fn make_pkcs7_signed_message(data: &[u8]) -> Vec<u8> {
    use openssl::pkcs7::{Pkcs7, Pkcs7Flags};
    use openssl::pkey::PKey;
    use openssl::stack::Stack;
    use openssl::x509::X509;

    let pkey = PKey::private_key_from_pem(TEST_KEY_PEM.as_bytes()).expect("parse private key");
    let cert = X509::from_pem(TEST_CERT_PEM.as_bytes()).expect("parse certificate");
    let extra_certs = Stack::new().expect("create cert stack");

    let pkcs7 = Pkcs7::sign(
        &cert,
        &pkey,
        &extra_certs,
        data,
        Pkcs7Flags::BINARY | Pkcs7Flags::NOATTR,
    )
    .expect("pkcs7 sign");
    pkcs7.to_der().expect("pkcs7 DER")
}

fn make_pkcs7_detached_signature(data: &[u8]) -> Vec<u8> {
    use openssl::pkcs7::{Pkcs7, Pkcs7Flags};
    use openssl::pkey::PKey;
    use openssl::stack::Stack;
    use openssl::x509::X509;

    let pkey = PKey::private_key_from_pem(TEST_KEY_PEM.as_bytes()).expect("parse private key");
    let cert = X509::from_pem(TEST_CERT_PEM.as_bytes()).expect("parse certificate");
    let extra_certs = Stack::new().expect("create cert stack");

    let pkcs7 = Pkcs7::sign(
        &cert,
        &pkey,
        &extra_certs,
        data,
        // NOATTR keeps the output deterministic (avoids adding a SigningTime attribute).
        Pkcs7Flags::BINARY | Pkcs7Flags::DETACHED | Pkcs7Flags::NOATTR,
    )
    .expect("pkcs7 sign");
    pkcs7.to_der().expect("pkcs7 DER")
}

#[test]
fn extracts_signer_certificate_metadata_from_pkcs7() {
    let pkcs7 = make_pkcs7_signed_message(b"formula-vba-test");
    let info = extract_signer_certificate_info(&pkcs7).expect("expected embedded certificate info");

    assert!(
        info.subject.contains("Formula VBA Test"),
        "expected subject to mention test CN, got: {}",
        info.subject
    );
    assert!(
        !info.issuer.is_empty(),
        "expected issuer to be present, got empty issuer"
    );
    assert_eq!(
        info.issuer, info.subject,
        "test certificate is self-signed, expected issuer == subject"
    );
    assert_eq!(
        info.sha256_fingerprint_hex.len(),
        64,
        "expected SHA-256 fingerprint to be 64 hex chars, got {}",
        info.sha256_fingerprint_hex
    );
    assert!(
        info.sha256_fingerprint_hex
            .chars()
            .all(|c| c.is_ascii_digit() || ('a'..='f').contains(&c)),
        "expected lowercase hex fingerprint, got {}",
        info.sha256_fingerprint_hex
    );
    assert!(
        !info.serial_hex.is_empty(),
        "expected serial_hex to be non-empty"
    );
}

fn wrap_in_digsig_info_serialized(pkcs7: &[u8]) -> Vec<u8> {
    // Synthetic DigSigInfoSerialized-like blob:
    // [cbSignature, cbSigningCertStore, cchProjectName] (LE u32)
    // [projectName UTF-16LE] [certStore bytes] [signature bytes]
    let project_name_utf16: Vec<u16> = "VBAProject\0".encode_utf16().collect();
    let mut project_name_bytes = Vec::new();
    for ch in &project_name_utf16 {
        project_name_bytes.extend_from_slice(&ch.to_le_bytes());
    }

    let cert_store = vec![0xAA, 0xBB, 0xCC, 0xDD];

    let cb_signature = pkcs7.len() as u32;
    let cb_cert_store = cert_store.len() as u32;
    let cch_project = project_name_utf16.len() as u32;

    let mut out = Vec::new();
    out.extend_from_slice(&cb_signature.to_le_bytes());
    out.extend_from_slice(&cb_cert_store.to_le_bytes());
    out.extend_from_slice(&cch_project.to_le_bytes());
    out.extend_from_slice(&project_name_bytes);
    out.extend_from_slice(&cert_store);
    out.extend_from_slice(pkcs7);
    out
}

fn build_vba_project_bin_with_signature_streams(streams: &[(&str, &[u8])]) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create compound file");

    for (path, bytes) in streams {
        let mut stream = ole.create_stream(path).expect("create signature stream");
        stream.write_all(bytes).expect("write signature bytes");
    }

    ole.into_inner().into_inner()
}

fn build_vba_project_bin_with_signature(signature_blob: Option<&[u8]>) -> Vec<u8> {
    match signature_blob {
        Some(sig) => build_vba_project_bin_with_signature_streams(&[("\u{0005}DigitalSignature", sig)]),
        None => build_vba_project_bin_with_signature_streams(&[]),
    }
}

#[test]
fn unsigned_project_reports_no_signature() {
    let vba = build_vba_project_bin_with_signature(None);
    let sig = verify_vba_digital_signature(&vba).expect("signature inspection should succeed");
    assert!(sig.is_none(), "expected no signature");
}

#[test]
fn valid_pkcs7_signature_is_reported_as_verified() {
    let pkcs7 = make_pkcs7_signed_message(b"formula-vba-test");
    let vba = build_vba_project_bin_with_signature(Some(&pkcs7));

    let sig = verify_vba_digital_signature(&vba)
        .expect("signature inspection should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert!(
        sig.signer_subject
            .as_deref()
            .is_some_and(|s| s.contains("Formula VBA Test")),
        "expected signer subject to mention test CN, got: {:?}",
        sig.signer_subject
    );
}

#[test]
fn corrupting_signature_bytes_marks_signature_invalid() {
    let mut pkcs7 = make_pkcs7_signed_message(b"formula-vba-test");
    let last = pkcs7.len().saturating_sub(1);
    pkcs7[last] ^= 0xFF;
    let vba = build_vba_project_bin_with_signature(Some(&pkcs7));

    let sig = verify_vba_digital_signature(&vba)
        .expect("signature inspection should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedInvalid);
}

#[test]
fn pkcs7_signature_with_prefix_is_still_verified() {
    let pkcs7 = make_pkcs7_signed_message(b"formula-vba-test");
    let mut prefixed = b"VBA\0SIG\0".to_vec();
    prefixed.extend_from_slice(&pkcs7);
    let vba = build_vba_project_bin_with_signature(Some(&prefixed));

    let sig = verify_vba_digital_signature(&vba)
        .expect("signature inspection should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert!(
        sig.signer_subject
            .as_deref()
            .is_some_and(|s| s.contains("Formula VBA Test")),
        "expected signer subject to mention test CN, got: {:?}",
        sig.signer_subject
    );
}

#[test]
fn pkcs7_signature_wrapped_in_digsig_info_serialized_is_still_verified() {
    let pkcs7 = make_pkcs7_signed_message(b"formula-vba-test");
    let wrapped = wrap_in_digsig_info_serialized(&pkcs7);
    let vba = build_vba_project_bin_with_signature(Some(&wrapped));

    let sig = verify_vba_digital_signature(&vba)
        .expect("signature inspection should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert!(
        sig.signer_subject
            .as_deref()
            .is_some_and(|s| s.contains("Formula VBA Test")),
        "expected signer subject to mention test CN, got: {:?}",
        sig.signer_subject
    );
}

#[test]
fn detached_pkcs7_signature_with_prefixed_content_is_verified() {
    let content = b"formula-vba-detached-test";
    let pkcs7 = make_pkcs7_detached_signature(content);
    let mut blob = content.to_vec();
    blob.extend_from_slice(&pkcs7);
    let vba = build_vba_project_bin_with_signature(Some(&blob));

    let sig = verify_vba_digital_signature(&vba)
        .expect("signature inspection should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert!(
        sig.signer_subject
            .as_deref()
            .is_some_and(|s| s.contains("Formula VBA Test")),
        "expected signer subject to mention test CN, got: {:?}",
        sig.signer_subject
    );
}

#[test]
fn prefers_verified_signature_stream_over_invalid_candidate() {
    let mut invalid = make_pkcs7_signed_message(b"formula-vba-test");
    let last = invalid.len().saturating_sub(1);
    invalid[last] ^= 0xFF;

    let valid = make_pkcs7_signed_message(b"formula-vba-test");

    let vba = build_vba_project_bin_with_signature_streams(&[
        ("\u{0005}DigitalSignature", &invalid),
        ("\u{0005}DigitalSignatureEx", &valid),
    ]);

    let sig = verify_vba_digital_signature(&vba)
        .expect("signature inspection should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert!(
        sig.stream_path.contains("DigitalSignatureEx"),
        "expected to pick verified signature stream, got {}",
        sig.stream_path
    );
    assert!(
        sig.signer_subject
            .as_deref()
            .is_some_and(|s| s.contains("Formula VBA Test")),
        "expected signer subject to mention test CN, got: {:?}",
        sig.signer_subject
    );
}

#[test]
fn unparseable_signature_stream_is_reported_as_parse_error() {
    let blob = b"not-a-pkcs7".to_vec();
    let vba = build_vba_project_bin_with_signature_streams(&[("\u{0005}DigitalSignature", &blob)]);

    let sig = verify_vba_digital_signature(&vba)
        .expect("signature inspection should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedParseError);
}

#[test]
fn prefers_verified_stream_over_parse_error_candidate() {
    let bad = b"not-a-pkcs7".to_vec();
    let valid = make_pkcs7_signed_message(b"formula-vba-test");

    let vba = build_vba_project_bin_with_signature_streams(&[
        ("\u{0005}DigitalSignature", &bad),
        ("\u{0005}DigitalSignatureEx", &valid),
    ]);

    let sig = verify_vba_digital_signature(&vba)
        .expect("signature inspection should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert!(
        sig.stream_path.contains("DigitalSignatureEx"),
        "expected to pick verified signature stream, got {}",
        sig.stream_path
    );
}
