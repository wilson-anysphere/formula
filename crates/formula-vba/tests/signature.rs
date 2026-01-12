#![cfg(not(target_arch = "wasm32"))]

use std::io::{Cursor, Write};

use formula_vba::{
    extract_signer_certificate_info, list_vba_digital_signatures, verify_vba_digital_signature,
    VbaSignatureVerification,
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

fn der_encode_len(len: usize) -> Vec<u8> {
    if len < 128 {
        return vec![len as u8];
    }
    let mut bytes = Vec::new();
    let mut n = len;
    while n > 0 {
        bytes.push((n & 0xFF) as u8);
        n >>= 8;
    }
    bytes.reverse();
    let mut out = Vec::with_capacity(1 + bytes.len());
    out.push(0x80 | (bytes.len() as u8));
    out.extend_from_slice(&bytes);
    out
}

fn der_tlv(tag: u8, value: &[u8]) -> Vec<u8> {
    let mut out = Vec::new();
    out.push(tag);
    out.extend_from_slice(&der_encode_len(value.len()));
    out.extend_from_slice(value);
    out
}

fn der_oid(oid: &str) -> Vec<u8> {
    let parts = oid
        .split('.')
        .filter(|p| !p.is_empty())
        .map(|p| p.parse::<u32>().expect("valid OID component"))
        .collect::<Vec<_>>();
    assert!(parts.len() >= 2, "OID must have at least two components");
    let first = parts[0];
    let second = parts[1];
    assert!(first <= 2, "invalid OID first component");
    if first < 2 {
        assert!(second < 40, "invalid OID second component");
    }

    fn encode_subid(mut n: u32) -> Vec<u8> {
        let mut chunks = Vec::new();
        chunks.push((n & 0x7F) as u8);
        n >>= 7;
        while n > 0 {
            chunks.push(((n & 0x7F) as u8) | 0x80);
            n >>= 7;
        }
        chunks.reverse();
        chunks
    }

    let mut encoded = Vec::new();
    encoded.extend_from_slice(&encode_subid(first * 40 + second));

    for &n in &parts[2..] {
        encoded.extend_from_slice(&encode_subid(n));
    }

    der_tlv(0x06, &encoded)
}

fn der_null() -> Vec<u8> {
    vec![0x05, 0x00]
}

fn der_sequence(contents: &[u8]) -> Vec<u8> {
    der_tlv(0x30, contents)
}

fn der_octet_string(bytes: &[u8]) -> Vec<u8> {
    der_tlv(0x04, bytes)
}

fn make_spc_indirect_data_content(digest_algorithm_oid: &str, digest: &[u8]) -> Vec<u8> {
    // Minimal SpcIndirectDataContent:
    //   SEQUENCE {
    //     data SEQUENCE { type OID },
    //     messageDigest SEQUENCE {
    //       digestAlgorithm SEQUENCE { algorithm OID, parameters NULL },
    //       digest OCTET STRING
    //     }
    //   }
    let spc_data = {
        let mut out = Vec::new();
        out.extend_from_slice(&der_oid("1.2.3.4"));
        der_sequence(&out)
    };

    let digest_info = {
        let alg_id = {
            let mut out = Vec::new();
            out.extend_from_slice(&der_oid(digest_algorithm_oid));
            out.extend_from_slice(&der_null());
            der_sequence(&out)
        };

        let mut out = Vec::new();
        out.extend_from_slice(&alg_id);
        out.extend_from_slice(&der_octet_string(digest));
        der_sequence(&out)
    };

    let mut out = Vec::new();
    out.extend_from_slice(&spc_data);
    out.extend_from_slice(&digest_info);
    der_sequence(&out)
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

#[test]
fn prefers_digital_signature_ex_over_legacy_when_both_verify() {
    let legacy = make_pkcs7_signed_message(b"legacy-signed");
    let ex = make_pkcs7_signed_message(b"ex-signed");

    let vba = build_vba_project_bin_with_signature_streams(&[
        ("\u{0005}DigitalSignature", &legacy),
        ("\u{0005}DigitalSignatureEx", &ex),
    ]);

    let sig = verify_vba_digital_signature(&vba)
        .expect("signature inspection should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert!(
        sig.stream_path.contains("DigitalSignatureEx"),
        "expected DigitalSignatureEx to be treated as authoritative when both legacy and Ex signatures verify, got {}",
        sig.stream_path
    );
    assert_eq!(sig.signature, ex);
}

#[test]
fn prefers_digital_signature_ext_over_ex_when_both_verify() {
    let ex = make_pkcs7_signed_message(b"ex-signed");
    let ext = make_pkcs7_signed_message(b"ext-signed");

    let vba = build_vba_project_bin_with_signature_streams(&[
        ("\u{0005}DigitalSignatureEx", &ex),
        ("\u{0005}DigitalSignatureExt", &ext),
    ]);

    let sig = verify_vba_digital_signature(&vba)
        .expect("signature inspection should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert!(
        sig.stream_path.contains("DigitalSignatureExt"),
        "expected DigitalSignatureExt to be treated as authoritative when both Ex and Ext signatures verify, got {}",
        sig.stream_path
    );
    assert_eq!(sig.signature, ext);
}

#[test]
fn lists_all_signature_streams_and_extracts_digest_info_per_stream() {
    // Use a deterministic digest so we can assert exact output.
    let digest_algorithm_oid = "2.16.840.1.101.3.4.2.1"; // sha256
    let digest = (0u8..32u8).collect::<Vec<u8>>();
    let spc = make_spc_indirect_data_content(digest_algorithm_oid, &digest);

    let valid = make_pkcs7_signed_message(&spc);
    let mut invalid = valid.clone();
    let last = invalid.len().saturating_sub(1);
    invalid[last] ^= 0xFF;

    let vba = build_vba_project_bin_with_signature_streams(&[
        ("\u{0005}DigitalSignature", &invalid),
        ("\u{0005}DigitalSignatureEx", &valid),
    ]);

    let sigs = list_vba_digital_signatures(&vba).expect("signature enumeration should succeed");
    assert_eq!(sigs.len(), 2, "expected two signature streams");

    // Deterministic Excel-like ordering: DigitalSignatureEx is preferred over the legacy
    // DigitalSignature stream.
    assert!(
        sigs[0].stream_path.ends_with("\u{0005}DigitalSignatureEx"),
        "unexpected first stream path: {}",
        sigs[0].stream_path
    );
    assert!(
        sigs[1].stream_path.ends_with("\u{0005}DigitalSignature"),
        "unexpected second stream path: {}",
        sigs[1].stream_path
    );

    assert_eq!(sigs[0].verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sigs[1].verification, VbaSignatureVerification::SignedInvalid);

    assert!(
        sigs[0]
            .signer_subject
            .as_deref()
            .is_some_and(|s| s.contains("Formula VBA Test")),
        "expected signer subject to mention test CN, got: {:?}",
        sigs[0].signer_subject
    );

    for sig in &sigs {
        assert_eq!(
            sig.signed_digest_algorithm_oid.as_deref(),
            Some(digest_algorithm_oid),
            "expected per-stream digest algorithm OID"
        );
        assert_eq!(
            sig.signed_digest.as_deref(),
            Some(digest.as_slice()),
            "expected per-stream digest bytes"
        );
    }
}
