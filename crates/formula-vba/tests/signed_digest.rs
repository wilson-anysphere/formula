#![cfg(not(target_arch = "wasm32"))]

use formula_vba::extract_vba_signature_signed_digest;

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
        Pkcs7Flags::BINARY | Pkcs7Flags::DETACHED | Pkcs7Flags::NOATTR,
    )
    .expect("pkcs7 sign");
    pkcs7.to_der().expect("pkcs7 DER")
}

fn der_len(len: usize) -> Vec<u8> {
    if len < 0x80 {
        return vec![len as u8];
    }
    let mut bytes = Vec::new();
    let mut tmp = len;
    while tmp > 0 {
        bytes.push((tmp & 0xFF) as u8);
        tmp >>= 8;
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
    out.extend_from_slice(&der_len(value.len()));
    out.extend_from_slice(value);
    out
}

fn der_sequence(children: &[Vec<u8>]) -> Vec<u8> {
    let mut value = Vec::new();
    for child in children {
        value.extend_from_slice(child);
    }
    der_tlv(0x30, &value)
}

fn der_octet_string(bytes: &[u8]) -> Vec<u8> {
    der_tlv(0x04, bytes)
}

fn der_null() -> Vec<u8> {
    vec![0x05, 0x00]
}

fn der_oid(oid: &str) -> Vec<u8> {
    let arcs: Vec<u32> = oid
        .split('.')
        .map(|s| s.parse::<u32>().expect("numeric arc"))
        .collect();
    assert!(arcs.len() >= 2, "OID needs at least 2 arcs");
    let mut out = Vec::new();
    out.push((arcs[0] * 40 + arcs[1]) as u8);
    for &arc in &arcs[2..] {
        let mut tmp = arc;
        let mut buf = Vec::new();
        buf.push((tmp & 0x7F) as u8);
        tmp >>= 7;
        while tmp > 0 {
            buf.push(((tmp & 0x7F) as u8) | 0x80);
            tmp >>= 7;
        }
        buf.reverse();
        out.extend_from_slice(&buf);
    }
    der_tlv(0x06, &out)
}

fn make_spc_indirect_data_content_sha256(digest: &[u8]) -> Vec<u8> {
    // data SpcAttributeTypeAndOptionalValue ::= SEQUENCE { type OBJECT IDENTIFIER, value [0] EXPLICIT ANY OPTIONAL }
    let data = der_sequence(&[der_oid("1.3.6.1.4.1.311.2.1.15")]);

    // messageDigest DigestInfo ::= SEQUENCE { digestAlgorithm AlgorithmIdentifier, digest OCTET STRING }
    let alg = der_sequence(&[der_oid("2.16.840.1.101.3.4.2.1"), der_null()]);
    let digest_info = der_sequence(&[alg, der_octet_string(digest)]);

    der_sequence(&[data, digest_info])
}

#[test]
fn extracts_signed_digest_from_embedded_pkcs7() {
    let digest = (0u8..32).collect::<Vec<_>>();
    let spc = make_spc_indirect_data_content_sha256(&digest);
    let pkcs7 = make_pkcs7_signed_message(&spc);

    let got = extract_vba_signature_signed_digest(&pkcs7)
        .expect("extract digest")
        .expect("digest present");
    assert_eq!(got.digest_algorithm_oid, "2.16.840.1.101.3.4.2.1");
    assert_eq!(got.digest, digest);
}

#[test]
fn extracts_signed_digest_from_detached_pkcs7_using_prefix_content() {
    let digest = (42u8..74).collect::<Vec<_>>();
    let spc = make_spc_indirect_data_content_sha256(&digest);
    let pkcs7 = make_pkcs7_detached_signature(&spc);

    let mut stream = spc.clone();
    stream.extend_from_slice(&pkcs7);

    let got = extract_vba_signature_signed_digest(&stream)
        .expect("extract digest")
        .expect("digest present");
    assert_eq!(got.digest_algorithm_oid, "2.16.840.1.101.3.4.2.1");
    assert_eq!(got.digest, digest);
}

#[test]
fn extracts_signed_digest_when_pkcs7_is_prefixed_by_header_bytes() {
    let digest = b"this-is-a-test-digest-32-bytes!!".to_vec();
    assert_eq!(digest.len(), 32);
    let spc = make_spc_indirect_data_content_sha256(&digest);
    let pkcs7 = make_pkcs7_signed_message(&spc);

    let mut stream = b"VBA\0SIG\0HDR".to_vec();
    stream.extend_from_slice(&pkcs7);

    let got = extract_vba_signature_signed_digest(&stream)
        .expect("extract digest")
        .expect("digest present");
    assert_eq!(got.digest_algorithm_oid, "2.16.840.1.101.3.4.2.1");
    assert_eq!(got.digest, digest);
}
