#![cfg(not(target_arch = "wasm32"))]

use std::io::{Cursor, Write};

use formula_vba::{
    compress_container, content_normalized_data, verify_vba_digital_signature,
    verify_vba_project_signature_binding, VbaProjectBindingVerification, VbaSignatureBinding,
    VbaSignatureVerification,
};
use md5::{Digest as _, Md5};

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

#[test]
fn ms_oshared_md5_digest_bytes_even_when_signeddata_uses_sha256() {
    use openssl::pkcs7::{Pkcs7, Pkcs7Flags};
    use openssl::pkey::PKey;
    use openssl::stack::Stack;
    use openssl::x509::X509;

    // ---- 1) Build a minimal spec-ish VBA project and compute its MS-OVBA content hash (MD5). ----
    let module_source = b"Sub Hello()\r\nEnd Sub\r\n";
    let module_container = compress_container(module_source);

    // Minimal `dir` stream (decompressed form) with a single module.
    let dir_decompressed = {
        let mut out = Vec::new();
        // PROJECTNAME
        push_record(&mut out, 0x0004, b"VBAProject");
        // MODULENAME
        push_record(&mut out, 0x0019, b"Module1");
        // MODULESTREAMNAME + reserved u16
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"Module1");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);
        // MODULETYPE (standard)
        push_record(&mut out, 0x0021, &0u16.to_le_bytes());
        // MODULETEXTOFFSET
        push_record(&mut out, 0x0031, &0u32.to_le_bytes());
        out
    };
    let dir_container = compress_container(&dir_decompressed);

    let project_stream_bytes: &[u8] = b"Name=\"VBAProject\"\r\nModule=Module1\r\n";
    let vba_project_stream_bytes: &[u8] = b"dummy";

    let unsigned_vba_project_bin = build_vba_project_bin_with_streams(&[
        ("PROJECT", project_stream_bytes),
        ("VBA/_VBA_PROJECT", vba_project_stream_bytes),
        ("VBA/dir", &dir_container),
        ("VBA/Module1", &module_container),
    ]);

    let normalized = content_normalized_data(&unsigned_vba_project_bin).expect("ContentNormalizedData");
    let project_md5: [u8; 16] = Md5::digest(&normalized).into();
    assert_eq!(project_md5.len(), 16, "VBA project digest must be 16-byte MD5");

    // ---- 2) Construct SpcIndirectDataContent with sha256 OID but MD5 digest bytes. ----
    // DigestInfo.digestAlgorithm.algorithm = sha256 (2.16.840.1.101.3.4.2.1)
    // DigestInfo.digestAlgorithm.parameters = NULL (per MS-OSHARED)
    // DigestInfo.digest = MD5(project)
    let spc_indirect_data_content =
        build_spc_indirect_data_content_sha256_oid_with_md5_digest(&project_md5);

    // ---- 3) Produce PKCS#7 SignedData using OpenSSL (signing with SHA-256 by default). ----
    let pkey = PKey::private_key_from_pem(TEST_KEY_PEM.as_bytes()).expect("parse private key");
    let cert = X509::from_pem(TEST_CERT_PEM.as_bytes()).expect("parse certificate");
    let extra_certs = Stack::new().expect("create cert stack");

    let pkcs7 = Pkcs7::sign(
        &cert,
        &pkey,
        &extra_certs,
        &spc_indirect_data_content,
        // NOATTR keeps output deterministic.
        Pkcs7Flags::BINARY | Pkcs7Flags::NOATTR,
    )
    .expect("pkcs7 sign");
    let pkcs7_der = pkcs7.to_der().expect("pkcs7 DER");

    // ---- 4) Store signature in a \x05DigitalSignature stream. ----
    let signed_streams = vec![
        ("PROJECT", project_stream_bytes),
        ("VBA/_VBA_PROJECT", vba_project_stream_bytes),
        ("VBA/dir", dir_container.as_slice()),
        ("VBA/Module1", module_container.as_slice()),
        ("\u{0005}DigitalSignature", pkcs7_der.as_slice()),
    ];
    let vba_project_bin = build_vba_project_bin_with_streams(&signed_streams);

    // ---- 5) Verify ----
    let sig = verify_vba_digital_signature(&vba_project_bin)
        .expect("signature inspection should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(
        sig.binding,
        VbaSignatureBinding::Bound,
        "expected signature binding to be Bound even when DigestInfo.digestAlgorithm is sha256 but digest bytes are MD5"
    );
}

#[test]
fn verify_vba_project_signature_binding_md5_digest_bytes_even_when_oid_is_sha256() {
    use openssl::pkcs7::{Pkcs7, Pkcs7Flags};
    use openssl::pkey::PKey;
    use openssl::stack::Stack;
    use openssl::x509::X509;

    // ---- 1) Build a minimal spec-ish VBA project and compute its MS-OVBA content hash (MD5). ----
    let module_source = b"Sub Hello()\r\nEnd Sub\r\n";
    let module_container = compress_container(module_source);

    // Minimal `dir` stream (decompressed form) with a single module.
    let dir_decompressed = {
        let mut out = Vec::new();
        // PROJECTNAME
        push_record(&mut out, 0x0004, b"VBAProject");
        // MODULENAME
        push_record(&mut out, 0x0019, b"Module1");
        // MODULESTREAMNAME + reserved u16
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"Module1");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);
        // MODULETYPE (standard)
        push_record(&mut out, 0x0021, &0u16.to_le_bytes());
        // MODULETEXTOFFSET
        push_record(&mut out, 0x0031, &0u32.to_le_bytes());
        out
    };
    let dir_container = compress_container(&dir_decompressed);

    let project_stream_bytes: &[u8] = b"Name=\"VBAProject\"\r\nModule=Module1\r\n";
    let vba_project_stream_bytes: &[u8] = b"dummy";

    let unsigned_vba_project_bin = build_vba_project_bin_with_streams(&[
        ("PROJECT", project_stream_bytes),
        ("VBA/_VBA_PROJECT", vba_project_stream_bytes),
        ("VBA/dir", &dir_container),
        ("VBA/Module1", &module_container),
    ]);

    let normalized =
        content_normalized_data(&unsigned_vba_project_bin).expect("ContentNormalizedData");
    let project_md5: [u8; 16] = Md5::digest(&normalized).into();
    assert_eq!(project_md5.len(), 16, "VBA project digest must be 16-byte MD5");

    // ---- 2) Construct SpcIndirectDataContent with sha256 OID but MD5 digest bytes. ----
    let spc_indirect_data_content =
        build_spc_indirect_data_content_sha256_oid_with_md5_digest(&project_md5);

    // ---- 3) Produce PKCS#7 SignedData using OpenSSL (signing with SHA-256 by default). ----
    let pkey = PKey::private_key_from_pem(TEST_KEY_PEM.as_bytes()).expect("parse private key");
    let cert = X509::from_pem(TEST_CERT_PEM.as_bytes()).expect("parse certificate");
    let extra_certs = Stack::new().expect("create cert stack");

    let pkcs7 = Pkcs7::sign(
        &cert,
        &pkey,
        &extra_certs,
        &spc_indirect_data_content,
        Pkcs7Flags::BINARY | Pkcs7Flags::NOATTR,
    )
    .expect("pkcs7 sign");
    let pkcs7_der = pkcs7.to_der().expect("pkcs7 DER");

    // ---- 4) Store signature in a separate signature OLE container (like `vbaProjectSignature.bin`). ----
    let signature_container_bin =
        build_vba_project_bin_with_streams(&[("\u{0005}DigitalSignature", pkcs7_der.as_slice())]);

    // ---- 5) Verify binding ----
    let binding = verify_vba_project_signature_binding(&unsigned_vba_project_bin, &signature_container_bin)
        .expect("binding verification should succeed");

    let debug = match binding {
        VbaProjectBindingVerification::BoundVerified(debug) => debug,
        other => panic!("expected BoundVerified, got {:?}", other),
    };

    assert_eq!(
        debug.hash_algorithm_oid.as_deref(),
        Some("2.16.840.1.101.3.4.2.1")
    );
    assert_eq!(debug.hash_algorithm_name.as_deref(), Some("SHA-256"));
    assert_eq!(debug.signed_digest.as_deref(), Some(project_md5.as_ref()));
    assert_eq!(debug.computed_digest.as_deref(), Some(project_md5.as_ref()));
}

fn build_vba_project_bin_with_streams(streams: &[(&str, &[u8])]) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create compound file");
    ole.create_storage("VBA").expect("create VBA storage");

    for (path, bytes) in streams {
        let mut stream = ole.create_stream(path).expect("create stream");
        stream.write_all(bytes).expect("write bytes");
    }

    ole.into_inner().into_inner()
}

fn push_record(out: &mut Vec<u8>, id: u16, data: &[u8]) {
    out.extend_from_slice(&id.to_le_bytes());
    out.extend_from_slice(&(data.len() as u32).to_le_bytes());
    out.extend_from_slice(data);
}

fn build_spc_indirect_data_content_sha256_oid_with_md5_digest(md5_digest: &[u8]) -> Vec<u8> {
    // AlgorithmIdentifier ::= SEQUENCE { algorithm OBJECT IDENTIFIER, parameters NULL }
    let sha256_oid = der_oid(&[0x60, 0x86, 0x48, 0x01, 0x65, 0x03, 0x04, 0x02, 0x01]);
    let alg_id = der_sequence(&[sha256_oid, der_null()]);

    // DigestInfo ::= SEQUENCE { digestAlgorithm AlgorithmIdentifier, digest OCTET STRING }
    let digest_info = der_sequence(&[alg_id, der_octet_string(md5_digest)]);

    // SpcAttributeTypeAndOptionalValue ::= SEQUENCE { type OID, value [0] EXPLICIT ANY OPTIONAL }
    // For VBA signatures the precise `type` value is not relevant to this regression; we only
    // care about `messageDigest`.
    let dummy_type_oid = der_oid(&[
        0x2b, 0x06, 0x01, 0x04, 0x01, 0x82, 0x37, 0x02, 0x01, 0x1e,
    ]); // 1.3.6.1.4.1.311.2.1.30 (SpcSipInfo)
    let data = der_sequence(&[dummy_type_oid]);

    // SpcIndirectDataContent ::= SEQUENCE { data, messageDigest }
    der_sequence(&[data, digest_info])
}

fn der_sequence(items: &[Vec<u8>]) -> Vec<u8> {
    let mut content = Vec::new();
    for item in items {
        content.extend_from_slice(item);
    }
    der_tlv(0x30, &content)
}

fn der_oid(oid_content: &[u8]) -> Vec<u8> {
    der_tlv(0x06, oid_content)
}

fn der_null() -> Vec<u8> {
    vec![0x05, 0x00]
}

fn der_octet_string(bytes: &[u8]) -> Vec<u8> {
    der_tlv(0x04, bytes)
}

fn der_tlv(tag: u8, content: &[u8]) -> Vec<u8> {
    let mut out = Vec::new();
    out.push(tag);
    out.extend_from_slice(&der_len(content.len()));
    out.extend_from_slice(content);
    out
}

fn der_len(len: usize) -> Vec<u8> {
    if len < 0x80 {
        return vec![len as u8];
    }
    let mut buf = Vec::new();
    let mut n = len;
    while n > 0 {
        buf.push((n & 0xFF) as u8);
        n >>= 8;
    }
    buf.reverse();
    let mut out = Vec::with_capacity(1 + buf.len());
    out.push(0x80 | (buf.len() as u8));
    out.extend_from_slice(&buf);
    out
}
