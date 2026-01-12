#![cfg(all(feature = "vba", not(target_arch = "wasm32")))]

use std::io::{Cursor, Write};

use formula_vba::{
    compress_container, content_normalized_data, verify_vba_digital_signature,
    VbaProjectBindingVerification, VbaSignatureVerification,
};
use formula_xlsx::XlsxPackage;
use openssl::hash::{hash, MessageDigest};
use zip::write::FileOptions;

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
    let mut out = Vec::new();
    let mut n = len;
    let mut buf = Vec::new();
    while n > 0 {
        buf.push((n & 0xFF) as u8);
        n >>= 8;
    }
    buf.reverse();
    out.push(0x80 | (buf.len() as u8));
    out.extend_from_slice(&buf);
    out
}

fn der_tlv(tag: u8, value: &[u8]) -> Vec<u8> {
    let mut out = Vec::new();
    out.push(tag);
    out.extend_from_slice(&der_len(value.len()));
    out.extend_from_slice(value);
    out
}

fn der_sequence(content: &[u8]) -> Vec<u8> {
    der_tlv(0x30, content)
}

fn der_null() -> Vec<u8> {
    vec![0x05, 0x00]
}

fn der_oid_raw(oid: &[u8]) -> Vec<u8> {
    der_tlv(0x06, oid)
}

fn der_octet_string(bytes: &[u8]) -> Vec<u8> {
    der_tlv(0x04, bytes)
}

fn build_spc_indirect_data_content_sha256_oid_with_md5_digest(md5_digest: &[u8]) -> Vec<u8> {
    // SHA-256 OID: 2.16.840.1.101.3.4.2.1
    let sha256_oid = [0x60, 0x86, 0x48, 0x01, 0x65, 0x03, 0x04, 0x02, 0x01];
    build_spc_indirect_data_content(&sha256_oid, md5_digest)
}

fn build_spc_indirect_data_content(oid: &[u8], digest: &[u8]) -> Vec<u8> {
    let mut alg_id = Vec::new();
    alg_id.extend_from_slice(&der_oid_raw(oid));
    alg_id.extend_from_slice(&der_null());
    let alg_id = der_sequence(&alg_id);

    let mut digest_info = Vec::new();
    digest_info.extend_from_slice(&alg_id);
    digest_info.extend_from_slice(&der_octet_string(digest));
    let digest_info = der_sequence(&digest_info);

    let mut spc = Vec::new();
    // `data` (ignored by our parser) – use NULL.
    spc.extend_from_slice(&der_null());
    spc.extend_from_slice(&digest_info);
    der_sequence(&spc)
}

fn build_vba_project_bin(module_byte: u8) -> Vec<u8> {
    fn push_record(out: &mut Vec<u8>, id: u16, data: &[u8]) {
        out.extend_from_slice(&id.to_le_bytes());
        out.extend_from_slice(&(data.len() as u32).to_le_bytes());
        out.extend_from_slice(data);
    }

    // Minimal MS-OVBA-ish VBA project structure that is valid for `content_normalized_data`.
    let module_source = {
        let mut out = Vec::new();
        out.extend_from_slice(b"Sub Hello");
        out.push(module_byte);
        out.extend_from_slice(b"()\r\nEnd Sub\r\n");
        out
    };
    let module_container = compress_container(&module_source);

    // Minimal `dir` stream (decompressed form) describing a single module named `Module1`.
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

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create compound file");

    {
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(b"Name=\"VBAProject\"\r\nModule=Module1\r\n")
            .unwrap();
    }

    ole.create_storage("VBA").expect("VBA storage");

    {
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).unwrap();
    }
    {
        let mut s = ole.create_stream("VBA/Module1").expect("module stream");
        s.write_all(&module_container).unwrap();
    }

    ole.into_inner().into_inner()
}

fn build_signature_part(signature_stream_payload: &[u8]) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create compound file");
    {
        let mut s = ole
            .create_stream("\u{0005}DigitalSignature")
            .expect("DigitalSignature stream");
        s.write_all(signature_stream_payload).unwrap();
    }
    ole.into_inner().into_inner()
}

fn build_xlsm_with_external_signature(project_ole: &[u8], signature_ole: &[u8]) -> Vec<u8> {
    let rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature" Target="vbaProjectSignature.bin"/>
</Relationships>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options =
        FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("xl/vbaProject.bin", options).unwrap();
    zip.write_all(project_ole).unwrap();

    zip.start_file("xl/_rels/vbaProject.bin.rels", options)
        .unwrap();
    zip.write_all(rels).unwrap();

    zip.start_file("xl/vbaProjectSignature.bin", options).unwrap();
    zip.write_all(signature_ole).unwrap();

    zip.finish().unwrap().into_inner()
}

#[test]
fn verifies_project_digest_binding_with_external_signature_part() {
    let project_ole = build_vba_project_bin(b'A');
    let normalized = content_normalized_data(&project_ole).expect("ContentNormalizedData");
    let digest = hash(MessageDigest::md5(), &normalized).expect("md5 digest");
    assert_eq!(digest.as_ref().len(), 16, "expected 16-byte MD5 digest");

    // VBA signatures sign an Authenticode `SpcIndirectDataContent` whose DigestInfo digest bytes
    // are the MS-OVBA project ContentNormalizedData hash (MD5), even when the DigestInfo algorithm
    // OID indicates SHA-256. We store it as:
    //   signed_content || pkcs7_detached_signature(signed_content)
    let signed_content = build_spc_indirect_data_content_sha256_oid_with_md5_digest(digest.as_ref());
    let pkcs7 = make_pkcs7_detached_signature(&signed_content);
    let mut signature_stream_payload = signed_content.clone();
    signature_stream_payload.extend_from_slice(&pkcs7);

    let signature_ole = build_signature_part(&signature_stream_payload);

    // Sanity: the PKCS#7 blob verifies even after we mutate the project bytes later.
    let sig = verify_vba_digital_signature(&signature_ole)
        .expect("signature verification should succeed")
        .expect("signature should be present");
    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);

    let xlsm = build_xlsm_with_external_signature(&project_ole, &signature_ole);
    let pkg = XlsxPackage::from_bytes(&xlsm).expect("read xlsm package");
    let binding = pkg
        .vba_project_signature_binding()
        .expect("binding verification")
        .expect("project present");
    assert!(
        matches!(binding, VbaProjectBindingVerification::BoundVerified(_)),
        "expected BoundVerified, got {binding:?}"
    );

    // Mutate a covered project stream (module byte) and ensure binding fails even though the
    // PKCS#7 signature remains valid (it signs the old digest).
    let mutated_project = build_vba_project_bin(b'B');
    let xlsm2 = build_xlsm_with_external_signature(&mutated_project, &signature_ole);
    let pkg2 = XlsxPackage::from_bytes(&xlsm2).expect("read xlsm package");
    let binding2 = pkg2
        .vba_project_signature_binding()
        .expect("binding verification")
        .expect("project present");
    assert!(
        matches!(binding2, VbaProjectBindingVerification::BoundMismatch(_)),
        "unexpected binding result after mutation: {binding2:?}"
    );

    let sig2 = verify_vba_digital_signature(&signature_ole)
        .expect("signature verification should succeed")
        .expect("signature should be present");
    assert_eq!(sig2.verification, VbaSignatureVerification::SignedVerified);
}
