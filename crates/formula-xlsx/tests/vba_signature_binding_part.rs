#![cfg(not(target_arch = "wasm32"))]

use std::io::{Cursor, Write};

use formula_vba::{
    compress_container, content_normalized_data, VbaSignatureBinding, VbaSignatureVerification,
};
use formula_xlsx::XlsxPackage;
use openssl::hash::{hash, MessageDigest};

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
        // NOATTR keeps the output deterministic (avoids adding a SigningTime attribute).
        Pkcs7Flags::BINARY | Pkcs7Flags::DETACHED | Pkcs7Flags::NOATTR,
    )
    .expect("pkcs7 sign");
    pkcs7.to_der().expect("pkcs7 DER")
}

fn build_minimal_vba_project_bin(module1: &[u8]) -> Vec<u8> {
    fn push_record(out: &mut Vec<u8>, id: u16, data: &[u8]) {
        out.extend_from_slice(&id.to_le_bytes());
        out.extend_from_slice(&(data.len() as u32).to_le_bytes());
        out.extend_from_slice(data);
    }

    // `content_normalized_data` expects a decompressed-and-parsable `VBA/dir` stream and module
    // streams containing MS-OVBA compressed containers.
    let module_container = compress_container(module1);

    let dir_decompressed = {
        let mut out = Vec::new();
        // Minimal module record group.
        push_record(&mut out, 0x0019, b"Module1"); // MODULENAME

        // MODULESTREAMNAME + reserved u16.
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"Module1");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);

        // MODULETYPE (standard)
        push_record(&mut out, 0x0021, &0u16.to_le_bytes());
        // MODULETEXTOFFSET: our module stream is just the compressed container.
        push_record(&mut out, 0x0031, &0u32.to_le_bytes());
        out
    };
    let dir_container = compress_container(&dir_decompressed);

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");

    {
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(b"Name=\"VBAProject\"\r\nModule=Module1\r\n")
            .expect("write PROJECT");
    }

    ole.create_storage("VBA").expect("VBA storage");

    {
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }

    {
        let mut s = ole.create_stream("VBA/Module1").expect("module stream");
        s.write_all(&module_container).expect("write module");
    }

    ole.into_inner().into_inner()
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

fn build_spc_indirect_data_content_md5(project_digest: &[u8]) -> Vec<u8> {
    // MD5 OID: 1.2.840.113549.2.5
    let md5_oid = [0x2A, 0x86, 0x48, 0x86, 0xF7, 0x0D, 0x02, 0x05];

    let mut alg_id = Vec::new();
    alg_id.extend_from_slice(&der_oid_raw(&md5_oid));
    alg_id.extend_from_slice(&der_null());
    let alg_id = der_sequence(&alg_id);

    let mut digest_info = Vec::new();
    digest_info.extend_from_slice(&alg_id);
    digest_info.extend_from_slice(&der_octet_string(project_digest));
    let digest_info = der_sequence(&digest_info);

    let mut spc = Vec::new();
    // `data` (ignored by our parser) – use NULL.
    spc.extend_from_slice(&der_null());
    spc.extend_from_slice(&digest_info);
    der_sequence(&spc)
}

fn build_ole_with_streams(streams: &[(&str, &[u8])]) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create compound file");

    for (path, bytes) in streams {
        let mut stream = ole.create_stream(path).expect("create stream");
        stream.write_all(bytes).expect("write stream bytes");
    }

    ole.into_inner().into_inner()
}

fn build_vba_signature_ole(signature_blob: &[u8]) -> Vec<u8> {
    build_ole_with_streams(&[("\u{0005}DigitalSignature", signature_blob)])
}

fn build_xlsm_zip(vba_project_bin: &[u8], vba_project_signature_bin: &[u8]) -> Vec<u8> {
    let vba_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSig" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature" Target="vbaProjectSignature.bin"/>
</Relationships>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options =
        zip::write::FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    for (name, bytes) in [
        ("xl/vbaProject.bin", vba_project_bin),
        ("xl/_rels/vbaProject.bin.rels", vba_rels.as_slice()),
        ("xl/vbaProjectSignature.bin", vba_project_signature_bin),
    ] {
        zip.start_file(name, options).expect("start zip file");
        zip.write_all(bytes).expect("write zip file");
    }

    zip.finish().expect("finish zip").into_inner()
}

#[test]
fn verifies_external_signature_part_binding_against_vba_project_bin() {
    let module1 = b"module1-bytes";
    let vba_project_bin = build_minimal_vba_project_bin(module1);
    let normalized = content_normalized_data(&vba_project_bin).expect("content normalized data");
    let digest = hash(MessageDigest::md5(), &normalized)
        .expect("md5 digest")
        .to_vec();

    let signed_content = build_spc_indirect_data_content_md5(&digest);
    let pkcs7 = make_pkcs7_detached_signature(&signed_content);

    let mut signature_stream = signed_content.clone();
    signature_stream.extend_from_slice(&pkcs7);
    let signature_part = build_vba_signature_ole(&signature_stream);

    // Untampered project: binding should verify.
    let xlsm_bytes = build_xlsm_zip(&vba_project_bin, &signature_part);
    let pkg = XlsxPackage::from_bytes(&xlsm_bytes).expect("read xlsm");

    let sig = pkg
        .verify_vba_digital_signature()
        .expect("verify signature")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.binding, VbaSignatureBinding::Bound);

    // Tamper with a project stream but keep the signature payload intact:
    // PKCS#7 verification should still succeed, but binding must fail.
    let mut tampered_module = module1.to_vec();
    tampered_module[0] ^= 0xFF;
    let tampered_project = build_minimal_vba_project_bin(&tampered_module);
    let xlsm_tampered = build_xlsm_zip(&tampered_project, &signature_part);
    let pkg = XlsxPackage::from_bytes(&xlsm_tampered).expect("read tampered xlsm");

    let sig = pkg
        .verify_vba_digital_signature()
        .expect("verify signature")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.binding, VbaSignatureBinding::NotBound);
}
