#![cfg(not(target_arch = "wasm32"))]

use std::io::{Cursor, Write};

use formula_vba::{
    compress_container, content_normalized_data, forms_normalized_data, verify_vba_digital_signature,
    VbaSignatureBinding, VbaSignatureVerification,
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
        // NOATTR keeps output deterministic (avoids SigningTime).
        Pkcs7Flags::BINARY | Pkcs7Flags::NOATTR,
    )
    .expect("pkcs7 sign");
    pkcs7.to_der().expect("pkcs7 DER")
}

fn push_record(out: &mut Vec<u8>, id: u16, data: &[u8]) {
    out.extend_from_slice(&id.to_le_bytes());
    out.extend_from_slice(&(data.len() as u32).to_le_bytes());
    out.extend_from_slice(data);
}

fn der_len(len: usize) -> Vec<u8> {
    if len < 0x80 {
        return vec![len as u8];
    }
    let mut tmp = Vec::new();
    let mut n = len;
    while n > 0 {
        tmp.push((n & 0xFF) as u8);
        n >>= 8;
    }
    tmp.reverse();
    let mut out = Vec::with_capacity(1 + tmp.len());
    out.push(0x80 | (tmp.len() as u8));
    out.extend_from_slice(&tmp);
    out
}

fn der_tlv(tag: u8, value: &[u8]) -> Vec<u8> {
    let mut out = Vec::new();
    out.push(tag);
    out.extend_from_slice(&der_len(value.len()));
    out.extend_from_slice(value);
    out
}

fn der_oid(arcs: &[u32]) -> Vec<u8> {
    assert!(arcs.len() >= 2, "OID requires at least two arcs");
    let mut body = Vec::new();
    let first = arcs[0] * 40 + arcs[1];
    body.push(first as u8);
    for &arc in &arcs[2..] {
        let mut stack = Vec::new();
        let mut v = arc;
        stack.push((v & 0x7F) as u8);
        v >>= 7;
        while v > 0 {
            stack.push(((v & 0x7F) as u8) | 0x80);
            v >>= 7;
        }
        stack.reverse();
        body.extend_from_slice(&stack);
    }
    der_tlv(0x06, &body)
}

fn build_spc_indirect_data_content(message_digest_md5: [u8; 16]) -> Vec<u8> {
    // SpcIndirectDataContent ::= SEQUENCE {
    //   data          SpcAttributeTypeAndOptionalValue,
    //   messageDigest DigestInfo
    // }
    // DigestInfo ::= SEQUENCE {
    //   digestAlgorithm AlgorithmIdentifier,
    //   digest          OCTET STRING
    // }
    // AlgorithmIdentifier ::= SEQUENCE { algorithm OBJECT IDENTIFIER, parameters NULL }
    //
    // We keep the `data` field minimal: SEQUENCE { type OID }, omitting the optional value.

    // type = 1.3.6.1.4.1.311.2.1.15 (SPC_PE_IMAGE_DATA_OBJID)
    let data = der_tlv(0x30, &der_oid(&[1, 3, 6, 1, 4, 1, 311, 2, 1, 15]));

    // digestAlgorithm = md5 (1.2.840.113549.2.5) + NULL params
    let mut alg_id = Vec::new();
    alg_id.extend_from_slice(&der_oid(&[1, 2, 840, 113549, 2, 5]));
    alg_id.extend_from_slice(&[0x05, 0x00]); // NULL
    let alg_id = der_tlv(0x30, &alg_id);

    let digest_octet = der_tlv(0x04, &message_digest_md5);
    let mut digest_info_body = Vec::new();
    digest_info_body.extend_from_slice(&alg_id);
    digest_info_body.extend_from_slice(&digest_octet);
    let digest_info = der_tlv(0x30, &digest_info_body);

    let mut outer = Vec::new();
    outer.extend_from_slice(&data);
    outer.extend_from_slice(&digest_info);
    der_tlv(0x30, &outer)
}

fn build_vba_project_bin_with_designer(designer_bytes: &[u8], signature_blob: Option<&[u8]>) -> Vec<u8> {
    // Build minimal VBA project:
    // - PROJECT stream with BaseClass=UserForm1
    // - VBA/dir with module UserForm1
    // - VBA/UserForm1 module stream (compressed)
    // - Root storage UserForm1 with a designer stream (e.g. "\x03VBFrame")
    // - Optional signature stream

    let module_code = b"Attribute VB_Name = \"UserForm1\"\r\nSub Foo()\r\nEnd Sub\r\n";
    let module_container = compress_container(module_code);

    let dir_decompressed = {
        let mut out = Vec::new();
        // PROJECTCODEPAGE (u16 LE)
        push_record(&mut out, 0x0003, &1252u16.to_le_bytes());
        // PROJECTNAME (MBCS bytes)
        push_record(&mut out, 0x0004, b"VBAProject");

        // PROJECTCONSTANTS record payload:
        // SizeOfConstants (u32) + Constants + Reserved (u16) + SizeOfConstantsUnicode (u32) + ConstantsUnicode
        let mut constants = Vec::new();
        constants.extend_from_slice(&0u32.to_le_bytes()); // SizeOfConstants
        constants.extend_from_slice(&0u16.to_le_bytes()); // Reserved
        constants.extend_from_slice(&0u32.to_le_bytes()); // SizeOfConstantsUnicode
        push_record(&mut out, 0x000C, &constants);

        // Module records.
        push_record(&mut out, 0x0019, b"UserForm1"); // MODULENAME
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"UserForm1");
        stream_name.extend_from_slice(&0u16.to_le_bytes()); // reserved u16
        push_record(&mut out, 0x001A, &stream_name); // MODULESTREAMNAME
        push_record(&mut out, 0x0021, &0x0003u16.to_le_bytes()); // MODULETYPE (UserForm)
        push_record(&mut out, 0x0031, &0u32.to_le_bytes()); // MODULETEXTOFFSET
        out
    };
    let dir_container = compress_container(&dir_decompressed);

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");

    // Root PROJECT stream (MBCS).
    {
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(b"BaseClass=UserForm1\r\n")
            .expect("write PROJECT");
    }

    // VBA storage.
    ole.create_storage("VBA").expect("VBA storage");
    {
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }
    {
        let mut s = ole
            .create_stream("VBA/UserForm1")
            .expect("module stream");
        s.write_all(&module_container).expect("write module");
    }

    // Designer storage.
    ole.create_storage("UserForm1")
        .expect("designer storage UserForm1");
    {
        let mut s = ole
            .create_stream("UserForm1/\u{0003}VBFrame")
            .expect("designer stream");
        s.write_all(designer_bytes).expect("write designer bytes");
    }

    // Signature stream.
    if let Some(sig) = signature_blob {
        let mut s = ole
            .create_stream("\u{0005}DigitalSignature")
            .expect("signature stream");
        s.write_all(sig).expect("write signature");
    }

    ole.into_inner().into_inner()
}

#[test]
fn agile_signature_binds_using_forms_normalized_data() {
    let designer_bytes = b"designer-stream-bytes";
    let unsigned = build_vba_project_bin_with_designer(designer_bytes, None);

    let content_normalized = content_normalized_data(&unsigned).expect("ContentNormalizedData");
    let forms_normalized = forms_normalized_data(&unsigned).expect("FormsNormalizedData");

    let mut hasher = Md5::new();
    hasher.update(&content_normalized);
    hasher.update(&forms_normalized);
    let digest: [u8; 16] = hasher.finalize().into();

    let spc = build_spc_indirect_data_content(digest);
    let pkcs7 = make_pkcs7_signed_message(&spc);
    let vba_project_bin = build_vba_project_bin_with_designer(designer_bytes, Some(&pkcs7));

    let sig = verify_vba_digital_signature(&vba_project_bin)
        .expect("verify should succeed")
        .expect("signature should be present");
    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.binding, VbaSignatureBinding::Bound);
}

#[test]
fn agile_signature_is_not_bound_if_designer_storage_changes() {
    let designer_bytes = b"designer-stream-bytes";
    let unsigned = build_vba_project_bin_with_designer(designer_bytes, None);

    let content_normalized = content_normalized_data(&unsigned).expect("ContentNormalizedData");
    let forms_normalized = forms_normalized_data(&unsigned).expect("FormsNormalizedData");

    let mut hasher = Md5::new();
    hasher.update(&content_normalized);
    hasher.update(&forms_normalized);
    let digest: [u8; 16] = hasher.finalize().into();
    let spc = build_spc_indirect_data_content(digest);
    let pkcs7 = make_pkcs7_signed_message(&spc);

    // Mutate the designer bytes after computing the signed digest.
    let mut mutated = designer_bytes.to_vec();
    mutated[0] ^= 0xFF;
    let vba_project_bin = build_vba_project_bin_with_designer(&mutated, Some(&pkcs7));

    let sig = verify_vba_digital_signature(&vba_project_bin)
        .expect("verify should succeed")
        .expect("signature should be present");
    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.binding, VbaSignatureBinding::NotBound);
}
