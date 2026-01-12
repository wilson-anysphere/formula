#![cfg(all(feature = "vba", not(target_arch = "wasm32")))]

use std::io::{Cursor, Write};

use formula_vba::{
    compress_container, compute_vba_project_digest, DigestAlg, VbaSignatureBinding,
    VbaSignatureVerification,
};
use formula_xlsx::XlsxPackage;

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
        // NOATTR keeps the output deterministic (avoids adding a SigningTime attribute).
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

fn build_minimal_vba_project_bin(module1: &[u8]) -> Vec<u8> {
    build_minimal_vba_project_bin_impl(module1, None)
}

fn build_minimal_vba_project_bin_with_signature(module1: &[u8], signature_blob: &[u8]) -> Vec<u8> {
    build_minimal_vba_project_bin_impl(module1, Some(signature_blob))
}

fn build_minimal_vba_project_bin_impl(module1: &[u8], signature_blob: Option<&[u8]>) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");

    // `compute_vba_project_digest` expects a parsable/decompressible `VBA/dir` stream and module
    // streams containing MS-OVBA compressed containers.
    let module_container = compress_container(module1);

    {
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(b"Name=\"VBAProject\"\r\nModule=Module1\r\n")
            .expect("write PROJECT");
    }

    ole.create_storage("VBA").expect("VBA storage");

    {
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
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }

    {
        // Used by Office and present in most real projects; not required by our binding logic
        // directly, but makes the fixture a closer match to real files.
        let mut s = ole
            .create_stream("VBA/_VBA_PROJECT")
            .expect("_VBA_PROJECT stream");
        s.write_all(b"dummy").expect("write _VBA_PROJECT");
    }

    {
        let mut s = ole.create_stream("VBA/Module1").expect("module stream");
        s.write_all(&module_container).expect("write module");
    }

    if let Some(sig) = signature_blob {
        let mut s = ole
            .create_stream("\u{0005}DigitalSignature")
            .expect("signature stream");
        s.write_all(sig).expect("write signature");
    }

    ole.into_inner().into_inner()
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

fn build_oshared_digsig_blob(valid_pkcs7: &[u8]) -> Vec<u8> {
    // MS-OSHARED describes a DigSigBlob wrapper around the PKCS#7 signature bytes.
    //
    // This test blob intentionally contains *multiple* PKCS#7 SignedData blobs:
    // - a corrupted (but still parseable) one early, and
    // - a corrupted one after the real signature.
    //
    // This ensures verification succeeds only when the DigSigBlob offsets are honored, not when
    // relying on the heuristic scan-for-0x30 fallback.
    let mut invalid_pkcs7 = valid_pkcs7.to_vec();
    let last = invalid_pkcs7.len().saturating_sub(1);
    invalid_pkcs7[last] ^= 0xFF;

    let digsig_blob_header_len = 8usize; // cb + serializedPointer
    let digsig_info_len = 0x24usize; // 9 DWORDs (cbSignature/signatureOffset + 7 reserved)

    let invalid_offset = digsig_blob_header_len + digsig_info_len; // 0x2C
    assert_eq!(invalid_offset, 0x2C);

    // Place the valid signature after the invalid one and align to 4 bytes.
    let mut signature_offset = invalid_offset + invalid_pkcs7.len();
    signature_offset = (signature_offset + 3) & !3;

    let cb_signature = u32::try_from(valid_pkcs7.len()).expect("pkcs7 fits u32");
    let signature_offset_u32 = u32::try_from(signature_offset).expect("offset fits u32");

    let mut out = Vec::new();
    // DigSigBlob.cb placeholder (filled later) + serializedPointer = 8.
    out.extend_from_slice(&0u32.to_le_bytes());
    out.extend_from_slice(&8u32.to_le_bytes());

    // DigSigInfoSerialized (MS-OSHARED): we only care about cbSignature and signatureOffset.
    out.extend_from_slice(&cb_signature.to_le_bytes());
    out.extend_from_slice(&signature_offset_u32.to_le_bytes());
    // Remaining fields set to 0 (cert store/project name/timestamp URL).
    for _ in 0..7 {
        out.extend_from_slice(&0u32.to_le_bytes());
    }
    assert_eq!(out.len(), invalid_offset);

    // Insert a corrupted PKCS#7 blob early in the stream (to break scan-first heuristics).
    out.extend_from_slice(&invalid_pkcs7);

    // Pad up to signatureOffset and append the actual signature bytes.
    if out.len() < signature_offset {
        out.resize(signature_offset, 0);
    }
    out.extend_from_slice(valid_pkcs7);

    // Append an invalid PKCS#7 blob after the real signature (to break scan-last heuristics).
    out.extend_from_slice(&invalid_pkcs7);

    // DigSigBlob.cb: size of the serialized signatureInfo payload (excluding the initial DWORDs).
    let cb =
        u32::try_from(out.len().saturating_sub(digsig_blob_header_len)).expect("blob fits u32");
    out[0..4].copy_from_slice(&cb.to_le_bytes());

    out
}

fn build_xlsm_zip(vba_project_bin: &[u8], vba_project_signature_bin: &[u8]) -> Vec<u8> {
    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>
  <Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>
  <Override PartName="/xl/vbaProjectSignature.bin" ContentType="application/vnd.ms-office.vbaProjectSignature"/>
</Types>"#;

    let root_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#;

    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets></sheets>
</workbook>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdVba" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>
</Relationships>"#;

    let vba_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSig" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature" Target="vbaProjectSignature.bin"/>
</Relationships>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    for (name, bytes) in [
        ("[Content_Types].xml", content_types.as_bytes()),
        ("_rels/.rels", root_rels.as_bytes()),
        ("xl/workbook.xml", workbook_xml.as_bytes()),
        ("xl/_rels/workbook.xml.rels", workbook_rels.as_bytes()),
        ("xl/vbaProject.bin", vba_project_bin),
        ("xl/_rels/vbaProject.bin.rels", vba_rels.as_bytes()),
        ("xl/vbaProjectSignature.bin", vba_project_signature_bin),
    ] {
        zip.start_file(name, options).expect("start zip file");
        zip.write_all(bytes).expect("write zip file");
    }

    zip.finish().expect("finish zip").into_inner()
}

#[test]
fn verify_prefers_vba_project_signature_part() {
    let pkcs7 = make_pkcs7_signed_message(b"formula-xlsx-test");
    let signature_part = build_vba_signature_ole(&pkcs7);

    // Use dummy bytes for `vbaProject.bin` to ensure verification doesn't try to open it if the
    // signature part is present and valid.
    let xlsm_bytes = build_xlsm_zip(b"not-an-ole", &signature_part);
    let pkg = XlsxPackage::from_bytes(&xlsm_bytes).expect("read xlsm");

    let sig = pkg
        .verify_vba_digital_signature()
        .expect("verify signature")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
}

#[test]
fn verify_signature_part_with_digsig_blob_wrapper() {
    let pkcs7 = make_pkcs7_signed_message(b"formula-xlsx-test");
    let wrapped = build_oshared_digsig_blob(&pkcs7);
    let signature_part = build_vba_signature_ole(&wrapped);

    // Use dummy bytes for `vbaProject.bin` to ensure verification doesn't try to open it if the
    // signature part is present and valid.
    let xlsm_bytes = build_xlsm_zip(b"not-an-ole", &signature_part);
    let pkg = XlsxPackage::from_bytes(&xlsm_bytes).expect("read xlsm");

    let sig = pkg
        .verify_vba_digital_signature()
        .expect("verify signature")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
}

#[test]
fn verify_falls_back_to_vba_project_bin_when_signature_part_is_garbage() {
    let pkcs7 = make_pkcs7_signed_message(b"formula-xlsx-test");
    let vba_project_bin = build_vba_signature_ole(&pkcs7);

    // Non-OLE garbage: signature verification should fall back to `vbaProject.bin`.
    let xlsm_bytes = build_xlsm_zip(&vba_project_bin, b"not-an-ole");
    let pkg = XlsxPackage::from_bytes(&xlsm_bytes).expect("read xlsm");

    let sig = pkg
        .verify_vba_digital_signature()
        .expect("verify signature")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
}

#[test]
fn verify_signature_part_binding_matches_vba_project_bin() {
    let vba_project_bin = build_minimal_vba_project_bin(b"Sub Hello()\r\nEnd Sub\r\n");
    let digest = compute_vba_project_digest(&vba_project_bin, DigestAlg::Md5)
        .expect("compute project digest");
    assert_eq!(digest.len(), 16, "VBA project digest must be 16-byte MD5");
    let spc = make_spc_indirect_data_content_sha256(&digest);

    let pkcs7 = make_pkcs7_signed_message(&spc);
    let signature_part = build_vba_signature_ole(&pkcs7);

    let xlsm_bytes = build_xlsm_zip(&vba_project_bin, &signature_part);
    let pkg = XlsxPackage::from_bytes(&xlsm_bytes).expect("read xlsm");

    let sig = pkg
        .verify_vba_digital_signature()
        .expect("verify signature")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.binding, VbaSignatureBinding::Bound);
}

#[test]
fn verify_signature_part_binding_matches_vba_project_bin_when_digsig_blob_wrapped() {
    let vba_project_bin = build_minimal_vba_project_bin(b"Sub Hello()\r\nEnd Sub\r\n");
    let digest = compute_vba_project_digest(&vba_project_bin, DigestAlg::Md5)
        .expect("compute project digest");
    assert_eq!(digest.len(), 16, "VBA project digest must be 16-byte MD5");
    let spc = make_spc_indirect_data_content_sha256(&digest);

    let pkcs7 = make_pkcs7_signed_message(&spc);
    let wrapped = build_oshared_digsig_blob(&pkcs7);
    let signature_part = build_vba_signature_ole(&wrapped);

    let xlsm_bytes = build_xlsm_zip(&vba_project_bin, &signature_part);
    let pkg = XlsxPackage::from_bytes(&xlsm_bytes).expect("read xlsm");

    let sig = pkg
        .verify_vba_digital_signature()
        .expect("verify signature")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.binding, VbaSignatureBinding::Bound);
}

#[test]
fn verify_signature_part_binding_detects_tampered_vba_project_bin() {
    let vba_project_bin = build_minimal_vba_project_bin(b"Sub Hello()\r\nEnd Sub\r\n");
    let digest = compute_vba_project_digest(&vba_project_bin, DigestAlg::Md5)
        .expect("compute project digest");
    assert_eq!(digest.len(), 16, "VBA project digest must be 16-byte MD5");
    let spc = make_spc_indirect_data_content_sha256(&digest);

    let pkcs7 = make_pkcs7_signed_message(&spc);
    let signature_part = build_vba_signature_ole(&pkcs7);

    // Change the project without changing the signature blob.
    let tampered_project = build_minimal_vba_project_bin(b"Sub HELLO()\r\nEnd Sub\r\n");

    let xlsm_bytes = build_xlsm_zip(&tampered_project, &signature_part);
    let pkg = XlsxPackage::from_bytes(&xlsm_bytes).expect("read xlsm");

    let sig = pkg
        .verify_vba_digital_signature()
        .expect("verify signature")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.binding, VbaSignatureBinding::NotBound);
}

#[test]
fn verify_falls_back_to_vba_project_bin_when_signature_part_is_garbage_and_embedded_signature_is_digsig_blob_wrapped(
) {
    // Build a minimal, digestable VBA project so binding evaluation can succeed.
    let module_bytes = b"Sub Hello()\r\nEnd Sub\r\n";
    let unsigned_project = build_minimal_vba_project_bin(module_bytes);
    let digest = compute_vba_project_digest(&unsigned_project, DigestAlg::Md5)
        .expect("compute project digest");
    assert_eq!(digest.len(), 16, "VBA project digest must be 16-byte MD5");
    let spc = make_spc_indirect_data_content_sha256(&digest);

    // Build a valid PKCS#7 signature over the SpcIndirectDataContent and wrap it in a DigSigBlob
    // offset table with decoy corrupted PKCS#7 blobs.
    let pkcs7 = make_pkcs7_signed_message(&spc);
    let wrapped = build_oshared_digsig_blob(&pkcs7);

    // Embed the wrapped signature into `vbaProject.bin` itself.
    let signed_project = build_minimal_vba_project_bin_with_signature(module_bytes, &wrapped);

    // Provide a garbage signature part so `XlsxPackage::verify_vba_digital_signature` is forced to
    // fall back to `vbaProject.bin` for signature inspection.
    let xlsm_bytes = build_xlsm_zip(&signed_project, b"not-an-ole");
    let pkg = XlsxPackage::from_bytes(&xlsm_bytes).expect("read xlsm");

    let sig = pkg
        .verify_vba_digital_signature()
        .expect("verify signature")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.binding, VbaSignatureBinding::Bound);
}
