#![cfg(all(feature = "vba", not(target_arch = "wasm32")))]

use std::io::{Cursor, Write};

use formula_vba::{
    compress_container, content_normalized_data, verify_vba_digital_signature,
    VbaProjectBindingVerification, VbaSignatureVerification,
};
use formula_xlsx::XlsxPackage;
use openssl::hash::{hash, MessageDigest};
use zip::write::FileOptions;

mod vba_signature_test_utils;
use vba_signature_test_utils::{build_vba_signature_ole, make_pkcs7_detached_signature};

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

    // VBA signatures sign an Authenticode `SpcIndirectDataContent` whose DigestInfo.digest bytes
    // are the MS-OVBA v1 Content Hash (`MD5(ContentNormalizedData)`; MS-OVBA §2.4.2.3).
    //
    // Per MS-OSHARED §4.3, Office stores these MD5 digest bytes for legacy signature streams even
    // when DigestInfo.digestAlgorithm.algorithm advertises SHA-256. We store it as:
    //   signed_content || pkcs7_detached_signature(signed_content)
    let signed_content = build_spc_indirect_data_content_sha256_oid_with_md5_digest(digest.as_ref());
    let pkcs7 = make_pkcs7_detached_signature(&signed_content);
    let mut signature_stream_payload = signed_content.clone();
    signature_stream_payload.extend_from_slice(&pkcs7);

    let signature_ole = build_vba_signature_ole(&signature_stream_payload);

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
