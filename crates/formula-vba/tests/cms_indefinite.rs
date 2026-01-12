use formula_vba::extract_vba_signature_signed_digest;

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

#[test]
fn extracts_spc_indirect_data_digest_from_ber_indefinite_cms_at_offset_0() {
    let cms = include_bytes!("fixtures/cms_indefinite.der");

    let digest_info = extract_vba_signature_signed_digest(cms)
        .expect("extract should succeed")
        .expect("digest info should be present");

    assert_eq!(digest_info.digest_algorithm_oid, "2.16.840.1.101.3.4.2.1"); // SHA-256
    assert_eq!(digest_info.digest, (0u8..0x20).collect::<Vec<_>>());
}

#[test]
fn extracts_spc_indirect_data_digest_from_ber_indefinite_cms() {
    let cms = include_bytes!("fixtures/cms_indefinite.der");

    // Real-world VBA signature streams sometimes include a small prefix/header before the CMS blob.
    let mut stream = b"VBA\0SIG\0".to_vec();
    stream.extend_from_slice(cms);

    let digest_info = extract_vba_signature_signed_digest(&stream)
        .expect("extract should succeed")
        .expect("digest info should be present");

    assert_eq!(digest_info.digest_algorithm_oid, "2.16.840.1.101.3.4.2.1"); // SHA-256
    assert_eq!(digest_info.digest, (0u8..0x20).collect::<Vec<_>>());
}

#[test]
fn extracts_spc_indirect_data_digest_from_ber_indefinite_cms_wrapped_in_digsig_info_serialized() {
    let cms = include_bytes!("fixtures/cms_indefinite.der");
    let stream = wrap_in_digsig_info_serialized(cms);

    let digest_info = extract_vba_signature_signed_digest(&stream)
        .expect("extract should succeed")
        .expect("digest info should be present");

    assert_eq!(digest_info.digest_algorithm_oid, "2.16.840.1.101.3.4.2.1"); // SHA-256
    assert_eq!(digest_info.digest, (0u8..0x20).collect::<Vec<_>>());
}

#[test]
fn extracts_digest_from_indefinite_length_detached_signeddata_with_indefinite_encap_content_info() {
    // SpcIndirectDataContent (DER) whose digest is 0..31.
    let mut spc = vec![
        0x30, 0x41, // SEQUENCE
        0x30, 0x0c, // SEQUENCE
        0x06, 0x0a, 0x2b, 0x06, 0x01, 0x04, 0x01, 0x82, 0x37, 0x02, 0x01, 0x0f, // 1.3.6.1.4.1.311.2.1.15
        0x30, 0x31, // SEQUENCE
        0x30, 0x0d, // SEQUENCE
        0x06, 0x09, 0x60, 0x86, 0x48, 0x01, 0x65, 0x03, 0x04, 0x02, 0x01, // 2.16.840.1.101.3.4.2.1 (sha256)
        0x05, 0x00, // NULL
        0x04, 0x20, // OCTET STRING (32 bytes)
    ];
    spc.extend(0u8..0x20);

    // Minimal BER-indefinite CMS ContentInfo/SignedData wrapper with:
    // - signedData ContentInfo
    // - SignedData.encapContentInfo encoded as an *indefinite-length* SEQUENCE with no eContent
    //
    // This exercises the \"optional eContent\" handling where the next bytes in an indefinite-length
    // SEQUENCE are the EOC marker.
    let pkcs7 = vec![
        0x30, 0x80, // ContentInfo SEQUENCE (indefinite)
        0x06, 0x09, 0x2A, 0x86, 0x48, 0x86, 0xF7, 0x0D, 0x01, 0x07, 0x02, // OID 1.2.840.113549.1.7.2 (signedData)
        0xA0, 0x80, // [0] EXPLICIT (indefinite)
        0x30, 0x80, // SignedData SEQUENCE (indefinite)
        0x02, 0x01, 0x03, // version INTEGER 3
        0x31, 0x00, // digestAlgorithms SET (empty; we don't validate it here)
        0x30, 0x80, // encapContentInfo SEQUENCE (indefinite)
        0x06, 0x0A, 0x2B, 0x06, 0x01, 0x04, 0x01, 0x82, 0x37, 0x02, 0x01, 0x04, // OID 1.3.6.1.4.1.311.2.1.4 (SpcIndirectDataContent)
        0x00, 0x00, // EOC for encapContentInfo
        0x00, 0x00, // EOC for SignedData
        0x00, 0x00, // EOC for [0] EXPLICIT
        0x00, 0x00, // EOC for ContentInfo
    ];

    // Detached signature stream layout: signed content prefix + CMS signature.
    let mut stream = spc.clone();
    stream.extend_from_slice(&pkcs7);

    let digest_info = extract_vba_signature_signed_digest(&stream)
        .expect("extract should succeed")
        .expect("digest info should be present");
    assert_eq!(digest_info.digest_algorithm_oid, "2.16.840.1.101.3.4.2.1"); // SHA-256
    assert_eq!(digest_info.digest, (0u8..0x20).collect::<Vec<_>>());
}
