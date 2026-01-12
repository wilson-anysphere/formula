use formula_vba::extract_vba_signature_signed_digest;

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
