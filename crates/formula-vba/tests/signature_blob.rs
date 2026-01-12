#![cfg(not(target_arch = "wasm32"))]

use formula_vba::{verify_vba_signature_blob, VbaSignatureVerification};

mod signature_test_utils;

use signature_test_utils::make_pkcs7_signed_message;

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
fn verifies_raw_signature_blob_and_extracts_signer_subject() {
    let pkcs7 = make_pkcs7_signed_message(b"formula-vba-test");

    let (verification, signer_subject) = verify_vba_signature_blob(&pkcs7);

    assert_eq!(verification, VbaSignatureVerification::SignedVerified);
    assert!(
        signer_subject
            .as_deref()
            .is_some_and(|s| s.contains("Formula VBA Test")),
        "expected signer subject to mention test CN, got: {signer_subject:?}"
    );
}

#[test]
fn verifies_raw_signature_blob_wrapped_in_digsig_info_serialized() {
    let pkcs7 = make_pkcs7_signed_message(b"formula-vba-test");
    let wrapped = wrap_in_digsig_info_serialized(&pkcs7);

    let (verification, signer_subject) = verify_vba_signature_blob(&wrapped);

    assert_eq!(verification, VbaSignatureVerification::SignedVerified);
    assert!(
        signer_subject
            .as_deref()
            .is_some_and(|s| s.contains("Formula VBA Test")),
        "expected signer subject to mention test CN, got: {signer_subject:?}"
    );
}
