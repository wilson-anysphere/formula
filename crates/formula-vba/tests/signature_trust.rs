#![cfg(not(target_arch = "wasm32"))]

mod signature_test_utils;

use formula_vba::{
    verify_vba_digital_signature_with_trust, VbaCertificateTrust, VbaSignatureTrustOptions,
    VbaSignatureVerification,
};
use openssl::x509::X509;

use signature_test_utils::{
    build_vba_project_bin_with_signature, make_pkcs7_signed_message, make_unrelated_root_cert_der,
    TEST_CERT_PEM,
};

#[test]
fn trust_is_unknown_when_no_roots_provided() {
    let pkcs7 = make_pkcs7_signed_message(b"formula-vba-trust-test");
    let vba = build_vba_project_bin_with_signature(Some(&pkcs7));

    let options = VbaSignatureTrustOptions {
        trusted_root_certs_der: Vec::new(),
    };

    let sig = verify_vba_digital_signature_with_trust(&vba, &options)
        .expect("signature inspection should succeed")
        .expect("signature should be present");

    assert_eq!(sig.signature.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.cert_trust, VbaCertificateTrust::Unknown);
}

#[test]
fn trust_is_trusted_when_root_matches_signer() {
    let pkcs7 = make_pkcs7_signed_message(b"formula-vba-trust-test");
    let vba = build_vba_project_bin_with_signature(Some(&pkcs7));

    let cert_der = X509::from_pem(TEST_CERT_PEM.as_bytes())
        .expect("parse test certificate")
        .to_der()
        .expect("convert test certificate to DER");

    let options = VbaSignatureTrustOptions {
        trusted_root_certs_der: vec![cert_der],
    };

    let sig = verify_vba_digital_signature_with_trust(&vba, &options)
        .expect("signature inspection should succeed")
        .expect("signature should be present");

    assert_eq!(sig.signature.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.cert_trust, VbaCertificateTrust::Trusted);
}

#[test]
fn trust_is_untrusted_when_root_does_not_match_signer() {
    let pkcs7 = make_pkcs7_signed_message(b"formula-vba-trust-test");
    let vba = build_vba_project_bin_with_signature(Some(&pkcs7));

    let wrong_root_der = make_unrelated_root_cert_der();

    let options = VbaSignatureTrustOptions {
        trusted_root_certs_der: vec![wrong_root_der],
    };

    let sig = verify_vba_digital_signature_with_trust(&vba, &options)
        .expect("signature inspection should succeed")
        .expect("signature should be present");

    assert_eq!(sig.signature.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.cert_trust, VbaCertificateTrust::Untrusted);
}
