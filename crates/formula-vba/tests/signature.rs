#![cfg(not(target_arch = "wasm32"))]

use formula_vba::{
    extract_signer_certificate_info, list_vba_digital_signatures, verify_vba_digital_signature,
    VbaSignatureStreamKind, VbaSignatureVerification,
};

mod signature_test_utils;

use signature_test_utils::{
    build_vba_project_bin_with_signature, build_vba_project_bin_with_signature_streams,
    make_pkcs7_detached_signature, make_pkcs7_signed_message,
};

#[test]
fn extracts_signer_certificate_metadata_from_pkcs7() {
    let pkcs7 = make_pkcs7_signed_message(b"formula-vba-test");
    let info = extract_signer_certificate_info(&pkcs7).expect("expected embedded certificate info");

    assert!(
        info.subject.contains("Formula VBA Test"),
        "expected subject to mention test CN, got: {}",
        info.subject
    );
    assert!(
        !info.issuer.is_empty(),
        "expected issuer to be present, got empty issuer"
    );
    assert_eq!(
        info.issuer, info.subject,
        "test certificate is self-signed, expected issuer == subject"
    );
    assert_eq!(
        info.sha256_fingerprint_hex.len(),
        64,
        "expected SHA-256 fingerprint to be 64 hex chars, got {}",
        info.sha256_fingerprint_hex
    );
    assert!(
        info.sha256_fingerprint_hex
            .chars()
            .all(|c| c.is_ascii_digit() || ('a'..='f').contains(&c)),
        "expected lowercase hex fingerprint, got {}",
        info.sha256_fingerprint_hex
    );
    assert!(
        !info.serial_hex.is_empty(),
        "expected serial_hex to be non-empty"
    );
}

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

fn der_encode_len(len: usize) -> Vec<u8> {
    if len < 128 {
        return vec![len as u8];
    }
    let mut bytes = Vec::new();
    let mut n = len;
    while n > 0 {
        bytes.push((n & 0xFF) as u8);
        n >>= 8;
    }
    bytes.reverse();
    let mut out = Vec::new();
    let _ = out.try_reserve_exact(1usize.saturating_add(bytes.len()));
    out.push(0x80 | (bytes.len() as u8));
    out.extend_from_slice(&bytes);
    out
}

fn der_tlv(tag: u8, value: &[u8]) -> Vec<u8> {
    let mut out = Vec::new();
    out.push(tag);
    out.extend_from_slice(&der_encode_len(value.len()));
    out.extend_from_slice(value);
    out
}

fn der_oid(oid: &str) -> Vec<u8> {
    let parts = oid
        .split('.')
        .filter(|p| !p.is_empty())
        .map(|p| p.parse::<u32>().expect("valid OID component"))
        .collect::<Vec<_>>();
    assert!(parts.len() >= 2, "OID must have at least two components");
    let first = parts[0];
    let second = parts[1];
    assert!(first <= 2, "invalid OID first component");
    if first < 2 {
        assert!(second < 40, "invalid OID second component");
    }

    fn encode_subid(mut n: u32) -> Vec<u8> {
        let mut chunks = Vec::new();
        chunks.push((n & 0x7F) as u8);
        n >>= 7;
        while n > 0 {
            chunks.push(((n & 0x7F) as u8) | 0x80);
            n >>= 7;
        }
        chunks.reverse();
        chunks
    }

    let mut encoded = Vec::new();
    encoded.extend_from_slice(&encode_subid(first * 40 + second));

    for &n in &parts[2..] {
        encoded.extend_from_slice(&encode_subid(n));
    }

    der_tlv(0x06, &encoded)
}

fn der_null() -> Vec<u8> {
    vec![0x05, 0x00]
}

fn der_sequence(contents: &[u8]) -> Vec<u8> {
    der_tlv(0x30, contents)
}

fn der_octet_string(bytes: &[u8]) -> Vec<u8> {
    der_tlv(0x04, bytes)
}

fn make_spc_indirect_data_content(digest_algorithm_oid: &str, digest: &[u8]) -> Vec<u8> {
    // Minimal SpcIndirectDataContent:
    //   SEQUENCE {
    //     data SEQUENCE { type OID },
    //     messageDigest SEQUENCE {
    //       digestAlgorithm SEQUENCE { algorithm OID, parameters NULL },
    //       digest OCTET STRING
    //     }
    //   }
    let spc_data = {
        let mut out = Vec::new();
        out.extend_from_slice(&der_oid("1.2.3.4"));
        der_sequence(&out)
    };

    let digest_info = {
        let alg_id = {
            let mut out = Vec::new();
            out.extend_from_slice(&der_oid(digest_algorithm_oid));
            out.extend_from_slice(&der_null());
            der_sequence(&out)
        };

        let mut out = Vec::new();
        out.extend_from_slice(&alg_id);
        out.extend_from_slice(&der_octet_string(digest));
        der_sequence(&out)
    };

    let mut out = Vec::new();
    out.extend_from_slice(&spc_data);
    out.extend_from_slice(&digest_info);
    der_sequence(&out)
}

#[test]
fn unsigned_project_reports_no_signature() {
    let vba = build_vba_project_bin_with_signature(None);
    let sig = verify_vba_digital_signature(&vba).expect("signature inspection should succeed");
    assert!(sig.is_none(), "expected no signature");
}

#[test]
fn valid_pkcs7_signature_is_reported_as_verified() {
    let pkcs7 = make_pkcs7_signed_message(b"formula-vba-test");
    let vba = build_vba_project_bin_with_signature(Some(&pkcs7));

    let sig = verify_vba_digital_signature(&vba)
        .expect("signature inspection should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert!(
        sig.signer_subject
            .as_deref()
            .is_some_and(|s| s.contains("Formula VBA Test")),
        "expected signer subject to mention test CN, got: {:?}",
        sig.signer_subject
    );
}

#[test]
fn ber_indefinite_pkcs7_signature_is_reported_as_verified() {
    // BER-indefinite SignedData fixture (OpenSSL `cms -stream` style).
    let pkcs7 = include_bytes!("fixtures/cms_indefinite.der");
    let vba = build_vba_project_bin_with_signature(Some(pkcs7));

    let sig = verify_vba_digital_signature(&vba)
        .expect("signature inspection should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
}

#[test]
fn ber_indefinite_pkcs7_signature_with_prefix_is_still_verified() {
    let pkcs7 = include_bytes!("fixtures/cms_indefinite.der");
    let mut prefixed = b"VBA\0SIG\0".to_vec();
    prefixed.extend_from_slice(pkcs7);
    let vba = build_vba_project_bin_with_signature(Some(&prefixed));

    let sig = verify_vba_digital_signature(&vba)
        .expect("signature inspection should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
}

#[test]
fn ber_indefinite_pkcs7_wrapped_in_digsig_info_serialized_is_still_verified() {
    let pkcs7 = include_bytes!("fixtures/cms_indefinite.der");
    let wrapped = wrap_in_digsig_info_serialized(pkcs7);
    let vba = build_vba_project_bin_with_signature(Some(&wrapped));

    let sig = verify_vba_digital_signature(&vba)
        .expect("signature inspection should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
}

#[test]
fn corrupting_signature_bytes_marks_signature_invalid() {
    let mut pkcs7 = make_pkcs7_signed_message(b"formula-vba-test");
    let last = pkcs7.len().saturating_sub(1);
    pkcs7[last] ^= 0xFF;
    let vba = build_vba_project_bin_with_signature(Some(&pkcs7));

    let sig = verify_vba_digital_signature(&vba)
        .expect("signature inspection should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedInvalid);
}

#[test]
fn pkcs7_signature_with_prefix_is_still_verified() {
    let pkcs7 = make_pkcs7_signed_message(b"formula-vba-test");
    let mut prefixed = b"VBA\0SIG\0".to_vec();
    prefixed.extend_from_slice(&pkcs7);
    let vba = build_vba_project_bin_with_signature(Some(&prefixed));

    let sig = verify_vba_digital_signature(&vba)
        .expect("signature inspection should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert!(
        sig.signer_subject
            .as_deref()
            .is_some_and(|s| s.contains("Formula VBA Test")),
        "expected signer subject to mention test CN, got: {:?}",
        sig.signer_subject
    );
}

#[test]
fn pkcs7_signature_wrapped_in_digsig_info_serialized_is_still_verified() {
    let pkcs7 = make_pkcs7_signed_message(b"formula-vba-test");
    let wrapped = wrap_in_digsig_info_serialized(&pkcs7);
    let vba = build_vba_project_bin_with_signature(Some(&wrapped));

    let sig = verify_vba_digital_signature(&vba)
        .expect("signature inspection should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert!(
        sig.signer_subject
            .as_deref()
            .is_some_and(|s| s.contains("Formula VBA Test")),
        "expected signer subject to mention test CN, got: {:?}",
        sig.signer_subject
    );
}

#[test]
fn detached_pkcs7_signature_with_prefixed_content_is_verified() {
    let content = b"formula-vba-detached-test";
    let pkcs7 = make_pkcs7_detached_signature(content);
    let mut blob = content.to_vec();
    blob.extend_from_slice(&pkcs7);
    let vba = build_vba_project_bin_with_signature(Some(&blob));

    let sig = verify_vba_digital_signature(&vba)
        .expect("signature inspection should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert!(
        sig.signer_subject
            .as_deref()
            .is_some_and(|s| s.contains("Formula VBA Test")),
        "expected signer subject to mention test CN, got: {:?}",
        sig.signer_subject
    );
}

#[test]
fn prefers_verified_signature_stream_over_invalid_candidate() {
    let mut invalid = make_pkcs7_signed_message(b"formula-vba-test");
    let last = invalid.len().saturating_sub(1);
    invalid[last] ^= 0xFF;

    let valid = make_pkcs7_signed_message(b"formula-vba-test");

    let vba = build_vba_project_bin_with_signature_streams(&[
        ("\u{0005}DigitalSignature", &invalid),
        ("\u{0005}DigitalSignatureEx", &valid),
    ]);

    let sig = verify_vba_digital_signature(&vba)
        .expect("signature inspection should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(
        sig.stream_kind,
        VbaSignatureStreamKind::DigitalSignatureEx,
        "expected to pick verified signature stream, got {}",
        sig.stream_path
    );
    assert!(
        sig.signer_subject
            .as_deref()
            .is_some_and(|s| s.contains("Formula VBA Test")),
        "expected signer subject to mention test CN, got: {:?}",
        sig.signer_subject
    );
}

#[test]
fn unparseable_signature_stream_is_reported_as_parse_error() {
    let blob = b"not-a-pkcs7".to_vec();
    let vba = build_vba_project_bin_with_signature_streams(&[("\u{0005}DigitalSignature", &blob)]);

    let sig = verify_vba_digital_signature(&vba)
        .expect("signature inspection should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedParseError);
}

#[test]
fn prefers_verified_stream_over_parse_error_candidate() {
    let bad = b"not-a-pkcs7".to_vec();
    let valid = make_pkcs7_signed_message(b"formula-vba-test");

    let vba = build_vba_project_bin_with_signature_streams(&[
        ("\u{0005}DigitalSignature", &bad),
        ("\u{0005}DigitalSignatureEx", &valid),
    ]);

    let sig = verify_vba_digital_signature(&vba)
        .expect("signature inspection should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(
        sig.stream_kind,
        VbaSignatureStreamKind::DigitalSignatureEx,
        "expected to pick verified signature stream, got {}",
        sig.stream_path
    );
}

#[test]
fn prefers_digital_signature_ex_over_legacy_when_both_verify() {
    let legacy = make_pkcs7_signed_message(b"legacy-signed");
    let ex = make_pkcs7_signed_message(b"ex-signed");

    let vba = build_vba_project_bin_with_signature_streams(&[
        ("\u{0005}DigitalSignature", &legacy),
        ("\u{0005}DigitalSignatureEx", &ex),
    ]);

    let sig = verify_vba_digital_signature(&vba)
        .expect("signature inspection should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(
        sig.stream_kind,
        VbaSignatureStreamKind::DigitalSignatureEx,
        "expected DigitalSignatureEx to be treated as authoritative when both legacy and Ex signatures verify, got {}",
        sig.stream_path
    );
    assert_eq!(sig.signature, ex);
}

#[test]
fn prefers_digital_signature_ext_over_ex_when_both_verify() {
    let ex = make_pkcs7_signed_message(b"ex-signed");
    let ext = make_pkcs7_signed_message(b"ext-signed");

    let vba = build_vba_project_bin_with_signature_streams(&[
        ("\u{0005}DigitalSignatureEx", &ex),
        ("\u{0005}DigitalSignatureExt", &ext),
    ]);

    let sig = verify_vba_digital_signature(&vba)
        .expect("signature inspection should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(
        sig.stream_kind,
        VbaSignatureStreamKind::DigitalSignatureExt,
        "expected DigitalSignatureExt to be treated as authoritative when both Ex and Ext signatures verify, got {}",
        sig.stream_path
    );
    assert_eq!(sig.signature, ext);
}

#[test]
fn prefers_digital_signature_ex_over_invalid_ext_candidate() {
    // `DigitalSignatureExt` sorts before `DigitalSignatureEx`, but if the Ext stream is present and
    // parses yet fails verification, we should still select a later verified candidate.
    let ex = make_pkcs7_signed_message(b"ex-signed");
    let mut ext_invalid = make_pkcs7_signed_message(b"ext-signed");
    let last = ext_invalid.len().saturating_sub(1);
    ext_invalid[last] ^= 0xFF;

    let vba = build_vba_project_bin_with_signature_streams(&[
        ("\u{0005}DigitalSignatureExt", &ext_invalid),
        ("\u{0005}DigitalSignatureEx", &ex),
    ]);

    let sig = verify_vba_digital_signature(&vba)
        .expect("signature inspection should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(
        sig.stream_kind,
        VbaSignatureStreamKind::DigitalSignatureEx,
        "expected DigitalSignatureEx to be selected when DigitalSignatureExt is invalid, got {}",
        sig.stream_path
    );
    assert_eq!(sig.signature, ex);
}

#[test]
fn lists_signature_streams_in_ext_ex_legacy_order() {
    let legacy = make_pkcs7_signed_message(b"legacy-signed");
    let ex = make_pkcs7_signed_message(b"ex-signed");
    let ext = make_pkcs7_signed_message(b"ext-signed");

    let vba = build_vba_project_bin_with_signature_streams(&[
        ("\u{0005}DigitalSignature", &legacy),
        ("\u{0005}DigitalSignatureEx", &ex),
        ("\u{0005}DigitalSignatureExt", &ext),
    ]);

    let sigs = list_vba_digital_signatures(&vba).expect("signature enumeration should succeed");
    assert_eq!(sigs.len(), 3, "expected three signature streams");

    // Deterministic Excel-like ordering: newest stream first.
    assert!(
        sigs[0].stream_path.ends_with("\u{0005}DigitalSignatureExt"),
        "unexpected first stream path: {}",
        sigs[0].stream_path
    );
    assert!(
        sigs[1].stream_path.ends_with("\u{0005}DigitalSignatureEx"),
        "unexpected second stream path: {}",
        sigs[1].stream_path
    );
    assert!(
        sigs[2].stream_path.ends_with("\u{0005}DigitalSignature"),
        "unexpected third stream path: {}",
        sigs[2].stream_path
    );

    for sig in &sigs {
        assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    }

    assert_eq!(sigs[0].signature.as_slice(), ext.as_slice());
    assert_eq!(sigs[1].signature.as_slice(), ex.as_slice());
    assert_eq!(sigs[2].signature.as_slice(), legacy.as_slice());
}

#[test]
fn lists_all_signature_streams_and_extracts_digest_info_per_stream() {
    // Use a deterministic digest so we can assert exact output.
    let digest_algorithm_oid = "2.16.840.1.101.3.4.2.1"; // sha256
    let digest = (0u8..32u8).collect::<Vec<u8>>();
    let spc = make_spc_indirect_data_content(digest_algorithm_oid, &digest);

    let valid = make_pkcs7_signed_message(&spc);
    let mut invalid = valid.clone();
    let last = invalid.len().saturating_sub(1);
    invalid[last] ^= 0xFF;

    let vba = build_vba_project_bin_with_signature_streams(&[
        ("\u{0005}DigitalSignature", &invalid),
        ("\u{0005}DigitalSignatureEx", &valid),
    ]);

    let sigs = list_vba_digital_signatures(&vba).expect("signature enumeration should succeed");
    assert_eq!(sigs.len(), 2, "expected two signature streams");

    // Deterministic Excel-like ordering: DigitalSignatureEx is preferred over the legacy
    // DigitalSignature stream.
    assert!(
        sigs[0].stream_path.ends_with("\u{0005}DigitalSignatureEx"),
        "unexpected first stream path: {}",
        sigs[0].stream_path
    );
    assert!(
        sigs[1].stream_path.ends_with("\u{0005}DigitalSignature"),
        "unexpected second stream path: {}",
        sigs[1].stream_path
    );

    assert_eq!(sigs[0].verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sigs[1].verification, VbaSignatureVerification::SignedInvalid);

    assert!(
        sigs[0]
            .signer_subject
            .as_deref()
            .is_some_and(|s| s.contains("Formula VBA Test")),
        "expected signer subject to mention test CN, got: {:?}",
        sigs[0].signer_subject
    );

    for sig in &sigs {
        assert_eq!(
            sig.signed_digest_algorithm_oid.as_deref(),
            Some(digest_algorithm_oid),
            "expected per-stream digest algorithm OID"
        );
        assert_eq!(
            sig.signed_digest.as_deref(),
            Some(digest.as_slice()),
            "expected per-stream digest bytes"
        );
    }
}

#[test]
fn lists_and_verifies_ber_indefinite_signature_stream() {
    let digest_algorithm_oid = "2.16.840.1.101.3.4.2.1"; // sha256
    let digest = (0u8..32u8).collect::<Vec<u8>>();

    // Build a deterministic DER signature with the same digest bytes as the BER fixture so we can
    // assert per-stream extraction results.
    let spc = make_spc_indirect_data_content(digest_algorithm_oid, &digest);
    let der_sig = make_pkcs7_signed_message(&spc);

    // BER-indefinite SignedData fixture (OpenSSL `cms -stream` style).
    let ber_sig = include_bytes!("fixtures/cms_indefinite.der");

    let vba = build_vba_project_bin_with_signature_streams(&[
        ("\u{0005}DigitalSignature", ber_sig),
        ("\u{0005}DigitalSignatureEx", &der_sig),
    ]);

    let sigs = list_vba_digital_signatures(&vba).expect("signature enumeration should succeed");
    assert_eq!(sigs.len(), 2, "expected two signature streams");

    // Deterministic Excel-like ordering: DigitalSignatureEx is preferred over the legacy
    // DigitalSignature stream.
    assert!(sigs[0].stream_path.ends_with("\u{0005}DigitalSignatureEx"));
    assert!(sigs[1].stream_path.ends_with("\u{0005}DigitalSignature"));

    // Both signatures should verify.
    assert_eq!(sigs[0].verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sigs[1].verification, VbaSignatureVerification::SignedVerified);

    for sig in &sigs {
        assert_eq!(
            sig.signed_digest_algorithm_oid.as_deref(),
            Some(digest_algorithm_oid)
        );
        assert_eq!(sig.signed_digest.as_deref(), Some(digest.as_slice()));
    }
}
