#![cfg(not(target_arch = "wasm32"))]

use formula_vba::{
    extract_vba_signature_signed_digest, list_vba_digital_signatures, verify_vba_digital_signature,
    VbaSignatureBinding, VbaSignatureStreamKind, VbaSignatureVerification,
};

mod signature_test_utils;

use signature_test_utils::{build_vba_project_bin_with_signature, make_pkcs7_signed_message};

fn build_oshared_wordsig_blob(valid_pkcs7: &[u8]) -> Vec<u8> {
    // MS-OSHARED WordSigBlob wraps DigSigInfoSerialized with a UTF-16-length prefix (`cch`).
    //
    // This fixture mirrors `tests/digsig_blob.rs`:
    // - Put a corrupted-but-parseable PKCS#7 blob at the location where pbSignatureBuffer would
    //   typically begin (immediately after the DigSigInfoSerialized header).
    // - Set DigSigInfoSerialized.signatureOffset to point at the real signature later.
    // - Append another corrupted PKCS#7 blob after the real signature so naive scan-last heuristics
    //   would select the wrong blob without WordSigBlob parsing.

    // Corrupt the signature bytes while keeping the overall ASN.1 shape parseable.
    let mut invalid_pkcs7 = valid_pkcs7.to_vec();
    if let Some((idx, _)) = invalid_pkcs7
        .iter()
        .enumerate()
        .rev()
        .find(|&(_i, &b)| b != 0)
    {
        invalid_pkcs7[idx] ^= 0xFF;
    } else if let Some(first) = invalid_pkcs7.get_mut(0) {
        *first ^= 0xFF;
    }

    let base = 2usize; // WordSigBlob offsets are relative to cbSigInfo at offset 2.
    let wordsig_header_len = 10usize; // cch(u16) + cbSigInfo(u32) + serializedPointer(u32)
    let digsig_info_len = 0x24usize; // DigSigInfoSerialized fixed header: 9 DWORDs
    let invalid_offset = wordsig_header_len + digsig_info_len; // 0x2E

    // Place the valid signature after the invalid one and align to 2 bytes (WordSigBlob is a
    // length-prefixed Unicode string).
    let mut signature_offset = invalid_offset + invalid_pkcs7.len();
    signature_offset = (signature_offset + 1) & !1;
    let signature_offset_rel = signature_offset - base;

    let cb_signature = u32::try_from(valid_pkcs7.len()).expect("pkcs7 fits u32");
    let signature_offset_u32 = u32::try_from(signature_offset_rel).expect("offset fits u32");

    let mut out = Vec::new();
    // WordSigBlob.cch placeholder + cbSigInfo placeholder + serializedPointer = 8.
    out.extend_from_slice(&0u16.to_le_bytes());
    out.extend_from_slice(&0u32.to_le_bytes());
    out.extend_from_slice(&8u32.to_le_bytes());

    // DigSigInfoSerialized: only cbSignature and signatureOffset matter for our purposes.
    out.extend_from_slice(&cb_signature.to_le_bytes());
    out.extend_from_slice(&signature_offset_u32.to_le_bytes());
    for _ in 0..7 {
        out.extend_from_slice(&0u32.to_le_bytes());
    }
    assert_eq!(
        out.len(),
        invalid_offset,
        "unexpected DigSigInfoSerialized header size"
    );

    // Decoy PKCS#7 (scan-first heuristics would pick this).
    out.extend_from_slice(&invalid_pkcs7);

    // Pad up to signatureOffset and append the actual signature bytes.
    if out.len() < signature_offset {
        out.resize(signature_offset, 0);
    }
    out.extend_from_slice(valid_pkcs7);

    // Trailing decoy PKCS#7 (scan-last heuristics would pick this).
    out.extend_from_slice(&invalid_pkcs7);

    // WordSigBlob.cbSigInfo: size of the signatureInfo field in bytes (starts at offset 10).
    let signature_info_offset = wordsig_header_len;
    let cb_siginfo = out.len().saturating_sub(signature_info_offset);

    // WordSigBlob.padding: pad the *entire* structure to an even byte length.
    if cb_siginfo % 2 != 0 {
        out.push(0);
    }

    // WordSigBlob.cch: half the byte count of the remainder of the structure.
    let remainder_bytes = out.len().saturating_sub(2);
    assert_eq!(
        remainder_bytes % 2,
        0,
        "expected WordSigBlob remainder to be even"
    );
    let cch = remainder_bytes / 2;

    out[0..2].copy_from_slice(&(cch as u16).to_le_bytes());
    out[2..6].copy_from_slice(&(cb_siginfo as u32).to_le_bytes());

    out
}

#[test]
fn pkcs7_signature_wrapped_in_oshared_wordsig_blob_is_verified() {
    let pkcs7 = make_pkcs7_signed_message(b"formula-vba-test");
    let blob = build_oshared_wordsig_blob(&pkcs7);
    let vba = build_vba_project_bin_with_signature(Some(&blob));

    let sig = verify_vba_digital_signature(&vba)
        .expect("signature inspection should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(
        sig.binding,
        VbaSignatureBinding::Unknown,
        "expected binding to remain unknown without a full MS-OVBA project digest payload"
    );
}

#[test]
fn extracts_signed_digest_from_ber_indefinite_pkcs7_wrapped_in_oshared_wordsig_blob() {
    // This fixture is a CMS/PKCS#7 SignedData blob emitted by OpenSSL with `cms -stream`, which uses
    // BER indefinite-length encodings. Its embedded SpcIndirectDataContent digest is 0..31.
    let pkcs7 = include_bytes!("fixtures/cms_indefinite.der");
    let expected_digest = (0u8..0x20).collect::<Vec<_>>();

    let blob = build_oshared_wordsig_blob(pkcs7);

    let got = extract_vba_signature_signed_digest(&blob)
        .expect("extract should succeed")
        .expect("digest should be present");

    assert_eq!(got.digest_algorithm_oid, "2.16.840.1.101.3.4.2.1"); // SHA-256
    assert_eq!(got.digest, expected_digest);
}

#[test]
fn list_reports_wordsig_blob_pkcs7_offset_and_len() {
    let pkcs7 = make_pkcs7_signed_message(b"formula-vba-test");
    let blob = build_oshared_wordsig_blob(&pkcs7);
    let vba = build_vba_project_bin_with_signature(Some(&blob));

    let sigs = list_vba_digital_signatures(&vba).expect("signature enumeration should succeed");
    assert_eq!(sigs.len(), 1);
    let sig = &sigs[0];

    assert_eq!(sig.stream_kind, VbaSignatureStreamKind::DigitalSignature);

    // This is a synthetic WordSigBlob; the valid PKCS#7 begins after:
    // - WordSigBlob header (10 bytes)
    // - DigSigInfoSerialized fixed header (0x24 bytes)
    // - one decoy PKCS#7 blob (same length as the real one)
    // - optional 1-byte padding to align to 2 bytes
    let mut expected_offset = 0x2E + pkcs7.len();
    expected_offset = (expected_offset + 1) & !1;

    assert_eq!(
        sig.pkcs7_offset,
        Some(expected_offset),
        "expected deterministic pkcs7_offset to be reported"
    );
    assert_eq!(sig.pkcs7_len, Some(pkcs7.len()));
}
