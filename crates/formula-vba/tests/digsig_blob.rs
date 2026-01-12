#![cfg(not(target_arch = "wasm32"))]

use formula_vba::{
    list_vba_digital_signatures, parse_vba_digital_signature, verify_vba_digital_signature,
    VbaSignatureBinding, VbaSignatureVerification,
};

mod signature_test_utils;

use signature_test_utils::{build_vba_project_bin_with_signature, make_pkcs7_signed_message};

fn build_oshared_digsig_blob(valid_pkcs7: &[u8]) -> Vec<u8> {
    // MS-OSHARED describes a DigSigBlob wrapper around the PKCS#7 signature bytes.
    //
    // In this synthetic blob:
    // - We place a *corrupted but still parseable* PKCS#7 blob immediately after the
    //   DigSigInfoSerialized header (at offset 0x2C). A heuristic scan that picks the first
    //   SignedData blob would lock onto this and report `SignedInvalid`.
    // - DigSigInfoSerialized.signatureOffset points at the *valid* PKCS#7 blob later in the stream.
    // - We append another corrupted PKCS#7 blob *after* the real signature. The verifier's current
    //   heuristic scan prefers the last SignedData blob in the stream, so without DigSigBlob parsing
    //   it would pick the trailing corrupt blob and report `SignedInvalid`.
    //
    // This ensures we exercise deterministic DigSigBlob parsing (offset-based) instead of relying
    // on the heuristic scan-for-0x30 fallback.
    // Corrupt the signature bytes while keeping the overall ASN.1 shape parseable.
    //
    // For BER-indefinite encodings (e.g. OpenSSL `cms -stream` output) the PKCS#7 blob can end with
    // multiple `0x00 0x00` EOC terminators. Flipping the *last* byte would break the EOC and cause
    // the blob to become unparseable (defeating the purpose of this test, which is to ensure our
    // locator prefers the DigSigBlob offset rather than heuristic scanning).
    //
    // Instead, flip the last non-zero byte, which is typically inside the signature value.
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

    let digsig_blob_header_len = 8usize; // cb + serializedPointer
    // DigSigInfoSerialized is 9 DWORDs total in MS-OSHARED:
    // cbSignature, signatureOffset, cbSigningCertStore, certStoreOffset, cbProjectName,
    // projectNameOffset, fTimestamp, cbTimestampUrl, timestampUrlOffset.
    let digsig_info_len = 0x24usize;

    let invalid_offset = digsig_blob_header_len + digsig_info_len; // 0x2C (matches MS-OSHARED examples)
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
    // Remaining fields (cert store/project name/timestamp URL) set to 0.
    for _ in 0..7 {
        out.extend_from_slice(&0u32.to_le_bytes());
    }
    assert_eq!(
        out.len(),
        invalid_offset,
        "unexpected DigSigInfoSerialized size"
    );

    // Insert a corrupted PKCS#7 blob early in the stream to ensure scan-first heuristics fail.
    out.extend_from_slice(&invalid_pkcs7);

    // Pad up to signatureOffset and append the actual signature bytes.
    if out.len() < signature_offset {
        out.resize(signature_offset, 0);
    }
    out.extend_from_slice(valid_pkcs7);

    // Append an invalid PKCS#7 blob after the real signature to ensure heuristic scanning would
    // pick the wrong SignedData candidate.
    out.extend_from_slice(&invalid_pkcs7);

    // DigSigBlob.cb: size of the serialized signatureInfo payload (excluding the initial DWORDs).
    let cb =
        u32::try_from(out.len().saturating_sub(digsig_blob_header_len)).expect("blob fits u32");
    out[0..4].copy_from_slice(&cb.to_le_bytes());

    out
}

#[test]
fn pkcs7_signature_wrapped_in_oshared_digsig_blob_is_verified() {
    let pkcs7 = make_pkcs7_signed_message(b"formula-vba-test");
    let blob = build_oshared_digsig_blob(&pkcs7);
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
    assert!(
        sig.signer_subject
            .as_deref()
            .is_some_and(|s| s.contains("Formula VBA Test")),
        "expected signer subject to mention test CN, got: {:?}",
        sig.signer_subject
    );
}

#[test]
fn ber_indefinite_pkcs7_signature_wrapped_in_oshared_digsig_blob_is_verified() {
    // This fixture is a CMS/PKCS#7 SignedData blob emitted by OpenSSL with `cms -stream`, which
    // uses BER indefinite-length encodings.
    let pkcs7 = include_bytes!("fixtures/cms_indefinite.der");
    let blob = build_oshared_digsig_blob(pkcs7);
    let vba = build_vba_project_bin_with_signature(Some(&blob));

    let sig = verify_vba_digital_signature(&vba)
        .expect("signature inspection should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
}

#[test]
fn parse_returns_digsig_blob_bytes_intact() {
    let pkcs7 = make_pkcs7_signed_message(b"formula-vba-test");
    let blob = build_oshared_digsig_blob(&pkcs7);
    let vba = build_vba_project_bin_with_signature(Some(&blob));

    let sig = parse_vba_digital_signature(&vba)
        .expect("parse should succeed")
        .expect("signature should be present");

    assert_eq!(
        sig.signature, blob,
        "expected raw stream bytes to be preserved"
    );
}

#[test]
fn list_reports_digsig_blob_pkcs7_offset_and_len() {
    let pkcs7 = make_pkcs7_signed_message(b"formula-vba-test");
    let blob = build_oshared_digsig_blob(&pkcs7);
    let vba = build_vba_project_bin_with_signature(Some(&blob));

    let sigs = list_vba_digital_signatures(&vba).expect("signature enumeration should succeed");
    assert_eq!(sigs.len(), 1, "expected one signature stream");
    let sig = &sigs[0];

    let invalid_offset = 0x2Cusize;
    let mut expected_offset = invalid_offset + pkcs7.len();
    expected_offset = (expected_offset + 3) & !3;

    assert_eq!(sig.pkcs7_offset, Some(expected_offset));
    assert_eq!(sig.pkcs7_len, Some(pkcs7.len()));
}
