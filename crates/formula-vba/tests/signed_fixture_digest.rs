use std::io::Read;

use formula_vba::{
    extract_signer_certificate_info, extract_vba_signature_signed_digest, list_vba_digital_signatures,
    parse_vba_digital_signature, verify_vba_digital_signature, VbaSignatureBinding,
    VbaSignatureVerification,
};

fn load_fixture_vba_bin() -> Vec<u8> {
    let fixture_path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/xlsx/macros/signed-basic.xlsm"
    );
    let data = std::fs::read(fixture_path).expect("fixture xlsm exists");
    let reader = std::io::Cursor::new(data);
    let mut zip = zip::ZipArchive::new(reader).expect("valid zip");
    let mut file = zip
        .by_name("xl/vbaProject.bin")
        .expect("vbaProject.bin in fixture");
    let mut buf = Vec::new();
    file.read_to_end(&mut buf).unwrap();
    buf
}

#[test]
fn extracts_spc_indirect_data_digest_from_signed_vba_fixture() {
    let vba_bin = load_fixture_vba_bin();
    let sig = parse_vba_digital_signature(&vba_bin)
        .expect("signature parse should succeed")
        .expect("signature should be present");

    assert!(
        sig.stream_path.contains("DigitalSignature"),
        "expected DigitalSignature stream, got {}",
        sig.stream_path
    );
    assert!(
        sig.signer_subject
            .as_deref()
            .is_some_and(|s| s.contains("Formula VBA Fixture")),
        "expected signer subject to mention fixture CN, got: {:?}",
        sig.signer_subject
    );

    // Many real-world files wrap the PKCS#7 blob in a length-prefixed DigSigInfoSerialized-like
    // header (distinct from the MS-OSHARED DigSigBlob/offset-based wrapper).
    assert_ne!(sig.signature.first(), Some(&0x30));
    assert!(
        sig.signature.len() >= 12,
        "expected at least DigSigInfoSerialized-like header"
    );

    let cb_signature = u32::from_le_bytes(sig.signature[0..4].try_into().unwrap()) as usize;
    let cb_cert_store = u32::from_le_bytes(sig.signature[4..8].try_into().unwrap()) as usize;
    let cch_project_name = u32::from_le_bytes(sig.signature[8..12].try_into().unwrap()) as usize;
    let project_name_bytes = cch_project_name * 2;

    let cert_store_offset = 12 + project_name_bytes;
    let pkcs7_offset = cert_store_offset + cb_cert_store;

    // The fixture intentionally includes a *decoy* PKCS#7 blob inside the certificate store bytes
    // so that naive scanning would pick the wrong payload. Correct handling should use the
    // DigSigInfoSerialized-like length fields to locate the real signature.
    assert_eq!(sig.signature.get(cert_store_offset), Some(&0x30));
    assert_eq!(sig.signature.get(pkcs7_offset), Some(&0x30));
    assert_eq!(cb_signature, sig.signature.len().saturating_sub(pkcs7_offset));

    let signed_digest = extract_vba_signature_signed_digest(&sig.signature)
        .expect("digest extraction should succeed")
        .expect("digest info should be present");

    assert_eq!(
        signed_digest.digest_algorithm_oid,
        "2.16.840.1.101.3.4.2.1"
    );
    assert_eq!(signed_digest.digest, (0u8..32).collect::<Vec<_>>());

    let cert_info =
        extract_signer_certificate_info(&sig.signature).expect("expected embedded certificate info");
    assert!(
        cert_info.subject.contains("Formula VBA Fixture"),
        "expected certificate subject to mention fixture CN, got: {}",
        cert_info.subject
    );
    assert!(
        !cert_info.sha256_fingerprint_hex.is_empty(),
        "expected certificate fingerprint"
    );
}

#[test]
fn extracts_digest_even_when_digsig_header_is_corrupt() {
    let vba_bin = load_fixture_vba_bin();
    let sig = parse_vba_digital_signature(&vba_bin)
        .expect("signature parse should succeed")
        .expect("signature should be present");

    // Corrupt the length-prefixed DigSigInfoSerialized-like header length fields so the
    // deterministic unwrapping logic can't apply, forcing the extractor to fall back to scanning for
    // an embedded CMS SignedData.
    let mut corrupted = sig.signature.clone();
    corrupted[0..4].copy_from_slice(&0xFFFF_FFFFu32.to_le_bytes());

    let signed_digest = extract_vba_signature_signed_digest(&corrupted)
        .expect("digest extraction should succeed")
        .expect("digest info should be present");

    assert_eq!(
        signed_digest.digest_algorithm_oid,
        "2.16.840.1.101.3.4.2.1"
    );
    assert_eq!(signed_digest.digest, (0u8..32).collect::<Vec<_>>());
}

#[test]
#[cfg(not(target_arch = "wasm32"))]
fn verifies_signature_even_when_digsig_header_is_corrupt() {
    use std::io::{Cursor, Write};

    let vba_bin = load_fixture_vba_bin();
    let sig = parse_vba_digital_signature(&vba_bin)
        .expect("signature parse should succeed")
        .expect("signature should be present");

    // Corrupt DigSigInfoSerialized-like sizes so we exercise the fallback scanning logic in PKCS#7
    // verification.
    let mut corrupted = sig.signature.clone();
    corrupted[0..4].copy_from_slice(&0xFFFF_FFFFu32.to_le_bytes());

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole
            .create_stream("\u{0005}DigitalSignature")
            .expect("create signature stream");
        stream.write_all(&corrupted).expect("write signature");
    }
    let vba_project_bin = ole.into_inner().into_inner();

    let verified = verify_vba_digital_signature(&vba_project_bin)
        .expect("verification should succeed")
        .expect("signature should be present");
    assert_eq!(verified.verification, VbaSignatureVerification::SignedVerified);
}

#[test]
fn verifies_signed_vba_fixture_signature_and_reports_binding_mismatch() {
    let vba_bin = load_fixture_vba_bin();
    let sig = verify_vba_digital_signature(&vba_bin)
        .expect("signature verification should succeed")
        .expect("signature should be present");

    #[cfg(not(target_arch = "wasm32"))]
    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    #[cfg(target_arch = "wasm32")]
    assert_eq!(sig.verification, VbaSignatureVerification::SignedButUnverified);

    // The fixture embeds a synthetic digest value; it is not intended to match the computed digest
    // of the fixture's project streams.
    assert_eq!(sig.binding, VbaSignatureBinding::NotBound);
}

#[test]
fn lists_signature_stream_and_reports_digsig_pkcs7_location() {
    let vba_bin = load_fixture_vba_bin();
    let sigs = list_vba_digital_signatures(&vba_bin).expect("signature enumeration should succeed");
    assert_eq!(sigs.len(), 1);

    let sig = &sigs[0];
    assert!(
        sig.stream_path.contains("DigitalSignature"),
        "expected DigitalSignature stream, got {}",
        sig.stream_path
    );

    let signature_stream = &sig.signature;
    assert_ne!(signature_stream.first(), Some(&0x30));
    assert!(signature_stream.len() >= 12);

    let cb_signature = u32::from_le_bytes(signature_stream[0..4].try_into().unwrap()) as usize;
    let cb_cert_store = u32::from_le_bytes(signature_stream[4..8].try_into().unwrap()) as usize;
    let cch_project_name =
        u32::from_le_bytes(signature_stream[8..12].try_into().unwrap()) as usize;
    let project_name_bytes = cch_project_name * 2;

    let expected_pkcs7_offset = 12 + project_name_bytes + cb_cert_store;

    assert_eq!(
        sig.pkcs7_offset,
        Some(expected_pkcs7_offset),
        "expected DigSigInfoSerialized pkcs7_offset"
    );
    assert_eq!(
        sig.pkcs7_len,
        Some(cb_signature),
        "expected DigSigInfoSerialized pkcs7_len"
    );

    assert_eq!(
        sig.signed_digest_algorithm_oid.as_deref(),
        Some("2.16.840.1.101.3.4.2.1")
    );
    assert_eq!(
        sig.signed_digest.as_deref(),
        Some((0u8..32).collect::<Vec<_>>().as_slice())
    );

    #[cfg(not(target_arch = "wasm32"))]
    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    #[cfg(target_arch = "wasm32")]
    assert_eq!(sig.verification, VbaSignatureVerification::SignedButUnverified);
}
