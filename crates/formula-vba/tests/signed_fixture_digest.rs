use std::io::Read;

use formula_vba::{extract_vba_signature_signed_digest, parse_vba_digital_signature};

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

    // Many real-world files wrap the PKCS#7 blob in a [MS-OFFCRYPTO] DigSigInfoSerialized header.
    assert_ne!(sig.signature.first(), Some(&0x30));
    assert!(
        sig.signature.len() >= 12,
        "expected at least DigSigInfoSerialized header"
    );

    let cb_signature = u32::from_le_bytes(sig.signature[0..4].try_into().unwrap()) as usize;
    let cb_cert_store = u32::from_le_bytes(sig.signature[4..8].try_into().unwrap()) as usize;
    let cch_project_name = u32::from_le_bytes(sig.signature[8..12].try_into().unwrap()) as usize;
    let project_name_bytes = cch_project_name * 2;

    let cert_store_offset = 12 + project_name_bytes;
    let pkcs7_offset = cert_store_offset + cb_cert_store;

    // The fixture intentionally includes a *decoy* PKCS#7 blob inside the certificate store bytes
    // so that naive scanning would pick the wrong payload. Correct handling should use the
    // DigSigInfoSerialized length fields to locate the real signature.
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
}
