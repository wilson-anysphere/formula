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

    // Excel commonly prefixes the stream with a serialized header (e.g. DigSigInfoSerialized)
    // before the DER PKCS#7 blob. Ensure we can still locate the PKCS#7 payload.
    assert_ne!(sig.signature.first(), Some(&0x30));
    assert_eq!(sig.signature.get(4), Some(&0x30));

    let signed_digest = extract_vba_signature_signed_digest(&sig.signature)
        .expect("digest extraction should succeed")
        .expect("digest info should be present");

    match signed_digest.digest_algorithm_oid.as_str() {
        // SHA-1
        "1.3.14.3.2.26" => assert_eq!(signed_digest.digest.len(), 20),
        // SHA-256
        "2.16.840.1.101.3.4.2.1" => assert_eq!(signed_digest.digest.len(), 32),
        other => panic!("unexpected digest algorithm OID: {}", other),
    }
}
