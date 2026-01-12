#![cfg(all(feature = "vba", not(target_arch = "wasm32")))]

use std::io::{Cursor, Write};

use formula_vba::VbaSignatureVerification;
use formula_xlsx::XlsxPackage;

mod vba_signature_test_utils;
use vba_signature_test_utils::{build_vba_signature_ole, make_pkcs7_signed_message};

fn build_package(entries: &[(&str, &[u8])]) -> Vec<u8> {
    use zip::write::FileOptions;
    use zip::ZipWriter;

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    for (name, bytes) in entries {
        zip.start_file(*name, options).unwrap();
        zip.write_all(bytes).unwrap();
    }

    zip.finish().unwrap().into_inner()
}

#[test]
fn verify_vba_signature_is_resolved_via_vba_project_relationships() {
    let pkcs7 = make_pkcs7_signed_message(b"formula-xlsx-test");
    let sig_container = build_vba_signature_ole(&pkcs7);

    let rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature" Target="vba/customSig.bin"/>
</Relationships>"#;

    let bytes = build_package(&[
        // `verify_vba_digital_signature` should return `Ok(None)` when this is missing, so we
        // include it to exercise relationship-based signature resolution.
        ("xl/vbaProject.bin", b"fake-vba-project"),
        ("xl/_rels/vbaProject.bin.rels", rels.as_bytes()),
        ("xl/vba/customSig.bin", &sig_container),
    ]);

    let pkg = XlsxPackage::from_bytes(&bytes).expect("read package");
    let sig = pkg
        .verify_vba_digital_signature()
        .expect("verify signature")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
}
