use std::io::{Cursor, Write};

use formula_vba::parse_vba_digital_signature;

#[test]
fn parse_prefers_signature_component_even_when_stream_is_nested_under_storage() {
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create compound file");

    // Some producers store the signature as a storage containing a stream rather than a single
    // `\x05DigitalSignature*` stream. Ensure we still detect it, and that we apply Excel/MS-OVBA
    // precedence rules when multiple signature streams exist.
    //
    // `DigitalSignatureEx` is preferred over the legacy `DigitalSignature`.
    ole.create_storage("\u{0005}DigitalSignatureEx")
        .expect("create signature storage");
    {
        let mut stream = ole
            .create_stream("\u{0005}DigitalSignatureEx/sig")
            .expect("create nested signature stream");
        stream
            .write_all(b"nested-signature")
            .expect("write nested signature");
    }

    {
        let mut stream = ole
            .create_stream("\u{0005}DigitalSignature")
            .expect("create signature stream");
        stream
            .write_all(b"other-signature")
            .expect("write other signature");
    }

    let vba_project_bin = ole.into_inner().into_inner();
    let sig = parse_vba_digital_signature(&vba_project_bin)
        .expect("parse should succeed")
        .expect("signature should be present");

    assert!(
        sig.stream_path.contains("/sig"),
        "expected nested DigitalSignatureEx stream to be selected, got {}",
        sig.stream_path
    );
    assert_eq!(sig.signature, b"nested-signature");
}
