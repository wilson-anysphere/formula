use std::io::{Cursor, Write};

use formula_vba::parse_vba_digital_signature;

#[test]
fn parse_prefers_signature_component_even_when_stream_is_nested_under_storage() {
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create compound file");

    // Some producers store the signature as a storage containing a stream rather than a single
    // `\x05DigitalSignature*` stream. Ensure we still detect it, and that we apply Excel-like
    // stream-name precedence rules when multiple signature streams exist.
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

#[test]
fn parse_prefers_digital_signature_ext_over_ex() {
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create compound file");

    // Some producers store the signature as a storage containing a stream rather than a single
    // `\x05DigitalSignature*` stream. Ensure we still detect it, and that we apply Excel-like
    // stream-name precedence rules when multiple signature streams exist.
    //
    // `DigitalSignatureExt` is preferred over `DigitalSignatureEx`.
    ole.create_storage("\u{0005}DigitalSignatureExt")
        .expect("create signature storage");
    {
        let mut stream = ole
            .create_stream("\u{0005}DigitalSignatureExt/sig")
            .expect("create nested signature stream");
        stream
            .write_all(b"nested-ext-signature")
            .expect("write nested signature");
    }

    {
        let mut stream = ole
            .create_stream("\u{0005}DigitalSignatureEx")
            .expect("create signature stream");
        stream
            .write_all(b"ex-signature")
            .expect("write other signature");
    }

    let vba_project_bin = ole.into_inner().into_inner();
    let sig = parse_vba_digital_signature(&vba_project_bin)
        .expect("parse should succeed")
        .expect("signature should be present");

    assert!(
        sig.stream_path.contains("DigitalSignatureExt") && sig.stream_path.contains("/sig"),
        "expected nested DigitalSignatureExt stream to be selected, got {}",
        sig.stream_path
    );
    assert_eq!(sig.signature, b"nested-ext-signature");
}
