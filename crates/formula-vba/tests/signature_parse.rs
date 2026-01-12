use std::io::{Cursor, Write};

use formula_vba::{parse_vba_digital_signature, VbaSignatureStreamKind};

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

    assert_eq!(
        sig.stream_kind,
        VbaSignatureStreamKind::DigitalSignatureEx,
        "expected nested DigitalSignatureEx stream to be classified as DigitalSignatureEx, got {:?} (path {})",
        sig.stream_kind,
        sig.stream_path
    );
    assert_eq!(
        sig.stream_path,
        "\u{0005}DigitalSignatureEx/sig",
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

    assert_eq!(sig.stream_kind, VbaSignatureStreamKind::DigitalSignatureExt);
    assert!(
        sig.stream_path == "\u{0005}DigitalSignatureExt/sig",
        "expected nested DigitalSignatureExt stream to be selected, got {}",
        sig.stream_path
    );
    assert_eq!(sig.signature, b"nested-ext-signature");
}

#[test]
fn parse_root_digital_signature_ext_sets_stream_kind() {
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create compound file");

    {
        let mut stream = ole
            .create_stream("\u{0005}DigitalSignatureExt")
            .expect("create signature stream");
        stream
            .write_all(b"ext-signature")
            .expect("write signature bytes");
    }

    let vba_project_bin = ole.into_inner().into_inner();
    let sig = parse_vba_digital_signature(&vba_project_bin)
        .expect("parse should succeed")
        .expect("signature should be present");

    assert_eq!(sig.stream_kind, VbaSignatureStreamKind::DigitalSignatureExt);
    assert_eq!(sig.stream_path, "\u{0005}DigitalSignatureExt");
    assert_eq!(sig.signature, b"ext-signature");
}

#[test]
fn parse_signature_like_stream_sets_stream_kind_unknown() {
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create compound file");

    // Some producers may use a signature-like stream/storage name that doesn't exactly match one of
    // the known `DigitalSignature*` variants. We should still detect it as a signature candidate,
    // but surface the kind as `Unknown` so callers don't accidentally apply the wrong binding/digest
    // rules.
    ole.create_storage("\u{0005}DigitalSignatureExWeird")
        .expect("create signature-like storage");
    {
        let mut stream = ole
            .create_stream("\u{0005}DigitalSignatureExWeird/sig")
            .expect("create nested signature stream");
        stream
            .write_all(b"weird-signature")
            .expect("write signature bytes");
    }

    let vba_project_bin = ole.into_inner().into_inner();
    let sig = parse_vba_digital_signature(&vba_project_bin)
        .expect("parse should succeed")
        .expect("signature should be present");

    assert_eq!(sig.stream_kind, VbaSignatureStreamKind::Unknown);
    assert_eq!(sig.stream_path, "\u{0005}DigitalSignatureExWeird/sig");
    assert_eq!(sig.signature, b"weird-signature");
}
