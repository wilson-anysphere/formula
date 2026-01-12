use std::io::{Cursor, Write};

use formula_vba::list_vba_digital_signatures;

#[test]
fn exposes_digsig_info_serialized_version_in_signature_enumeration() {
    // Use the existing BER-indefinite CMS fixture (valid PKCS#7 SignedData).
    let pkcs7 = include_bytes!("fixtures/cms_indefinite.der");

    // Synthetic DigSigInfoSerialized-like stream with version prefix:
    // [version, cbSignature, cbSigningCertStore, cchProjectName] (LE u32)
    // [projectName UTF-16LE] [certStore bytes] [signature bytes]
    let version = 3u32;

    let project_name_utf16: Vec<u16> = "VBAProject\0".encode_utf16().collect();
    let mut project_name_bytes = Vec::new();
    for ch in &project_name_utf16 {
        project_name_bytes.extend_from_slice(&ch.to_le_bytes());
    }

    let cert_store = vec![0xAA, 0xBB, 0xCC, 0xDD];

    let cb_signature = pkcs7.len() as u32;
    let cb_cert_store = cert_store.len() as u32;
    let cch_project = project_name_utf16.len() as u32;

    let mut signature_stream = Vec::new();
    signature_stream.extend_from_slice(&version.to_le_bytes());
    signature_stream.extend_from_slice(&cb_signature.to_le_bytes());
    signature_stream.extend_from_slice(&cb_cert_store.to_le_bytes());
    signature_stream.extend_from_slice(&cch_project.to_le_bytes());
    signature_stream.extend_from_slice(&project_name_bytes);
    signature_stream.extend_from_slice(&cert_store);
    signature_stream.extend_from_slice(pkcs7);

    // Store the signature stream in an OLE container (vbaProject.bin).
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut s = ole
            .create_stream("\u{0005}DigitalSignature")
            .expect("create signature stream");
        s.write_all(&signature_stream)
            .expect("write signature stream bytes");
    }
    let vba_project_bin = ole.into_inner().into_inner();

    let sigs = list_vba_digital_signatures(&vba_project_bin).expect("should list signatures");
    assert_eq!(sigs.len(), 1, "expected one signature stream");
    assert_eq!(sigs[0].digsig_info_version, Some(version));

    // Ensure existing pkcs7_offset/pkcs7_len behavior remains unchanged.
    let expected_pkcs7_offset = 16 + project_name_bytes.len() + cert_store.len();
    assert_eq!(sigs[0].pkcs7_offset, Some(expected_pkcs7_offset));
    assert_eq!(sigs[0].pkcs7_len, Some(pkcs7.len()));
}

