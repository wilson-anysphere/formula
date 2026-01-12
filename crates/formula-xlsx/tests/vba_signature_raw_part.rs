#![cfg(all(feature = "vba", not(target_arch = "wasm32")))]

use std::io::{Cursor, Write};

use formula_vba::{
    compress_container, compute_vba_project_digest_v3, content_normalized_data, DigestAlg,
    VbaProjectBindingVerification, VbaSignatureBinding, VbaSignatureVerification,
};
use formula_xlsx::XlsxPackage;
use openssl::hash::{hash, MessageDigest};
use zip::write::FileOptions;

mod vba_signature_test_utils;
use vba_signature_test_utils::make_pkcs7_signed_message;

fn push_record(out: &mut Vec<u8>, id: u16, data: &[u8]) {
    out.extend_from_slice(&id.to_le_bytes());
    out.extend_from_slice(&(data.len() as u32).to_le_bytes());
    out.extend_from_slice(data);
}

fn build_minimal_vba_project_bin(module1: &[u8]) -> Vec<u8> {
    // `content_normalized_data` expects a decompressed-and-parsable `VBA/dir` stream and module
    // streams containing MS-OVBA compressed containers.
    let module_container = compress_container(module1);

    let dir_decompressed = {
        let mut out = Vec::new();
        // Minimal module record group.
        push_record(&mut out, 0x0019, b"Module1"); // MODULENAME

        // MODULESTREAMNAME + reserved u16.
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"Module1");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);

        // MODULETYPE (standard)
        push_record(&mut out, 0x0021, &0u16.to_le_bytes());
        // MODULETEXTOFFSET: our module stream is just the compressed container.
        push_record(&mut out, 0x0031, &0u32.to_le_bytes());
        out
    };
    let dir_container = compress_container(&dir_decompressed);

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create compound file");

    {
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(b"Name=\"VBAProject\"\r\nModule=Module1\r\n")
            .expect("write PROJECT");
    }

    ole.create_storage("VBA").expect("VBA storage");

    {
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }

    {
        let mut s = ole.create_stream("VBA/Module1").expect("module stream");
        s.write_all(&module_container).expect("write module");
    }

    ole.into_inner().into_inner()
}

fn build_minimal_vba_project_bin_v3(designer_payload: &[u8]) -> Vec<u8> {
    let module_source = b"Sub Hello()\r\nEnd Sub\r\n";
    let module_container = compress_container(module_source);
    let userform_source = b"Sub FormHello()\r\nEnd Sub\r\n";
    let userform_container = compress_container(userform_source);

    // Minimal `dir` stream (decompressed form) with:
    // - one standard module, and
    // - one UserForm module so FormsNormalizedData is non-empty.
    let dir_decompressed = {
        let mut out = Vec::new();
        // Include a v3-specific reference record type so the transcript depends on it.
        let libid_twiddled = b"REFCTRL-V3";
        let reserved1: u32 = 0;
        let reserved2: u16 = 0;
        let mut reference_control = Vec::new();
        reference_control.extend_from_slice(&(libid_twiddled.len() as u32).to_le_bytes());
        reference_control.extend_from_slice(libid_twiddled);
        reference_control.extend_from_slice(&reserved1.to_le_bytes());
        reference_control.extend_from_slice(&reserved2.to_le_bytes());
        push_record(&mut out, 0x002F, &reference_control);

        // MODULENAME (standard module)
        push_record(&mut out, 0x0019, b"Module1");
        // MODULESTREAMNAME + reserved u16
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"Module1");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);
        // MODULETYPE (standard)
        push_record(&mut out, 0x0021, &0u16.to_le_bytes());
        // MODULETEXTOFFSET
        push_record(&mut out, 0x0031, &0u32.to_le_bytes());

        // MODULENAME (UserForm/designer module referenced from PROJECT by BaseClass=)
        push_record(&mut out, 0x0019, b"UserForm1");
        // MODULESTREAMNAME + reserved u16
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"UserForm1");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);
        // MODULETYPE = UserForm (0x0003 per MS-OVBA).
        push_record(&mut out, 0x0021, &0x0003u16.to_le_bytes());
        // MODULETEXTOFFSET
        push_record(&mut out, 0x0031, &0u32.to_le_bytes());

        out
    };
    let dir_container = compress_container(&dir_decompressed);

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    ole.create_storage("VBA").expect("VBA storage");
    ole.create_storage("UserForm1").expect("designer storage");

    {
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(b"Name=\"VBAProject\"\r\nModule=Module1\r\nBaseClass=\"UserForm1\"\r\n")
            .expect("write PROJECT");
    }

    {
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }
    {
        let mut s = ole.create_stream("VBA/Module1").expect("module stream");
        s.write_all(&module_container).expect("write module");
    }
    {
        let mut s = ole
            .create_stream("VBA/UserForm1")
            .expect("userform module stream");
        s.write_all(&userform_container)
            .expect("write userform module");
    }

    // Designer payload so FormsNormalizedData is non-empty (and therefore bound by v3 digest).
    {
        let mut s = ole
            .create_stream("UserForm1/Payload")
            .expect("designer stream");
        s.write_all(designer_payload)
            .expect("write designer payload");
    }

    ole.into_inner().into_inner()
}

fn der_len(len: usize) -> Vec<u8> {
    if len < 0x80 {
        return vec![len as u8];
    }
    let mut bytes = Vec::new();
    let mut tmp = len;
    while tmp > 0 {
        bytes.push((tmp & 0xFF) as u8);
        tmp >>= 8;
    }
    bytes.reverse();
    let mut out = Vec::with_capacity(1 + bytes.len());
    out.push(0x80 | (bytes.len() as u8));
    out.extend_from_slice(&bytes);
    out
}

fn der_tlv(tag: u8, value: &[u8]) -> Vec<u8> {
    let mut out = Vec::new();
    out.push(tag);
    out.extend_from_slice(&der_len(value.len()));
    out.extend_from_slice(value);
    out
}

fn der_sequence(children: &[Vec<u8>]) -> Vec<u8> {
    let mut value = Vec::new();
    for child in children {
        value.extend_from_slice(child);
    }
    der_tlv(0x30, &value)
}

fn der_octet_string(bytes: &[u8]) -> Vec<u8> {
    der_tlv(0x04, bytes)
}

fn der_null() -> Vec<u8> {
    vec![0x05, 0x00]
}

fn der_oid(oid: &str) -> Vec<u8> {
    let arcs: Vec<u32> = oid
        .split('.')
        .map(|s| s.parse::<u32>().expect("numeric arc"))
        .collect();
    assert!(arcs.len() >= 2, "OID needs at least 2 arcs");
    let mut out = Vec::new();
    out.push((arcs[0] * 40 + arcs[1]) as u8);
    for &arc in &arcs[2..] {
        let mut tmp = arc;
        let mut buf = Vec::new();
        buf.push((tmp & 0x7F) as u8);
        tmp >>= 7;
        while tmp > 0 {
            buf.push(((tmp & 0x7F) as u8) | 0x80);
            tmp >>= 7;
        }
        buf.reverse();
        out.extend_from_slice(&buf);
    }
    der_tlv(0x06, &out)
}

fn make_spc_indirect_data_content_sha256(digest: &[u8]) -> Vec<u8> {
    // data SpcAttributeTypeAndOptionalValue ::= SEQUENCE { type OBJECT IDENTIFIER, value [0] EXPLICIT ANY OPTIONAL }
    let data = der_sequence(&[der_oid("1.3.6.1.4.1.311.2.1.15")]);

    // messageDigest DigestInfo ::= SEQUENCE { digestAlgorithm AlgorithmIdentifier, digest OCTET STRING }
    let alg = der_sequence(&[der_oid("2.16.840.1.101.3.4.2.1"), der_null()]);
    let digest_info = der_sequence(&[alg, der_octet_string(digest)]);

    der_sequence(&[data, digest_info])
}

fn build_oshared_digsig_blob(valid_pkcs7: &[u8]) -> Vec<u8> {
    // MS-OSHARED describes a DigSigBlob wrapper around the PKCS#7 signature bytes.
    //
    // This test blob intentionally contains *multiple* PKCS#7 SignedData blobs:
    // - an invalid (but parseable) one early, and
    // - an invalid one after the real signature.
    //
    // This ensures verification succeeds only when the DigSigBlob offsets are honored, not when
    // relying on heuristic scan-for-0x30 fallbacks.
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
    let digsig_info_len = 0x24usize; // 9 DWORDs (cbSignature/signatureOffset + 7 reserved)
    let invalid_offset = digsig_blob_header_len + digsig_info_len; // 0x2C
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
    // Remaining fields set to 0 (cert store/project name/timestamp URL).
    for _ in 0..7 {
        out.extend_from_slice(&0u32.to_le_bytes());
    }
    assert_eq!(out.len(), invalid_offset);

    // Insert a corrupted PKCS#7 blob early in the stream (to break scan-first heuristics).
    out.extend_from_slice(&invalid_pkcs7);

    // Pad up to signatureOffset and append the actual signature bytes.
    if out.len() < signature_offset {
        out.resize(signature_offset, 0);
    }
    out.extend_from_slice(valid_pkcs7);

    // Append an invalid PKCS#7 blob after the real signature (to break scan-last heuristics).
    out.extend_from_slice(&invalid_pkcs7);

    // DigSigBlob.cb: size of the serialized signatureInfo payload (excluding the initial DWORDs).
    let cb =
        u32::try_from(out.len().saturating_sub(digsig_blob_header_len)).expect("blob fits u32");
    out[0..4].copy_from_slice(&cb.to_le_bytes());

    out
}

fn build_oshared_wordsig_blob(valid_pkcs7: &[u8]) -> Vec<u8> {
    // MS-OSHARED WordSigBlob wraps DigSigInfoSerialized with a UTF-16-length prefix (`cch`).
    //
    // This fixture mirrors `build_oshared_digsig_blob`:
    // - Put a corrupted-but-parseable PKCS#7 blob at the location where pbSignatureBuffer would
    //   typically begin.
    // - Set DigSigInfoSerialized.signatureOffset to point at the real signature later.
    // - Append another corrupted PKCS#7 blob after the real signature so naive scan-last heuristics
    //   would select the wrong blob without WordSigBlob parsing.

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

fn build_digsig_info_serialized_wrapper(pkcs7: &[u8]) -> Vec<u8> {
    // Some real-world `\x05DigitalSignature*` streams are prefixed by a length-prefixed
    // DigSigInfoSerialized-like header rather than by a full MS-OSHARED DigSigBlob offset table.
    //
    // Layout (one common variant):
    //   [cbSignature, cbSigningCertStore, cchProjectName] (LE u32)
    //   [projectName UTF-16LE] [certStore bytes] [signature bytes]
    let project_name_utf16: Vec<u16> = "VBAProject\0".encode_utf16().collect();
    let mut project_name_bytes = Vec::new();
    for ch in &project_name_utf16 {
        project_name_bytes.extend_from_slice(&ch.to_le_bytes());
    }
    let cert_store = vec![0xAA, 0xBB, 0xCC, 0xDD];

    let cb_signature = u32::try_from(pkcs7.len()).expect("pkcs7 fits u32");
    let cb_cert_store = u32::try_from(cert_store.len()).expect("cert store fits u32");
    let cch_project = u32::try_from(project_name_utf16.len()).expect("project name fits u32");

    let mut out = Vec::new();
    out.extend_from_slice(&cb_signature.to_le_bytes());
    out.extend_from_slice(&cb_cert_store.to_le_bytes());
    out.extend_from_slice(&cch_project.to_le_bytes());
    out.extend_from_slice(&project_name_bytes);
    out.extend_from_slice(&cert_store);
    out.extend_from_slice(pkcs7);
    out
}

#[test]
fn verifies_raw_vba_project_signature_part_when_not_ole() {
    let pkcs7 = make_pkcs7_signed_message(b"formula-vba-test");

    // `xl/vbaProject.bin` must be a valid OLE file (even if unsigned) so the
    // fallback embedded-signature scan can run without errors.
    let vba_project_bin = {
        let cursor = Cursor::new(Vec::new());
        let ole = cfb::CompoundFile::create(cursor).expect("create compound file");
        ole.into_inner().into_inner()
    };

    let vba_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature" Target="vbaProjectSignature.bin"/>
</Relationships>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("xl/vbaProject.bin", options).unwrap();
    zip.write_all(&vba_project_bin).unwrap();

    zip.start_file("xl/_rels/vbaProject.bin.rels", options)
        .unwrap();
    zip.write_all(vba_rels).unwrap();

    zip.start_file("xl/vbaProjectSignature.bin", options)
        .unwrap();
    zip.write_all(&pkcs7).unwrap();

    let bytes = zip.finish().unwrap().into_inner();
    let pkg = XlsxPackage::from_bytes(&bytes).expect("read package");

    let sig = pkg
        .verify_vba_digital_signature()
        .expect("signature verification should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert!(
        sig.signer_subject
            .as_deref()
            .is_some_and(|s| s.contains("Formula VBA Test")),
        "expected signer subject to mention test CN, got: {:?}",
        sig.signer_subject
    );
}

#[test]
fn verifies_raw_signature_part_binding_against_vba_project_bin() {
    let module1 = b"Sub Hello()\r\nEnd Sub\r\n";
    let vba_project_bin = build_minimal_vba_project_bin(module1);

    // Signed digest is MD5(ContentNormalizedData) per MS-OVBA.
    let normalized = content_normalized_data(&vba_project_bin).expect("content normalized data");
    let digest = hash(MessageDigest::md5(), &normalized)
        .expect("md5 digest")
        .to_vec();

    // Authenticode SpcIndirectDataContent: DigestInfo.algorithm is typically SHA-256 in practice,
    // but DigestInfo.digest bytes are still the 16-byte MD5 project digest for VBA signatures.
    let spc = make_spc_indirect_data_content_sha256(&digest);
    let pkcs7 = make_pkcs7_signed_message(&spc);

    let vba_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature" Target="vbaProjectSignature.bin"/>
</Relationships>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("xl/vbaProject.bin", options).unwrap();
    zip.write_all(&vba_project_bin).unwrap();

    zip.start_file("xl/_rels/vbaProject.bin.rels", options)
        .unwrap();
    zip.write_all(vba_rels).unwrap();

    // Raw PKCS#7 blob: not an OLE container.
    zip.start_file("xl/vbaProjectSignature.bin", options)
        .unwrap();
    zip.write_all(&pkcs7).unwrap();

    let bytes = zip.finish().unwrap().into_inner();
    let pkg = XlsxPackage::from_bytes(&bytes).expect("read package");

    let sig = pkg
        .verify_vba_digital_signature()
        .expect("signature verification should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.binding, VbaSignatureBinding::Bound);

    // Tamper with a covered project stream but keep the signature bytes the same.
    let mut tampered_module = module1.to_vec();
    tampered_module[0] ^= 0xFF;
    let tampered_project = build_minimal_vba_project_bin(&tampered_module);

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("xl/vbaProject.bin", options).unwrap();
    zip.write_all(&tampered_project).unwrap();

    zip.start_file("xl/_rels/vbaProject.bin.rels", options)
        .unwrap();
    zip.write_all(vba_rels).unwrap();

    zip.start_file("xl/vbaProjectSignature.bin", options)
        .unwrap();
    zip.write_all(&pkcs7).unwrap();

    let bytes = zip.finish().unwrap().into_inner();
    let pkg = XlsxPackage::from_bytes(&bytes).expect("read tampered package");

    let sig = pkg
        .verify_vba_digital_signature()
        .expect("signature verification should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.binding, VbaSignatureBinding::NotBound);
}

#[test]
fn verifies_raw_signature_part_binding_when_digsig_blob_wrapped() {
    let module1 = b"Sub Hello()\r\nEnd Sub\r\n";
    let vba_project_bin = build_minimal_vba_project_bin(module1);

    // Signed digest is MD5(ContentNormalizedData) per MS-OVBA.
    let normalized = content_normalized_data(&vba_project_bin).expect("content normalized data");
    let digest = hash(MessageDigest::md5(), &normalized)
        .expect("md5 digest")
        .to_vec();

    let spc = make_spc_indirect_data_content_sha256(&digest);
    let pkcs7 = make_pkcs7_signed_message(&spc);
    let wrapped = build_oshared_digsig_blob(&pkcs7);

    let vba_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature" Target="vbaProjectSignature.bin"/>
</Relationships>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("xl/vbaProject.bin", options).unwrap();
    zip.write_all(&vba_project_bin).unwrap();

    zip.start_file("xl/_rels/vbaProject.bin.rels", options).unwrap();
    zip.write_all(vba_rels).unwrap();

    // Not an OLE container, not raw PKCS#7: DigSigBlob-wrapped PKCS#7 payload.
    zip.start_file("xl/vbaProjectSignature.bin", options).unwrap();
    zip.write_all(&wrapped).unwrap();

    let bytes = zip.finish().unwrap().into_inner();
    let pkg = XlsxPackage::from_bytes(&bytes).expect("read package");

    let sig = pkg
        .verify_vba_digital_signature()
        .expect("signature verification should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.binding, VbaSignatureBinding::Bound);
    assert_eq!(sig.stream_path, "xl/vbaProjectSignature.bin");

    let binding = pkg
        .vba_project_signature_binding()
        .expect("binding verification")
        .expect("project should be present");
    assert!(
        matches!(binding, VbaProjectBindingVerification::BoundVerified(_)),
        "expected BoundVerified, got {binding:?}"
    );
}

#[test]
fn verifies_raw_signature_part_binding_when_wordsig_blob_wrapped() {
    let module1 = b"Sub Hello()\r\nEnd Sub\r\n";
    let vba_project_bin = build_minimal_vba_project_bin(module1);

    // Signed digest is MD5(ContentNormalizedData) per MS-OVBA.
    let normalized = content_normalized_data(&vba_project_bin).expect("content normalized data");
    let digest = hash(MessageDigest::md5(), &normalized)
        .expect("md5 digest")
        .to_vec();

    let spc = make_spc_indirect_data_content_sha256(&digest);
    let pkcs7 = make_pkcs7_signed_message(&spc);
    let wrapped = build_oshared_wordsig_blob(&pkcs7);

    let vba_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature" Target="vbaProjectSignature.bin"/>
</Relationships>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("xl/vbaProject.bin", options).unwrap();
    zip.write_all(&vba_project_bin).unwrap();

    zip.start_file("xl/_rels/vbaProject.bin.rels", options).unwrap();
    zip.write_all(vba_rels).unwrap();

    // Not an OLE container, not raw PKCS#7: WordSigBlob-wrapped PKCS#7 payload.
    zip.start_file("xl/vbaProjectSignature.bin", options).unwrap();
    zip.write_all(&wrapped).unwrap();

    let bytes = zip.finish().unwrap().into_inner();
    let pkg = XlsxPackage::from_bytes(&bytes).expect("read package");

    let sig = pkg
        .verify_vba_digital_signature()
        .expect("signature verification should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.binding, VbaSignatureBinding::Bound);
    assert_eq!(sig.stream_path, "xl/vbaProjectSignature.bin");

    let binding = pkg
        .vba_project_signature_binding()
        .expect("binding verification")
        .expect("project should be present");
    assert!(
        matches!(binding, VbaProjectBindingVerification::BoundVerified(_)),
        "expected BoundVerified, got {binding:?}"
    );

    // Tamper with a covered project stream but keep the signature bytes the same.
    let mut tampered_module = module1.to_vec();
    tampered_module[0] ^= 0xFF;
    let tampered_project = build_minimal_vba_project_bin(&tampered_module);

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("xl/vbaProject.bin", options).unwrap();
    zip.write_all(&tampered_project).unwrap();

    zip.start_file("xl/_rels/vbaProject.bin.rels", options).unwrap();
    zip.write_all(vba_rels).unwrap();

    zip.start_file("xl/vbaProjectSignature.bin", options).unwrap();
    zip.write_all(&wrapped).unwrap();

    let bytes = zip.finish().unwrap().into_inner();
    let pkg = XlsxPackage::from_bytes(&bytes).expect("read tampered package");

    let sig = pkg
        .verify_vba_digital_signature()
        .expect("signature verification should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.binding, VbaSignatureBinding::NotBound);
    assert_eq!(sig.stream_path, "xl/vbaProjectSignature.bin");

    let binding = pkg
        .vba_project_signature_binding()
        .expect("binding verification")
        .expect("project should be present");
    assert!(
        matches!(binding, VbaProjectBindingVerification::BoundMismatch(_)),
        "expected BoundMismatch, got {binding:?}"
    );
}

#[test]
fn verifies_raw_vba_project_signature_part_binding_for_v3_digest() {
    let vba_project_bin = build_minimal_vba_project_bin_v3(b"ABC");
    let digest =
        compute_vba_project_digest_v3(&vba_project_bin, DigestAlg::Sha256).expect("digest v3");
    assert_eq!(digest.len(), 32, "SHA-256 digest must be 32 bytes");

    let signed_content = make_spc_indirect_data_content_sha256(&digest);
    let pkcs7 = make_pkcs7_signed_message(&signed_content);

    let vba_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature" Target="vbaProjectSignature.bin"/>
</Relationships>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("xl/vbaProject.bin", options).unwrap();
    zip.write_all(&vba_project_bin).unwrap();

    zip.start_file("xl/_rels/vbaProject.bin.rels", options)
        .unwrap();
    zip.write_all(vba_rels).unwrap();

    // Raw PKCS#7/CMS bytes (not an OLE container).
    zip.start_file("xl/vbaProjectSignature.bin", options)
        .unwrap();
    zip.write_all(&pkcs7).unwrap();

    let bytes = zip.finish().unwrap().into_inner();
    let pkg = XlsxPackage::from_bytes(&bytes).expect("read package");

    let sig = pkg
        .verify_vba_digital_signature()
        .expect("signature verification should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(
        sig.binding,
        VbaSignatureBinding::Bound,
        "expected v3 digest binding to be verified for raw signature part"
    );
    assert_eq!(sig.stream_path, "xl/vbaProjectSignature.bin");

    let binding = pkg
        .vba_project_signature_binding()
        .expect("binding verification")
        .expect("project should be present");
    assert!(
        matches!(binding, VbaProjectBindingVerification::BoundVerified(_)),
        "expected BoundVerified, got {binding:?}"
    );
}

#[test]
fn verifies_raw_signature_part_binding_when_digsig_blob_wrapped_and_tampered_project_is_not_bound() {
    let module1 = b"Sub Hello()\r\nEnd Sub\r\n";
    let vba_project_bin = build_minimal_vba_project_bin(module1);

    // Signed digest is MD5(ContentNormalizedData) per MS-OVBA.
    let normalized = content_normalized_data(&vba_project_bin).expect("content normalized data");
    let digest = hash(MessageDigest::md5(), &normalized)
        .expect("md5 digest")
        .to_vec();

    let spc = make_spc_indirect_data_content_sha256(&digest);
    let pkcs7 = make_pkcs7_signed_message(&spc);
    let wrapped = build_oshared_digsig_blob(&pkcs7);

    let vba_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature" Target="vbaProjectSignature.bin"/>
</Relationships>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("xl/vbaProject.bin", options).unwrap();
    zip.write_all(&vba_project_bin).unwrap();

    zip.start_file("xl/_rels/vbaProject.bin.rels", options)
        .unwrap();
    zip.write_all(vba_rels).unwrap();

    // Raw DigSigBlob wrapper bytes (not an OLE container).
    zip.start_file("xl/vbaProjectSignature.bin", options)
        .unwrap();
    zip.write_all(&wrapped).unwrap();

    let bytes = zip.finish().unwrap().into_inner();
    let pkg = XlsxPackage::from_bytes(&bytes).expect("read package");

    let sig = pkg
        .verify_vba_digital_signature()
        .expect("signature verification should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.binding, VbaSignatureBinding::Bound);

    // Tamper with a covered project stream but keep the signature bytes the same.
    let mut tampered_module = module1.to_vec();
    tampered_module[0] ^= 0xFF;
    let tampered_project = build_minimal_vba_project_bin(&tampered_module);

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("xl/vbaProject.bin", options).unwrap();
    zip.write_all(&tampered_project).unwrap();

    zip.start_file("xl/_rels/vbaProject.bin.rels", options)
        .unwrap();
    zip.write_all(vba_rels).unwrap();

    zip.start_file("xl/vbaProjectSignature.bin", options)
        .unwrap();
    zip.write_all(&wrapped).unwrap();

    let bytes = zip.finish().unwrap().into_inner();
    let pkg = XlsxPackage::from_bytes(&bytes).expect("read tampered package");

    let sig = pkg
        .verify_vba_digital_signature()
        .expect("signature verification should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.binding, VbaSignatureBinding::NotBound);
}

#[test]
fn verifies_raw_signature_part_binding_when_digsig_info_serialized_wrapped() {
    let module1 = b"Sub Hello()\r\nEnd Sub\r\n";
    let vba_project_bin = build_minimal_vba_project_bin(module1);

    let normalized = content_normalized_data(&vba_project_bin).expect("content normalized data");
    let digest = hash(MessageDigest::md5(), &normalized)
        .expect("md5 digest")
        .to_vec();

    let spc = make_spc_indirect_data_content_sha256(&digest);
    let pkcs7 = make_pkcs7_signed_message(&spc);
    let wrapped = build_digsig_info_serialized_wrapper(&pkcs7);

    let vba_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature" Target="vbaProjectSignature.bin"/>
</Relationships>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("xl/vbaProject.bin", options).unwrap();
    zip.write_all(&vba_project_bin).unwrap();

    zip.start_file("xl/_rels/vbaProject.bin.rels", options)
        .unwrap();
    zip.write_all(vba_rels).unwrap();

    zip.start_file("xl/vbaProjectSignature.bin", options)
        .unwrap();
    zip.write_all(&wrapped).unwrap();

    let bytes = zip.finish().unwrap().into_inner();
    let pkg = XlsxPackage::from_bytes(&bytes).expect("read package");

    let sig = pkg
        .verify_vba_digital_signature()
        .expect("signature verification should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.binding, VbaSignatureBinding::Bound);

    let binding = pkg
        .vba_project_signature_binding()
        .expect("binding verification")
        .expect("project should be present");
    assert!(
        matches!(binding, VbaProjectBindingVerification::BoundVerified(_)),
        "expected BoundVerified, got {binding:?}"
    );

    // Tamper with a covered project stream but keep the signature bytes the same.
    let mut tampered_module = module1.to_vec();
    tampered_module[0] ^= 0xFF;
    let tampered_project = build_minimal_vba_project_bin(&tampered_module);

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("xl/vbaProject.bin", options).unwrap();
    zip.write_all(&tampered_project).unwrap();

    zip.start_file("xl/_rels/vbaProject.bin.rels", options)
        .unwrap();
    zip.write_all(vba_rels).unwrap();

    zip.start_file("xl/vbaProjectSignature.bin", options)
        .unwrap();
    zip.write_all(&wrapped).unwrap();

    let bytes = zip.finish().unwrap().into_inner();
    let pkg = XlsxPackage::from_bytes(&bytes).expect("read tampered package");

    let sig = pkg
        .verify_vba_digital_signature()
        .expect("signature verification should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.binding, VbaSignatureBinding::NotBound);

    let binding = pkg
        .vba_project_signature_binding()
        .expect("binding verification")
        .expect("project should be present");
    assert!(
        matches!(binding, VbaProjectBindingVerification::BoundMismatch(_)),
        "expected BoundMismatch, got {binding:?}"
    );
}

#[test]
fn verifies_raw_vba_project_signature_part_binding_for_v3_digest_when_digsig_blob_wrapped() {
    let vba_project_bin = build_minimal_vba_project_bin_v3(b"ABC");
    let digest =
        compute_vba_project_digest_v3(&vba_project_bin, DigestAlg::Sha256).expect("digest v3");
    assert_eq!(digest.len(), 32, "SHA-256 digest must be 32 bytes");

    let signed_content = make_spc_indirect_data_content_sha256(&digest);
    let pkcs7 = make_pkcs7_signed_message(&signed_content);
    let wrapped = build_oshared_digsig_blob(&pkcs7);

    let vba_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature" Target="vbaProjectSignature.bin"/>
</Relationships>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("xl/vbaProject.bin", options).unwrap();
    zip.write_all(&vba_project_bin).unwrap();

    zip.start_file("xl/_rels/vbaProject.bin.rels", options)
        .unwrap();
    zip.write_all(vba_rels).unwrap();

    zip.start_file("xl/vbaProjectSignature.bin", options)
        .unwrap();
    zip.write_all(&wrapped).unwrap();

    let bytes = zip.finish().unwrap().into_inner();
    let pkg = XlsxPackage::from_bytes(&bytes).expect("read package");

    let sig = pkg
        .verify_vba_digital_signature()
        .expect("signature verification should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.binding, VbaSignatureBinding::Bound);

    let binding = pkg
        .vba_project_signature_binding()
        .expect("binding verification")
        .expect("project should be present");
    assert!(
        matches!(binding, VbaProjectBindingVerification::BoundVerified(_)),
        "expected BoundVerified, got {binding:?}"
    );
}
