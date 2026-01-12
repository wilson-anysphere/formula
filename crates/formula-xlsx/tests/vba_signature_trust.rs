#![cfg(all(feature = "vba", not(target_arch = "wasm32")))]

use std::io::{Cursor, Write};

use formula_vba::{compute_vba_project_digest_v3, compress_container, DigestAlg, VbaSignatureBinding};
use formula_xlsx::vba::{VbaCertificateTrust, VbaSignatureTrustOptions, VbaSignatureVerification};
use formula_xlsx::XlsxPackage;
use openssl::x509::X509;
use zip::write::FileOptions;

mod vba_signature_test_utils;
use vba_signature_test_utils::{
    build_ole_with_streams, build_vba_signature_ole, make_pkcs7_signed_message, TEST_CERT_PEM,
    TEST_KEY_PEM,
};

fn build_zip(parts: &[(&str, &[u8])]) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    for (name, bytes) in parts {
        zip.start_file(name, options).expect("start zip file");
        zip.write_all(bytes).expect("write zip file");
    }

    zip.finish().expect("finish zip").into_inner()
}

fn push_record(out: &mut Vec<u8>, id: u16, data: &[u8]) {
    out.extend_from_slice(&id.to_le_bytes());
    out.extend_from_slice(&(data.len() as u32).to_le_bytes());
    out.extend_from_slice(data);
}

fn build_minimal_vba_project_bin_v3(designer_payload: &[u8]) -> Vec<u8> {
    let module_source = b"Sub Hello()\r\nEnd Sub\r\n";
    let module_container = compress_container(module_source);
    let userform_source = b"Sub FormHello()\r\nEnd Sub\r\n";
    let userform_container = compress_container(userform_source);

    // Minimal `dir` stream (decompressed) with:
    // - one standard module, and
    // - one UserForm module so FormsNormalizedData is non-empty.
    let dir_decompressed = {
        let mut out = Vec::new();

        // Include a v3-specific reference record so the transcript depends on it.
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

    // Designer payload so FormsNormalizedData is non-empty.
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
    let mut buf = Vec::new();
    let mut n = len;
    while n > 0 {
        buf.push((n & 0xFF) as u8);
        n >>= 8;
    }
    buf.reverse();
    let mut out = Vec::with_capacity(1 + buf.len());
    out.push(0x80 | (buf.len() as u8));
    out.extend_from_slice(&buf);
    out
}

fn der_tlv(tag: u8, content: &[u8]) -> Vec<u8> {
    let mut out = Vec::new();
    out.push(tag);
    out.extend_from_slice(&der_len(content.len()));
    out.extend_from_slice(content);
    out
}

fn der_sequence(items: &[Vec<u8>]) -> Vec<u8> {
    let mut content = Vec::new();
    for item in items {
        content.extend_from_slice(item);
    }
    der_tlv(0x30, &content)
}

fn der_oid(oid_content: &[u8]) -> Vec<u8> {
    der_tlv(0x06, oid_content)
}

fn der_null() -> Vec<u8> {
    vec![0x05, 0x00]
}

fn der_octet_string(bytes: &[u8]) -> Vec<u8> {
    der_tlv(0x04, bytes)
}

fn build_spc_indirect_data_content_sha256(project_digest: &[u8]) -> Vec<u8> {
    // SHA-256 OID: 2.16.840.1.101.3.4.2.1
    let sha256_oid = der_oid(&[0x60, 0x86, 0x48, 0x01, 0x65, 0x03, 0x04, 0x02, 0x01]);
    let alg_id = der_sequence(&[sha256_oid, der_null()]);

    // DigestInfo ::= SEQUENCE { digestAlgorithm AlgorithmIdentifier, digest OCTET STRING }
    let digest_info = der_sequence(&[alg_id, der_octet_string(project_digest)]);

    // SpcIndirectDataContent ::= SEQUENCE { data, messageDigest }
    // `data` is ignored by our parser; use NULL.
    der_sequence(&[der_null(), digest_info])
}

fn build_oshared_digsig_blob(valid_pkcs7: &[u8]) -> Vec<u8> {
    // MS-OSHARED describes a DigSigBlob wrapper around the PKCS#7 signature bytes.
    //
    // This test blob intentionally contains *multiple* PKCS#7 SignedData blobs:
    // - a corrupted (but still parseable) one early, and
    // - a corrupted one after the real signature.
    //
    // This ensures verification succeeds only when the DigSigBlob offsets are honored, not when
    // relying on heuristic scan-for-0x30 fallbacks.
    let mut invalid_pkcs7 = valid_pkcs7.to_vec();
    let last = invalid_pkcs7.len().saturating_sub(1);
    invalid_pkcs7[last] ^= 0xFF;

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

    // Corrupt the signature bytes while keeping the overall ASN.1 shape parseable.
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

fn make_unrelated_root_cert_der() -> Vec<u8> {
    use openssl::asn1::Asn1Time;
    use openssl::bn::BigNum;
    use openssl::hash::MessageDigest;
    use openssl::pkey::PKey;
    use openssl::x509::{X509Builder, X509NameBuilder};

    let pkey = PKey::private_key_from_pem(TEST_KEY_PEM.as_bytes()).expect("parse private key");

    let mut name_builder = X509NameBuilder::new().expect("x509 name builder");
    name_builder
        .append_entry_by_text("CN", "Formula VBA Unrelated Root")
        .expect("CN");
    let name = name_builder.build();

    let mut builder = X509Builder::new().expect("x509 builder");
    builder.set_version(2).expect("set version");

    let serial = BigNum::from_u32(2)
        .expect("serial bn")
        .to_asn1_integer()
        .expect("serial integer");
    builder.set_serial_number(&serial).expect("serial");

    builder.set_subject_name(&name).expect("subject name");
    builder.set_issuer_name(&name).expect("issuer name");
    builder.set_pubkey(&pkey).expect("pubkey");

    builder
        .set_not_before(&Asn1Time::days_from_now(0).expect("not before"))
        .expect("set not before");
    builder
        .set_not_after(&Asn1Time::days_from_now(3650).expect("not after"))
        .expect("set not after");

    builder
        .sign(&pkey, MessageDigest::sha256())
        .expect("sign");
    builder.build().to_der().expect("DER")
}

#[test]
fn embedded_signature_stream_trust_is_reported() {
    let pkcs7 = make_pkcs7_signed_message(b"formula-xlsx-trust-test");
    let vba_project_bin = build_vba_signature_ole(&pkcs7);
    let zip_bytes = build_zip(&[("xl/vbaProject.bin", &vba_project_bin)]);

    let pkg = XlsxPackage::from_bytes(&zip_bytes).expect("read package");

    let signer_der = X509::from_pem(TEST_CERT_PEM.as_bytes())
        .expect("parse test certificate")
        .to_der()
        .expect("convert test certificate to DER");
    let wrong_root_der = make_unrelated_root_cert_der();

    let sig = pkg
        .verify_vba_digital_signature_with_trust(&VbaSignatureTrustOptions::default())
        .expect("verify signature")
        .expect("signature should be present");
    assert_eq!(sig.signature.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.cert_trust, VbaCertificateTrust::Unknown);

    let sig = pkg
        .verify_vba_digital_signature_with_trust(&VbaSignatureTrustOptions {
            trusted_root_certs_der: vec![signer_der.clone()],
        })
        .expect("verify signature")
        .expect("signature should be present");
    assert_eq!(sig.signature.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.cert_trust, VbaCertificateTrust::Trusted);

    let sig = pkg
        .verify_vba_digital_signature_with_trust(&VbaSignatureTrustOptions {
            trusted_root_certs_der: vec![wrong_root_der],
        })
        .expect("verify signature")
        .expect("signature should be present");
    assert_eq!(sig.signature.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.cert_trust, VbaCertificateTrust::Untrusted);
}

#[test]
fn embedded_signature_stream_trust_is_reported_with_leading_slash_part_names() {
    // Some producers incorrectly store OPC part names with a leading `/` in the ZIP.
    // Ensure signature verification and trust evaluation still work.
    let pkcs7 = make_pkcs7_signed_message(b"formula-xlsx-trust-test");
    let vba_project_bin = build_vba_signature_ole(&pkcs7);
    let zip_bytes = build_zip(&[("/xl/vbaProject.bin", &vba_project_bin)]);

    let pkg = XlsxPackage::from_bytes(&zip_bytes).expect("read package");

    let signer_der = X509::from_pem(TEST_CERT_PEM.as_bytes())
        .expect("parse test certificate")
        .to_der()
        .expect("convert test certificate to DER");
    let wrong_root_der = make_unrelated_root_cert_der();

    let sig = pkg
        .verify_vba_digital_signature_with_trust(&VbaSignatureTrustOptions::default())
        .expect("verify signature")
        .expect("signature should be present");
    assert_eq!(sig.signature.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.cert_trust, VbaCertificateTrust::Unknown);

    let sig = pkg
        .verify_vba_digital_signature_with_trust(&VbaSignatureTrustOptions {
            trusted_root_certs_der: vec![signer_der.clone()],
        })
        .expect("verify signature")
        .expect("signature should be present");
    assert_eq!(sig.signature.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.cert_trust, VbaCertificateTrust::Trusted);

    let sig = pkg
        .verify_vba_digital_signature_with_trust(&VbaSignatureTrustOptions {
            trusted_root_certs_der: vec![wrong_root_der],
        })
        .expect("verify signature")
        .expect("signature should be present");
    assert_eq!(sig.signature.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.cert_trust, VbaCertificateTrust::Untrusted);
}

#[test]
fn external_signature_part_trust_is_reported() {
    let pkcs7 = make_pkcs7_signed_message(b"formula-xlsx-trust-test");

    // `xl/vbaProject.bin` must be a valid OLE file (even if unsigned) so the fallback embedded
    // signature scan can run without errors.
    let vba_project_bin = build_ole_with_streams(&[]);

    let vba_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature" Target="vbaProjectSignature.bin"/>
</Relationships>"#;

    let zip_bytes = build_zip(&[
        ("xl/vbaProject.bin", &vba_project_bin),
        ("xl/_rels/vbaProject.bin.rels", vba_rels),
        ("xl/vbaProjectSignature.bin", &pkcs7),
    ]);
    let pkg = XlsxPackage::from_bytes(&zip_bytes).expect("read package");

    let signer_der = X509::from_pem(TEST_CERT_PEM.as_bytes())
        .expect("parse test certificate")
        .to_der()
        .expect("convert test certificate to DER");
    let wrong_root_der = make_unrelated_root_cert_der();

    let sig = pkg
        .verify_vba_digital_signature_with_trust(&VbaSignatureTrustOptions::default())
        .expect("verify signature")
        .expect("signature should be present");
    assert_eq!(sig.signature.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.cert_trust, VbaCertificateTrust::Unknown);

    let sig = pkg
        .verify_vba_digital_signature_with_trust(&VbaSignatureTrustOptions {
            trusted_root_certs_der: vec![signer_der.clone()],
        })
        .expect("verify signature")
        .expect("signature should be present");
    assert_eq!(sig.signature.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.cert_trust, VbaCertificateTrust::Trusted);

    let sig = pkg
        .verify_vba_digital_signature_with_trust(&VbaSignatureTrustOptions {
            trusted_root_certs_der: vec![wrong_root_der],
        })
        .expect("verify signature")
        .expect("signature should be present");
    assert_eq!(sig.signature.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.cert_trust, VbaCertificateTrust::Untrusted);
}

#[test]
fn external_signature_part_trust_is_reported_with_leading_slash_part_names() {
    let pkcs7 = make_pkcs7_signed_message(b"formula-xlsx-trust-test");

    // `xl/vbaProject.bin` must be a valid OLE file (even if unsigned) so the fallback embedded
    // signature scan can run without errors.
    let vba_project_bin = build_ole_with_streams(&[]);

    let vba_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature" Target="vbaProjectSignature.bin"/>
</Relationships>"#;

    let zip_bytes = build_zip(&[
        ("/xl/vbaProject.bin", &vba_project_bin),
        ("/xl/_rels/vbaProject.bin.rels", vba_rels),
        ("/xl/vbaProjectSignature.bin", &pkcs7),
    ]);
    let pkg = XlsxPackage::from_bytes(&zip_bytes).expect("read package");

    let signer_der = X509::from_pem(TEST_CERT_PEM.as_bytes())
        .expect("parse test certificate")
        .to_der()
        .expect("convert test certificate to DER");
    let wrong_root_der = make_unrelated_root_cert_der();

    let sig = pkg
        .verify_vba_digital_signature_with_trust(&VbaSignatureTrustOptions::default())
        .expect("verify signature")
        .expect("signature should be present");
    assert_eq!(sig.signature.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.cert_trust, VbaCertificateTrust::Unknown);

    let sig = pkg
        .verify_vba_digital_signature_with_trust(&VbaSignatureTrustOptions {
            trusted_root_certs_der: vec![signer_der.clone()],
        })
        .expect("verify signature")
        .expect("signature should be present");
    assert_eq!(sig.signature.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.cert_trust, VbaCertificateTrust::Trusted);

    let sig = pkg
        .verify_vba_digital_signature_with_trust(&VbaSignatureTrustOptions {
            trusted_root_certs_der: vec![wrong_root_der],
        })
        .expect("verify signature")
        .expect("signature should be present");
    assert_eq!(sig.signature.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.cert_trust, VbaCertificateTrust::Untrusted);
}

#[test]
fn external_raw_signature_part_binding_v3_and_trust_is_reported() {
    let vba_project_bin = build_minimal_vba_project_bin_v3(b"ABC");
    let digest =
        compute_vba_project_digest_v3(&vba_project_bin, DigestAlg::Sha256).expect("digest v3");

    let signed_content = build_spc_indirect_data_content_sha256(&digest);
    let pkcs7 = make_pkcs7_signed_message(&signed_content);

    let vba_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature" Target="vbaProjectSignature.bin"/>
</Relationships>"#;

    let zip_bytes = build_zip(&[
        ("xl/vbaProject.bin", &vba_project_bin),
        ("xl/_rels/vbaProject.bin.rels", vba_rels),
        ("xl/vbaProjectSignature.bin", &pkcs7),
    ]);
    let pkg = XlsxPackage::from_bytes(&zip_bytes).expect("read package");

    let signer_der = X509::from_pem(TEST_CERT_PEM.as_bytes())
        .expect("parse test certificate")
        .to_der()
        .expect("convert test certificate to DER");

    // No trust anchors: trust is unknown, but binding should still verify.
    let sig = pkg
        .verify_vba_digital_signature_with_trust(&VbaSignatureTrustOptions::default())
        .expect("verify signature")
        .expect("signature should be present");
    assert_eq!(sig.signature.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.signature.binding, VbaSignatureBinding::Bound);
    assert_eq!(sig.cert_trust, VbaCertificateTrust::Unknown);

    // Trusted publisher (root matches signer): binding should still be bound and trust trusted.
    let sig = pkg
        .verify_vba_digital_signature_with_trust(&VbaSignatureTrustOptions {
            trusted_root_certs_der: vec![signer_der],
        })
        .expect("verify signature")
        .expect("signature should be present");
    assert_eq!(sig.signature.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.signature.binding, VbaSignatureBinding::Bound);
    assert_eq!(sig.cert_trust, VbaCertificateTrust::Trusted);
}

#[test]
fn external_raw_signature_part_binding_v3_and_trust_is_reported_when_digsig_blob_wrapped() {
    let vba_project_bin = build_minimal_vba_project_bin_v3(b"ABC");
    let digest =
        compute_vba_project_digest_v3(&vba_project_bin, DigestAlg::Sha256).expect("digest v3");

    let signed_content = build_spc_indirect_data_content_sha256(&digest);
    let pkcs7 = make_pkcs7_signed_message(&signed_content);
    let wrapped = build_oshared_digsig_blob(&pkcs7);

    let vba_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature" Target="vbaProjectSignature.bin"/>
</Relationships>"#;

    let zip_bytes = build_zip(&[
        ("xl/vbaProject.bin", &vba_project_bin),
        ("xl/_rels/vbaProject.bin.rels", vba_rels),
        ("xl/vbaProjectSignature.bin", &wrapped),
    ]);
    let pkg = XlsxPackage::from_bytes(&zip_bytes).expect("read package");

    let signer_der = X509::from_pem(TEST_CERT_PEM.as_bytes())
        .expect("parse test certificate")
        .to_der()
        .expect("convert test certificate to DER");

    // No trust anchors: trust is unknown, but binding should still verify.
    let sig = pkg
        .verify_vba_digital_signature_with_trust(&VbaSignatureTrustOptions::default())
        .expect("verify signature")
        .expect("signature should be present");
    assert_eq!(sig.signature.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.signature.binding, VbaSignatureBinding::Bound);
    assert_eq!(sig.signature.stream_path, "xl/vbaProjectSignature.bin");
    assert_eq!(sig.cert_trust, VbaCertificateTrust::Unknown);

    // Trusted publisher (root matches signer): binding should still be bound and trust trusted.
    let sig = pkg
        .verify_vba_digital_signature_with_trust(&VbaSignatureTrustOptions {
            trusted_root_certs_der: vec![signer_der],
        })
        .expect("verify signature")
        .expect("signature should be present");
    assert_eq!(sig.signature.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.signature.binding, VbaSignatureBinding::Bound);
    assert_eq!(sig.signature.stream_path, "xl/vbaProjectSignature.bin");
    assert_eq!(sig.cert_trust, VbaCertificateTrust::Trusted);
}

#[test]
fn external_raw_signature_part_binding_v3_and_trust_is_reported_when_wordsig_blob_wrapped() {
    let vba_project_bin = build_minimal_vba_project_bin_v3(b"ABC");
    let digest =
        compute_vba_project_digest_v3(&vba_project_bin, DigestAlg::Sha256).expect("digest v3");

    let signed_content = build_spc_indirect_data_content_sha256(&digest);
    let pkcs7 = make_pkcs7_signed_message(&signed_content);
    let wrapped = build_oshared_wordsig_blob(&pkcs7);

    let vba_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature" Target="vbaProjectSignature.bin"/>
</Relationships>"#;

    let zip_bytes = build_zip(&[
        ("xl/vbaProject.bin", &vba_project_bin),
        ("xl/_rels/vbaProject.bin.rels", vba_rels),
        ("xl/vbaProjectSignature.bin", &wrapped),
    ]);
    let pkg = XlsxPackage::from_bytes(&zip_bytes).expect("read package");

    let signer_der = X509::from_pem(TEST_CERT_PEM.as_bytes())
        .expect("parse test certificate")
        .to_der()
        .expect("convert test certificate to DER");

    // No trust anchors: trust is unknown, but binding should still verify.
    let sig = pkg
        .verify_vba_digital_signature_with_trust(&VbaSignatureTrustOptions::default())
        .expect("verify signature")
        .expect("signature should be present");
    assert_eq!(sig.signature.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.signature.binding, VbaSignatureBinding::Bound);
    assert_eq!(sig.signature.stream_path, "xl/vbaProjectSignature.bin");
    assert_eq!(sig.cert_trust, VbaCertificateTrust::Unknown);

    // Trusted publisher (root matches signer): binding should still be bound and trust trusted.
    let sig = pkg
        .verify_vba_digital_signature_with_trust(&VbaSignatureTrustOptions {
            trusted_root_certs_der: vec![signer_der],
        })
        .expect("verify signature")
        .expect("signature should be present");
    assert_eq!(sig.signature.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.signature.binding, VbaSignatureBinding::Bound);
    assert_eq!(sig.signature.stream_path, "xl/vbaProjectSignature.bin");
    assert_eq!(sig.cert_trust, VbaCertificateTrust::Trusted);
}

#[test]
fn external_signature_part_ole_container_trust_is_reported() {
    let pkcs7 = make_pkcs7_signed_message(b"formula-xlsx-trust-test");
    let signature_part_ole = build_vba_signature_ole(&pkcs7);

    // Use a minimal (unsigned) OLE container for `xl/vbaProject.bin` so binding verification can
    // run if needed.
    let vba_project_bin = build_ole_with_streams(&[]);

    let vba_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature" Target="vbaProjectSignature.bin"/>
</Relationships>"#;

    let zip_bytes = build_zip(&[
        ("xl/vbaProject.bin", &vba_project_bin),
        ("xl/_rels/vbaProject.bin.rels", vba_rels),
        ("xl/vbaProjectSignature.bin", &signature_part_ole),
    ]);
    let pkg = XlsxPackage::from_bytes(&zip_bytes).expect("read package");

    let signer_der = X509::from_pem(TEST_CERT_PEM.as_bytes())
        .expect("parse test certificate")
        .to_der()
        .expect("convert test certificate to DER");
    let wrong_root_der = make_unrelated_root_cert_der();

    let sig = pkg
        .verify_vba_digital_signature_with_trust(&VbaSignatureTrustOptions::default())
        .expect("verify signature")
        .expect("signature should be present");
    assert_eq!(sig.signature.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.cert_trust, VbaCertificateTrust::Unknown);

    let sig = pkg
        .verify_vba_digital_signature_with_trust(&VbaSignatureTrustOptions {
            trusted_root_certs_der: vec![signer_der.clone()],
        })
        .expect("verify signature")
        .expect("signature should be present");
    assert_eq!(sig.signature.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.cert_trust, VbaCertificateTrust::Trusted);

    let sig = pkg
        .verify_vba_digital_signature_with_trust(&VbaSignatureTrustOptions {
            trusted_root_certs_der: vec![wrong_root_der],
        })
        .expect("verify signature")
        .expect("signature should be present");
    assert_eq!(sig.signature.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.cert_trust, VbaCertificateTrust::Untrusted);
}

#[test]
fn external_signature_part_ole_container_digital_signature_ext_binding_v3_and_trust_is_reported() {
    let vba_project_bin = build_minimal_vba_project_bin_v3(b"ABC");
    let digest =
        compute_vba_project_digest_v3(&vba_project_bin, DigestAlg::Sha256).expect("digest v3");

    let signed_content = build_spc_indirect_data_content_sha256(&digest);
    let pkcs7 = make_pkcs7_signed_message(&signed_content);
    let signature_part_ole =
        build_ole_with_streams(&[("\u{0005}DigitalSignatureExt", pkcs7.as_slice())]);

    let vba_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature" Target="vbaProjectSignature.bin"/>
</Relationships>"#;

    let zip_bytes = build_zip(&[
        ("xl/vbaProject.bin", &vba_project_bin),
        ("xl/_rels/vbaProject.bin.rels", vba_rels),
        ("xl/vbaProjectSignature.bin", &signature_part_ole),
    ]);
    let pkg = XlsxPackage::from_bytes(&zip_bytes).expect("read package");

    let signer_der = X509::from_pem(TEST_CERT_PEM.as_bytes())
        .expect("parse test certificate")
        .to_der()
        .expect("convert test certificate to DER");

    // No trust anchors: trust is unknown, but binding should still verify.
    let sig = pkg
        .verify_vba_digital_signature_with_trust(&VbaSignatureTrustOptions::default())
        .expect("verify signature")
        .expect("signature should be present");
    assert_eq!(sig.signature.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.signature.binding, VbaSignatureBinding::Bound);
    assert!(
        sig.signature.stream_path.contains("DigitalSignatureExt"),
        "expected stream path to contain DigitalSignatureExt, got: {}",
        sig.signature.stream_path
    );
    assert_eq!(sig.cert_trust, VbaCertificateTrust::Unknown);

    // Trusted publisher (root matches signer): binding should still be bound and trust trusted.
    let sig = pkg
        .verify_vba_digital_signature_with_trust(&VbaSignatureTrustOptions {
            trusted_root_certs_der: vec![signer_der],
        })
        .expect("verify signature")
        .expect("signature should be present");
    assert_eq!(sig.signature.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.signature.binding, VbaSignatureBinding::Bound);
    assert_eq!(sig.cert_trust, VbaCertificateTrust::Trusted);
}

#[test]
fn external_signature_part_ole_container_digital_signature_ext_binding_v3_and_trust_is_reported_when_digsig_blob_wrapped(
) {
    let vba_project_bin = build_minimal_vba_project_bin_v3(b"ABC");
    let digest =
        compute_vba_project_digest_v3(&vba_project_bin, DigestAlg::Sha256).expect("digest v3");

    let signed_content = build_spc_indirect_data_content_sha256(&digest);
    let pkcs7 = make_pkcs7_signed_message(&signed_content);
    let wrapped = build_oshared_digsig_blob(&pkcs7);
    let signature_part_ole =
        build_ole_with_streams(&[("\u{0005}DigitalSignatureExt", wrapped.as_slice())]);

    let vba_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature" Target="vbaProjectSignature.bin"/>
</Relationships>"#;

    let zip_bytes = build_zip(&[
        ("xl/vbaProject.bin", &vba_project_bin),
        ("xl/_rels/vbaProject.bin.rels", vba_rels),
        ("xl/vbaProjectSignature.bin", &signature_part_ole),
    ]);
    let pkg = XlsxPackage::from_bytes(&zip_bytes).expect("read package");

    let signer_der = X509::from_pem(TEST_CERT_PEM.as_bytes())
        .expect("parse test certificate")
        .to_der()
        .expect("convert test certificate to DER");

    let sig = pkg
        .verify_vba_digital_signature_with_trust(&VbaSignatureTrustOptions::default())
        .expect("verify signature")
        .expect("signature should be present");
    assert_eq!(sig.signature.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.signature.binding, VbaSignatureBinding::Bound);
    assert_eq!(sig.cert_trust, VbaCertificateTrust::Unknown);

    let sig = pkg
        .verify_vba_digital_signature_with_trust(&VbaSignatureTrustOptions {
            trusted_root_certs_der: vec![signer_der],
        })
        .expect("verify signature")
        .expect("signature should be present");
    assert_eq!(sig.signature.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.signature.binding, VbaSignatureBinding::Bound);
    assert_eq!(sig.cert_trust, VbaCertificateTrust::Trusted);
}

#[test]
fn external_signature_part_ole_container_digital_signature_ext_binding_v3_and_trust_is_reported_when_wordsig_blob_wrapped(
) {
    let vba_project_bin = build_minimal_vba_project_bin_v3(b"ABC");
    let digest =
        compute_vba_project_digest_v3(&vba_project_bin, DigestAlg::Sha256).expect("digest v3");

    let signed_content = build_spc_indirect_data_content_sha256(&digest);
    let pkcs7 = make_pkcs7_signed_message(&signed_content);
    let wrapped = build_oshared_wordsig_blob(&pkcs7);
    let signature_part_ole =
        build_ole_with_streams(&[("\u{0005}DigitalSignatureExt", wrapped.as_slice())]);

    let vba_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature" Target="vbaProjectSignature.bin"/>
</Relationships>"#;

    let zip_bytes = build_zip(&[
        ("xl/vbaProject.bin", &vba_project_bin),
        ("xl/_rels/vbaProject.bin.rels", vba_rels),
        ("xl/vbaProjectSignature.bin", &signature_part_ole),
    ]);
    let pkg = XlsxPackage::from_bytes(&zip_bytes).expect("read package");

    let signer_der = X509::from_pem(TEST_CERT_PEM.as_bytes())
        .expect("parse test certificate")
        .to_der()
        .expect("convert test certificate to DER");

    let sig = pkg
        .verify_vba_digital_signature_with_trust(&VbaSignatureTrustOptions::default())
        .expect("verify signature")
        .expect("signature should be present");
    assert_eq!(sig.signature.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.signature.binding, VbaSignatureBinding::Bound);
    assert_eq!(sig.cert_trust, VbaCertificateTrust::Unknown);

    let sig = pkg
        .verify_vba_digital_signature_with_trust(&VbaSignatureTrustOptions {
            trusted_root_certs_der: vec![signer_der],
        })
        .expect("verify signature")
        .expect("signature should be present");
    assert_eq!(sig.signature.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.signature.binding, VbaSignatureBinding::Bound);
    assert_eq!(sig.cert_trust, VbaCertificateTrust::Trusted);
}

#[test]
fn returns_none_when_vba_project_bin_is_missing() {
    let pkcs7 = make_pkcs7_signed_message(b"formula-xlsx-trust-test");
    let zip_bytes = build_zip(&[("xl/vbaProjectSignature.bin", &pkcs7)]);

    let pkg = XlsxPackage::from_bytes(&zip_bytes).expect("read package");
    let sig = pkg
        .verify_vba_digital_signature_with_trust(&VbaSignatureTrustOptions::default())
        .expect("verify signature");
    assert!(sig.is_none());
}
