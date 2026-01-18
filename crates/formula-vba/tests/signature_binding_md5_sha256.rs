#![cfg(not(target_arch = "wasm32"))]

use std::io::{Cursor, Write};

use formula_vba::{
    compress_container, content_normalized_data, verify_vba_digital_signature,
    verify_vba_project_signature_binding, VbaProjectBindingVerification, VbaSignatureBinding,
    VbaSignatureVerification,
};
use md5::{Digest as _, Md5};

mod signature_test_utils;
use signature_test_utils::{TEST_CERT_PEM, TEST_KEY_PEM};

#[test]
fn ms_oshared_md5_digest_bytes_even_when_signeddata_uses_sha256() {
    use openssl::pkcs7::{Pkcs7, Pkcs7Flags};
    use openssl::pkey::PKey;
    use openssl::stack::Stack;
    use openssl::x509::X509;

    // ---- 1) Build a minimal spec-ish VBA project and compute its MS-OVBA content hash (MD5). ----
    let module_source = b"Sub Hello()\r\nEnd Sub\r\n";
    let module_container = compress_container(module_source);

    // Minimal `dir` stream (decompressed form) with a single module.
    let dir_decompressed = {
        let mut out = Vec::new();
        // PROJECTNAME
        push_record(&mut out, 0x0004, b"VBAProject");
        // MODULENAME
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
        out
    };
    let dir_container = compress_container(&dir_decompressed);

    let project_stream_bytes: &[u8] = b"Name=\"VBAProject\"\r\nModule=Module1\r\n";
    let vba_project_stream_bytes: &[u8] = b"dummy";

    let unsigned_vba_project_bin = build_vba_project_bin_with_streams(&[
        ("PROJECT", project_stream_bytes),
        ("VBA/_VBA_PROJECT", vba_project_stream_bytes),
        ("VBA/dir", &dir_container),
        ("VBA/Module1", &module_container),
    ]);

    let normalized = content_normalized_data(&unsigned_vba_project_bin).expect("ContentNormalizedData");
    let project_md5: [u8; 16] = Md5::digest(&normalized).into();
    assert_eq!(project_md5.len(), 16, "VBA project digest must be 16-byte MD5");

    // ---- 2) Construct SpcIndirectDataContent with sha256 OID but MD5 digest bytes. ----
    // DigestInfo.digestAlgorithm.algorithm = sha256 (2.16.840.1.101.3.4.2.1)
    // DigestInfo.digestAlgorithm.parameters = NULL (per MS-OSHARED)
    // DigestInfo.digest = MD5(project)
    let spc_indirect_data_content =
        build_spc_indirect_data_content_sha256_oid_with_md5_digest(&project_md5);

    // ---- 3) Produce PKCS#7 SignedData using OpenSSL (signing with SHA-256 by default). ----
    let pkey = PKey::private_key_from_pem(TEST_KEY_PEM.as_bytes()).expect("parse private key");
    let cert = X509::from_pem(TEST_CERT_PEM.as_bytes()).expect("parse certificate");
    let extra_certs = Stack::new().expect("create cert stack");

    let pkcs7 = Pkcs7::sign(
        &cert,
        &pkey,
        &extra_certs,
        &spc_indirect_data_content,
        // NOATTR keeps output deterministic.
        Pkcs7Flags::BINARY | Pkcs7Flags::NOATTR,
    )
    .expect("pkcs7 sign");
    let pkcs7_der = pkcs7.to_der().expect("pkcs7 DER");

    // ---- 4) Store signature in a \x05DigitalSignature stream. ----
    let signed_streams = vec![
        ("PROJECT", project_stream_bytes),
        ("VBA/_VBA_PROJECT", vba_project_stream_bytes),
        ("VBA/dir", dir_container.as_slice()),
        ("VBA/Module1", module_container.as_slice()),
        ("\u{0005}DigitalSignature", pkcs7_der.as_slice()),
    ];
    let vba_project_bin = build_vba_project_bin_with_streams(&signed_streams);

    // ---- 5) Verify ----
    let sig = verify_vba_digital_signature(&vba_project_bin)
        .expect("signature inspection should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(
        sig.binding,
        VbaSignatureBinding::Bound,
        "expected signature binding to be Bound even when DigestInfo.digestAlgorithm is sha256 but digest bytes are MD5"
    );

    // Also exercise the "binding-only" helper API which takes the signature bytes separately
    // (e.g. `xl/vbaProjectSignature.bin` in OOXML packages).
    let binding = verify_vba_project_signature_binding(&vba_project_bin, &pkcs7_der)
        .expect("binding verification should succeed");
    match binding {
        VbaProjectBindingVerification::BoundVerified(debug) => {
            assert_eq!(
                debug.hash_algorithm_oid.as_deref(),
                Some("2.16.840.1.101.3.4.2.1")
            );
            assert_eq!(debug.hash_algorithm_name.as_deref(), Some("SHA-256"));
            assert_eq!(debug.signed_digest.as_deref(), Some(project_md5.as_slice()));
            assert_eq!(debug.computed_digest.as_deref(), Some(project_md5.as_slice()));
        }
        other => panic!("expected BoundVerified, got {other:?}"),
    }
}

#[test]
fn ms_oshared_md5_source_hash_even_when_spc_indirect_data_content_v2_advertises_sha256() {
    use openssl::pkcs7::{Pkcs7, Pkcs7Flags};
    use openssl::pkey::PKey;
    use openssl::stack::Stack;
    use openssl::x509::X509;

    // ---- 1) Build a minimal spec-ish VBA project and compute its MS-OVBA content hash (MD5). ----
    let module_source = b"Sub Hello()\r\nEnd Sub\r\n";
    let module_container = compress_container(module_source);

    // Minimal `dir` stream (decompressed form) with a single module.
    let dir_decompressed = {
        let mut out = Vec::new();
        // PROJECTNAME
        push_record(&mut out, 0x0004, b"VBAProject");
        // MODULENAME
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
        out
    };
    let dir_container = compress_container(&dir_decompressed);

    let project_stream_bytes: &[u8] = b"Name=\"VBAProject\"\r\nModule=Module1\r\n";
    let vba_project_stream_bytes: &[u8] = b"dummy";

    let unsigned_vba_project_bin = build_vba_project_bin_with_streams(&[
        ("PROJECT", project_stream_bytes),
        ("VBA/_VBA_PROJECT", vba_project_stream_bytes),
        ("VBA/dir", &dir_container),
        ("VBA/Module1", &module_container),
    ]);

    let normalized =
        content_normalized_data(&unsigned_vba_project_bin).expect("ContentNormalizedData");
    let project_md5: [u8; 16] = Md5::digest(&normalized).into();
    assert_eq!(project_md5.len(), 16, "VBA project digest must be 16-byte MD5");

    // ---- 2) Construct SpcIndirectDataContentV2 with sha256 algorithm id but MD5 source hash. ----
    let spc_v2 =
        build_spc_indirect_data_content_v2_with_sha256_algorithm_and_md5_source_hash(&project_md5);

    // ---- 3) Produce PKCS#7 SignedData using OpenSSL (signing with SHA-256 by default). ----
    let pkey = PKey::private_key_from_pem(TEST_KEY_PEM.as_bytes()).expect("parse private key");
    let cert = X509::from_pem(TEST_CERT_PEM.as_bytes()).expect("parse certificate");
    let extra_certs = Stack::new().expect("create cert stack");

    let pkcs7 = Pkcs7::sign(
        &cert,
        &pkey,
        &extra_certs,
        &spc_v2,
        // NOATTR keeps output deterministic.
        Pkcs7Flags::BINARY | Pkcs7Flags::NOATTR,
    )
    .expect("pkcs7 sign");
    let pkcs7_der = pkcs7.to_der().expect("pkcs7 DER");

    // ---- 4) Store signature in a \x05DigitalSignature stream. ----
    let signed_streams = vec![
        ("PROJECT", project_stream_bytes),
        ("VBA/_VBA_PROJECT", vba_project_stream_bytes),
        ("VBA/dir", dir_container.as_slice()),
        ("VBA/Module1", module_container.as_slice()),
        ("\u{0005}DigitalSignature", pkcs7_der.as_slice()),
    ];
    let vba_project_bin = build_vba_project_bin_with_streams(&signed_streams);

    // ---- 5) Verify ----
    let sig = verify_vba_digital_signature(&vba_project_bin)
        .expect("signature inspection should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(
        sig.binding,
        VbaSignatureBinding::Bound,
        "expected signature binding to be Bound: SpcIndirectDataContentV2 sourceHash must be MD5 (16 bytes) even when algorithmId indicates SHA-256"
    );

    // Also exercise the binding-only helper API with a raw PKCS#7 blob (like OOXML external
    // signature parts).
    let binding = verify_vba_project_signature_binding(&unsigned_vba_project_bin, &pkcs7_der)
        .expect("binding verification should succeed");
    match binding {
        VbaProjectBindingVerification::BoundVerified(debug) => {
            // For V2 signatures we normalize to MD5 since the SigData sourceHash is always MD5 per
            // MS-OSHARED §4.3.
            assert_eq!(debug.hash_algorithm_oid.as_deref(), Some("1.2.840.113549.2.5"));
            assert_eq!(debug.hash_algorithm_name.as_deref(), Some("MD5"));
            assert_eq!(debug.signed_digest.as_deref(), Some(project_md5.as_slice()));
            assert_eq!(debug.computed_digest.as_deref(), Some(project_md5.as_slice()));
        }
        other => panic!("expected BoundVerified, got {other:?}"),
    }
}

#[test]
fn verify_vba_project_signature_binding_md5_digest_bytes_even_when_oid_is_sha256() {
    use openssl::pkcs7::{Pkcs7, Pkcs7Flags};
    use openssl::pkey::PKey;
    use openssl::stack::Stack;
    use openssl::x509::X509;

    // ---- 1) Build a minimal spec-ish VBA project and compute its MS-OVBA content hash (MD5). ----
    let module_source = b"Sub Hello()\r\nEnd Sub\r\n";
    let module_container = compress_container(module_source);

    // Minimal `dir` stream (decompressed form) with a single module.
    let dir_decompressed = {
        let mut out = Vec::new();
        // PROJECTNAME
        push_record(&mut out, 0x0004, b"VBAProject");
        // MODULENAME
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
        out
    };
    let dir_container = compress_container(&dir_decompressed);

    let project_stream_bytes: &[u8] = b"Name=\"VBAProject\"\r\nModule=Module1\r\n";
    let vba_project_stream_bytes: &[u8] = b"dummy";

    let unsigned_vba_project_bin = build_vba_project_bin_with_streams(&[
        ("PROJECT", project_stream_bytes),
        ("VBA/_VBA_PROJECT", vba_project_stream_bytes),
        ("VBA/dir", &dir_container),
        ("VBA/Module1", &module_container),
    ]);

    let normalized =
        content_normalized_data(&unsigned_vba_project_bin).expect("ContentNormalizedData");
    let project_md5: [u8; 16] = Md5::digest(&normalized).into();
    assert_eq!(project_md5.len(), 16, "VBA project digest must be 16-byte MD5");

    // ---- 2) Construct SpcIndirectDataContent with sha256 OID but MD5 digest bytes. ----
    let spc_indirect_data_content =
        build_spc_indirect_data_content_sha256_oid_with_md5_digest(&project_md5);

    // ---- 3) Produce PKCS#7 SignedData using OpenSSL (signing with SHA-256 by default). ----
    let pkey = PKey::private_key_from_pem(TEST_KEY_PEM.as_bytes()).expect("parse private key");
    let cert = X509::from_pem(TEST_CERT_PEM.as_bytes()).expect("parse certificate");
    let extra_certs = Stack::new().expect("create cert stack");

    let pkcs7 = Pkcs7::sign(
        &cert,
        &pkey,
        &extra_certs,
        &spc_indirect_data_content,
        Pkcs7Flags::BINARY | Pkcs7Flags::NOATTR,
    )
    .expect("pkcs7 sign");
    let pkcs7_der = pkcs7.to_der().expect("pkcs7 DER");

    // ---- 4) Store signature in a separate signature OLE container (like `vbaProjectSignature.bin`). ----
    let signature_container_bin =
        build_vba_project_bin_with_streams(&[("\u{0005}DigitalSignature", pkcs7_der.as_slice())]);

    // ---- 5) Verify binding ----
    let binding = verify_vba_project_signature_binding(&unsigned_vba_project_bin, &signature_container_bin)
        .expect("binding verification should succeed");

    let debug = match binding {
        VbaProjectBindingVerification::BoundVerified(debug) => debug,
        other => panic!("expected BoundVerified, got {:?}", other),
    };

    assert_eq!(
        debug.hash_algorithm_oid.as_deref(),
        Some("2.16.840.1.101.3.4.2.1")
    );
    assert_eq!(debug.hash_algorithm_name.as_deref(), Some("SHA-256"));
    assert_eq!(debug.signed_digest.as_deref(), Some(project_md5.as_ref()));
    assert_eq!(debug.computed_digest.as_deref(), Some(project_md5.as_ref()));
}

fn build_vba_project_bin_with_streams(streams: &[(&str, &[u8])]) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create compound file");
    ole.create_storage("VBA").expect("create VBA storage");

    for (path, bytes) in streams {
        let mut stream = ole.create_stream(path).expect("create stream");
        stream.write_all(bytes).expect("write bytes");
    }

    ole.into_inner().into_inner()
}

fn push_record(out: &mut Vec<u8>, id: u16, data: &[u8]) {
    out.extend_from_slice(&id.to_le_bytes());
    out.extend_from_slice(&(data.len() as u32).to_le_bytes());
    out.extend_from_slice(data);
}

fn build_spc_indirect_data_content_sha256_oid_with_md5_digest(md5_digest: &[u8]) -> Vec<u8> {
    // AlgorithmIdentifier ::= SEQUENCE { algorithm OBJECT IDENTIFIER, parameters NULL }
    let sha256_oid = der_oid(&[0x60, 0x86, 0x48, 0x01, 0x65, 0x03, 0x04, 0x02, 0x01]);
    let alg_id = der_sequence(&[sha256_oid, der_null()]);

    // DigestInfo ::= SEQUENCE { digestAlgorithm AlgorithmIdentifier, digest OCTET STRING }
    let digest_info = der_sequence(&[alg_id, der_octet_string(md5_digest)]);

    // SpcAttributeTypeAndOptionalValue ::= SEQUENCE { type OID, value [0] EXPLICIT ANY OPTIONAL }
    // For VBA signatures the precise `type` value is not relevant to this regression; we only
    // care about `messageDigest`.
    let dummy_type_oid = der_oid(&[
        0x2b, 0x06, 0x01, 0x04, 0x01, 0x82, 0x37, 0x02, 0x01, 0x1e,
    ]); // 1.3.6.1.4.1.311.2.1.30 (SpcSipInfo)
    let data = der_sequence(&[dummy_type_oid]);

    // SpcIndirectDataContent ::= SEQUENCE { data, messageDigest }
    der_sequence(&[data, digest_info])
}

fn build_spc_indirect_data_content_v2_with_sha256_algorithm_and_md5_source_hash(
    md5_digest: &[u8],
) -> Vec<u8> {
    // [MS-OSHARED] SpcIndirectDataContentV2 stores the VBA project digest as a 16-byte MD5
    // `SigDataV1Serialized.sourceHash` even when the signature advertises a different algorithm
    // (e.g. SHA-256).
    //
    // This constructs a minimal V2-like payload that:
    // - includes a SHA-256 AlgorithmIdentifier (best-effort simulation of `algorithmId`)
    // - includes a 16-byte OCTET STRING holding the MD5 project hash (simulating `sourceHash`)
    //
    // The "SigData" ASN.1 is wrapped in an OCTET STRING so the classic Authenticode
    // SpcIndirectDataContent parser (which expects a DigestInfo SEQUENCE) fails and our V2 parser
    // path is exercised.

    // AlgorithmIdentifier ::= SEQUENCE { algorithm OBJECT IDENTIFIER, parameters NULL }
    let sha256_oid = der_oid(&[0x60, 0x86, 0x48, 0x01, 0x65, 0x03, 0x04, 0x02, 0x01]);
    let alg_id = der_sequence(&[sha256_oid, der_null()]);

    // SigDataV1Serialized (ASN.1-ish): SEQUENCE { version INTEGER, algorithmId AlgorithmIdentifier, sourceHash OCTET STRING }
    // The production structure includes a small version field; include it so our strict V2 parser
    // can identify the blob without accidentally treating arbitrary 16-byte OCTET STRINGs as the
    // VBA project hash.
    let version = vec![0x02, 0x01, 0x01]; // INTEGER 1
    let sig_data = der_sequence(&[version, alg_id, der_octet_string(md5_digest)]);

    // SpcIndirectDataContentV2-like: SEQUENCE { data ANY, sigData OCTET STRING }
    der_sequence(&[der_null(), der_octet_string(&sig_data)])
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

fn der_tlv(tag: u8, content: &[u8]) -> Vec<u8> {
    let mut out = Vec::new();
    out.push(tag);
    out.extend_from_slice(&der_len(content.len()));
    out.extend_from_slice(content);
    out
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
    let mut out = Vec::new();
    let _ = out.try_reserve_exact(1usize.saturating_add(buf.len()));
    out.push(0x80 | (buf.len() as u8));
    out.extend_from_slice(&buf);
    out
}
