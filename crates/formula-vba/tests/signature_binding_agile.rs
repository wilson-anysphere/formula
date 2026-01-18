#![cfg(not(target_arch = "wasm32"))]

use std::io::{Cursor, Write};

use formula_vba::{
    compress_container, content_normalized_data, forms_normalized_data, verify_vba_digital_signature,
    VbaSignatureBinding, VbaSignatureVerification,
};
use md5::{Digest as _, Md5};

mod signature_test_utils;

use signature_test_utils::make_pkcs7_signed_message;

fn push_record(out: &mut Vec<u8>, id: u16, data: &[u8]) {
    out.extend_from_slice(&id.to_le_bytes());
    out.extend_from_slice(&(data.len() as u32).to_le_bytes());
    out.extend_from_slice(data);
}

fn der_len(len: usize) -> Vec<u8> {
    if len < 0x80 {
        return vec![len as u8];
    }
    let mut tmp = Vec::new();
    let mut n = len;
    while n > 0 {
        tmp.push((n & 0xFF) as u8);
        n >>= 8;
    }
    tmp.reverse();
    let mut out = Vec::new();
    let _ = out.try_reserve_exact(1usize.saturating_add(tmp.len()));
    out.push(0x80 | (tmp.len() as u8));
    out.extend_from_slice(&tmp);
    out
}

fn der_tlv(tag: u8, value: &[u8]) -> Vec<u8> {
    let mut out = Vec::new();
    out.push(tag);
    out.extend_from_slice(&der_len(value.len()));
    out.extend_from_slice(value);
    out
}

fn der_oid(arcs: &[u32]) -> Vec<u8> {
    assert!(arcs.len() >= 2, "OID requires at least two arcs");
    let mut body = Vec::new();
    let first = arcs[0] * 40 + arcs[1];
    body.push(first as u8);
    for &arc in &arcs[2..] {
        let mut stack = Vec::new();
        let mut v = arc;
        stack.push((v & 0x7F) as u8);
        v >>= 7;
        while v > 0 {
            stack.push(((v & 0x7F) as u8) | 0x80);
            v >>= 7;
        }
        stack.reverse();
        body.extend_from_slice(&stack);
    }
    der_tlv(0x06, &body)
}

fn build_spc_indirect_data_content(message_digest_md5: [u8; 16]) -> Vec<u8> {
    // SpcIndirectDataContent ::= SEQUENCE {
    //   data          SpcAttributeTypeAndOptionalValue,
    //   messageDigest DigestInfo
    // }
    // DigestInfo ::= SEQUENCE {
    //   digestAlgorithm AlgorithmIdentifier,
    //   digest          OCTET STRING
    // }
    // AlgorithmIdentifier ::= SEQUENCE { algorithm OBJECT IDENTIFIER, parameters NULL }
    //
    // We keep the `data` field minimal: SEQUENCE { type OID }, omitting the optional value.

    // type = 1.3.6.1.4.1.311.2.1.15 (SPC_PE_IMAGE_DATA_OBJID)
    let data = der_tlv(0x30, &der_oid(&[1, 3, 6, 1, 4, 1, 311, 2, 1, 15]));

    // digestAlgorithm = md5 (1.2.840.113549.2.5) + NULL params
    let mut alg_id = Vec::new();
    alg_id.extend_from_slice(&der_oid(&[1, 2, 840, 113549, 2, 5]));
    alg_id.extend_from_slice(&[0x05, 0x00]); // NULL
    let alg_id = der_tlv(0x30, &alg_id);

    let digest_octet = der_tlv(0x04, &message_digest_md5);
    let mut digest_info_body = Vec::new();
    digest_info_body.extend_from_slice(&alg_id);
    digest_info_body.extend_from_slice(&digest_octet);
    let digest_info = der_tlv(0x30, &digest_info_body);

    let mut outer = Vec::new();
    outer.extend_from_slice(&data);
    outer.extend_from_slice(&digest_info);
    der_tlv(0x30, &outer)
}

fn build_vba_project_bin_with_designer(designer_bytes: &[u8], signature_blob: Option<&[u8]>) -> Vec<u8> {
    // Build minimal VBA project:
    // - PROJECT stream with BaseClass=UserForm1
    // - VBA/dir with module UserForm1
    // - VBA/UserForm1 module stream (compressed)
    // - Root storage UserForm1 with a designer stream (e.g. "\x03VBFrame")
    // - Optional signature stream

    let module_code = b"Attribute VB_Name = \"UserForm1\"\r\nSub Foo()\r\nEnd Sub\r\n";
    let module_container = compress_container(module_code);

    let dir_decompressed = {
        let mut out = Vec::new();
        // PROJECTCODEPAGE (u16 LE)
        push_record(&mut out, 0x0003, &1252u16.to_le_bytes());
        // PROJECTNAME (MBCS bytes)
        push_record(&mut out, 0x0004, b"VBAProject");

        // PROJECTCONSTANTS record payload:
        // SizeOfConstants (u32) + Constants + Reserved (u16) + SizeOfConstantsUnicode (u32) + ConstantsUnicode
        let mut constants = Vec::new();
        constants.extend_from_slice(&0u32.to_le_bytes()); // SizeOfConstants
        constants.extend_from_slice(&0u16.to_le_bytes()); // Reserved
        constants.extend_from_slice(&0u32.to_le_bytes()); // SizeOfConstantsUnicode
        push_record(&mut out, 0x000C, &constants);

        // Module records.
        push_record(&mut out, 0x0019, b"UserForm1"); // MODULENAME
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"UserForm1");
        stream_name.extend_from_slice(&0u16.to_le_bytes()); // reserved u16
        push_record(&mut out, 0x001A, &stream_name); // MODULESTREAMNAME
        push_record(&mut out, 0x0021, &0x0003u16.to_le_bytes()); // MODULETYPE (UserForm)
        push_record(&mut out, 0x0031, &0u32.to_le_bytes()); // MODULETEXTOFFSET
        out
    };
    let dir_container = compress_container(&dir_decompressed);

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");

    // Root PROJECT stream (MBCS).
    {
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(b"BaseClass=UserForm1\r\n")
            .expect("write PROJECT");
    }

    // VBA storage.
    ole.create_storage("VBA").expect("VBA storage");
    {
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }
    {
        let mut s = ole
            .create_stream("VBA/UserForm1")
            .expect("module stream");
        s.write_all(&module_container).expect("write module");
    }

    // Designer storage.
    ole.create_storage("UserForm1")
        .expect("designer storage UserForm1");
    {
        let mut s = ole
            .create_stream("UserForm1/\u{0003}VBFrame")
            .expect("designer stream");
        s.write_all(designer_bytes).expect("write designer bytes");
    }

    // Signature stream: Agile Content Hash binding is associated with `DigitalSignatureEx`.
    if let Some(sig) = signature_blob {
        let mut s = ole
            .create_stream("\u{0005}DigitalSignatureEx")
            .expect("signature stream");
        s.write_all(sig).expect("write signature");
    }

    ole.into_inner().into_inner()
}

#[test]
fn agile_signature_binds_using_forms_normalized_data() {
    let designer_bytes = b"designer-stream-bytes";
    let unsigned = build_vba_project_bin_with_designer(designer_bytes, None);

    let content_normalized = content_normalized_data(&unsigned).expect("ContentNormalizedData");
    let forms_normalized = forms_normalized_data(&unsigned).expect("FormsNormalizedData");

    let mut hasher = Md5::new();
    hasher.update(&content_normalized);
    hasher.update(&forms_normalized);
    let digest: [u8; 16] = hasher.finalize().into();

    let spc = build_spc_indirect_data_content(digest);
    let pkcs7 = make_pkcs7_signed_message(&spc);
    let vba_project_bin = build_vba_project_bin_with_designer(designer_bytes, Some(&pkcs7));

    let sig = verify_vba_digital_signature(&vba_project_bin)
        .expect("verify should succeed")
        .expect("signature should be present");
    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.binding, VbaSignatureBinding::Bound);
}

#[test]
fn agile_signature_is_not_bound_if_designer_storage_changes() {
    let designer_bytes = b"designer-stream-bytes";
    let unsigned = build_vba_project_bin_with_designer(designer_bytes, None);

    let content_normalized = content_normalized_data(&unsigned).expect("ContentNormalizedData");
    let forms_normalized = forms_normalized_data(&unsigned).expect("FormsNormalizedData");

    let mut hasher = Md5::new();
    hasher.update(&content_normalized);
    hasher.update(&forms_normalized);
    let digest: [u8; 16] = hasher.finalize().into();
    let spc = build_spc_indirect_data_content(digest);
    let pkcs7 = make_pkcs7_signed_message(&spc);

    // Mutate the designer bytes after computing the signed digest.
    let mut mutated = designer_bytes.to_vec();
    mutated[0] ^= 0xFF;
    let vba_project_bin = build_vba_project_bin_with_designer(&mutated, Some(&pkcs7));

    let sig = verify_vba_digital_signature(&vba_project_bin)
        .expect("verify should succeed")
        .expect("signature should be present");
    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.binding, VbaSignatureBinding::NotBound);
}
