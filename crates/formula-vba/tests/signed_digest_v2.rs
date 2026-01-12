#![cfg(not(target_arch = "wasm32"))]

use std::io::{Cursor, Write};

use formula_vba::{extract_vba_signature_signed_digest, OleFile};

mod signature_test_utils;

use signature_test_utils::make_pkcs7_signed_message;

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

fn build_sig_data_v1_serialized_with_source_hash(source_hash: &[u8]) -> Vec<u8> {
    // Minimal binary SigDataV1Serialized-ish blob:
    // [version u32 LE] [cbSourceHash u32 LE] [sourceHash bytes]
    let mut out = Vec::new();
    out.extend_from_slice(&1u32.to_le_bytes());
    out.extend_from_slice(&(source_hash.len() as u32).to_le_bytes());
    out.extend_from_slice(source_hash);
    out
}

fn build_spc_indirect_data_content_v2(source_hash: &[u8]) -> Vec<u8> {
    // Minimal SpcIndirectDataContentV2-like payload:
    // SEQUENCE { data ANY, sigData OCTET STRING }
    //
    // The classic Authenticode `SpcIndirectDataContent` stores the digest in a `DigestInfo`
    // (SEQUENCE) as the second element. By making the second element an OCTET STRING instead, we
    // ensure the classic parser fails and our V2 parser path is exercised.
    let data = der_null();
    let sig_data = build_sig_data_v1_serialized_with_source_hash(source_hash);
    der_sequence(&[data, der_octet_string(&sig_data)])
}

#[test]
fn extracts_signed_digest_from_spc_indirect_data_content_v2_source_hash() {
    let source_hash = (0u8..16).collect::<Vec<_>>();
    let spc_v2 = build_spc_indirect_data_content_v2(&source_hash);
    let pkcs7 = make_pkcs7_signed_message(&spc_v2);

    // Store PKCS#7 in a minimal OLE file signature stream.
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut s = ole
            .create_stream("\u{0005}DigitalSignature")
            .expect("create signature stream");
        s.write_all(&pkcs7).expect("write signature");
    }
    let ole_bytes = ole.into_inner().into_inner();

    // Read it back from the OLE file and extract the digest.
    let mut ole = OleFile::open(&ole_bytes).expect("open ole");
    let signature_stream = ole
        .read_stream_opt("\u{0005}DigitalSignature")
        .expect("read signature stream")
        .expect("signature stream present");

    let got = extract_vba_signature_signed_digest(&signature_stream)
        .expect("extract digest")
        .expect("digest present");
    assert_eq!(got.digest_algorithm_oid, "1.2.840.113549.2.5");
    assert_eq!(got.digest, source_hash);
}

