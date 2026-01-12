//! Parsers for a small subset of MS-OFFCRYPTO structures used by Office VBA project signatures.
//!
//! Excel stores a signed VBA project in an OLE stream named `\x05DigitalSignature*` (see MS-OVBA).
//! In many real-world files the PKCS#7/CMS `SignedData` payload is wrapped in a
//! `[MS-OFFCRYPTO] DigSigInfoSerialized` header. The header contains size fields for the
//! surrounding metadata, making it possible to locate the DER blob deterministically instead of
//! scanning the whole stream.

/// Parsed information from a `[MS-OFFCRYPTO] DigSigInfoSerialized` prefix.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct DigSigInfoSerialized {
    /// Offset (from the start of the stream) where the DER-encoded PKCS#7 `ContentInfo` begins.
    pub(crate) pkcs7_offset: usize,
    /// Length (in bytes) of the DER-encoded PKCS#7 `ContentInfo`.
    pub(crate) pkcs7_len: usize,
    /// Best-effort version field when present.
    pub(crate) version: Option<u32>,
}

/// Best-effort parse of `[MS-OFFCRYPTO] DigSigInfoSerialized`.
///
/// Returns `None` if the stream does not look like a DigSigInfoSerialized-wrapped PKCS#7 payload.
///
/// Notes:
/// - The MS-OFFCRYPTO structure contains several length-prefixed metadata blobs (project name,
///   certificate store, etc.). The order varies across producers/versions, so we try a small set of
///   deterministic layouts and validate by checking for a well-formed PKCS#7 `SignedData`
///   `ContentInfo` at the computed offset.
/// - Integer fields are little-endian.
pub(crate) fn parse_digsig_info_serialized(stream: &[u8]) -> Option<DigSigInfoSerialized> {
    // The minimum layout we support has three u32 length fields.
    if stream.len() < 12 {
        return None;
    }

    #[derive(Debug, Clone, Copy)]
    struct Header {
        version: Option<u32>,
        header_size: usize,
        sig_len: usize,
        cert_len: usize,
        proj_len: usize,
    }

    fn read_u32_le(bytes: &[u8], offset: usize) -> Option<u32> {
        let b = bytes.get(offset..offset + 4)?;
        Some(u32::from_le_bytes([b[0], b[1], b[2], b[3]]))
    }

    // Build a small set of header candidates. MS-OFFCRYPTO uses little-endian DWORD fields.
    let mut headers = Vec::<Header>::new();

    // Common layout: [cbSignature, cbSigningCertStore, cchProjectName] then variable data.
    if let (Some(sig_len), Some(cert_len), Some(proj_len)) = (
        read_u32_le(stream, 0),
        read_u32_le(stream, 4),
        read_u32_le(stream, 8),
    ) {
        headers.push(Header {
            version: None,
            header_size: 12,
            sig_len: sig_len as usize,
            cert_len: cert_len as usize,
            proj_len: proj_len as usize,
        });
    }

    // Variant with an explicit version field first: [version, cbSignature, cbSigningCertStore, cchProjectName].
    if stream.len() >= 16 {
        if let (Some(version), Some(sig_len), Some(cert_len), Some(proj_len)) = (
            read_u32_le(stream, 0),
            read_u32_le(stream, 4),
            read_u32_le(stream, 8),
            read_u32_le(stream, 12),
        ) {
            // Reject clearly bogus versions to avoid false positives on raw DER.
            if version <= 0x100 {
                headers.push(Header {
                    version: Some(version),
                    header_size: 16,
                    sig_len: sig_len as usize,
                    cert_len: cert_len as usize,
                    proj_len: proj_len as usize,
                });
            }
        }
    }

    // Variant with version + reserved: [version, reserved, cbSignature, cbSigningCertStore, cchProjectName].
    if stream.len() >= 20 {
        if let (Some(version), Some(_reserved), Some(sig_len), Some(cert_len), Some(proj_len)) = (
            read_u32_le(stream, 0),
            read_u32_le(stream, 4),
            read_u32_le(stream, 8),
            read_u32_le(stream, 12),
            read_u32_le(stream, 16),
        ) {
            if version <= 0x100 {
                headers.push(Header {
                    version: Some(version),
                    header_size: 20,
                    sig_len: sig_len as usize,
                    cert_len: cert_len as usize,
                    proj_len: proj_len as usize,
                });
            }
        }
    }

    // Try each header candidate and a small set of deterministic layouts.
    //
    // We don't attempt to parse all metadata; we only use the length fields to derive candidate
    // offsets for the PKCS#7 `ContentInfo` and validate the computed slice.
    let mut best: Option<(usize /* padding */, DigSigInfoSerialized)> = None;
    for header in headers {
        // Basic sanity checks on the length fields.
        if header.sig_len == 0 || header.header_size > stream.len() {
            continue;
        }
        if header.sig_len > stream.len().saturating_sub(header.header_size) {
            // Signature can't start before we know where it is, but this rules out obviously bogus
            // size fields.
            // We'll still allow it if signature is not directly after the header; compute later.
        }

        // Project name length can be either bytes or WCHAR count; try both interpretations.
        for proj_bytes in [
            Some(header.proj_len),
            header.proj_len.checked_mul(2),
        ]
        .into_iter()
        .flatten()
        {
            // Total size of all variable blobs must fit inside the stream for any ordering.
            let total_min = match header
                .header_size
                .checked_add(header.sig_len)
                .and_then(|n| n.checked_add(header.cert_len))
                .and_then(|n| n.checked_add(proj_bytes))
            {
                Some(n) => n,
                None => continue,
            };
            if total_min > stream.len() {
                continue;
            }

            // The signature can appear at a small number of offsets depending on the ordering of
            // the (project name, cert store, signature) blobs.
            let candidate_offsets = [
                header.header_size,                                      // sig first
                header.header_size.saturating_add(header.cert_len),      // cert then sig
                header.header_size.saturating_add(proj_bytes),           // project then sig
                header
                    .header_size
                    .saturating_add(header.cert_len)
                    .saturating_add(proj_bytes), // project+cert then sig (or cert+project then sig)
            ];

            for &pkcs7_offset in &candidate_offsets {
                let sig_end = match pkcs7_offset.checked_add(header.sig_len) {
                    Some(end) => end,
                    None => continue,
                };
                if sig_end > stream.len() {
                    continue;
                }

                let sig_slice = &stream[pkcs7_offset..sig_end];
                let Some(pkcs7_len) = ber_total_len(sig_slice) else {
                    continue;
                };
                if pkcs7_len == 0 || pkcs7_len > sig_slice.len() {
                    continue;
                }

                // Ensure the candidate is plausibly a PKCS#7 SignedData `ContentInfo`:
                // SEQUENCE { OID signedData, [0] EXPLICIT ... }.
                if !looks_like_pkcs7_signed_data(sig_slice) {
                    continue;
                }

                let padding = header.sig_len.saturating_sub(pkcs7_len);
                let info = DigSigInfoSerialized {
                    pkcs7_offset,
                    pkcs7_len,
                    version: header.version,
                };

                match best {
                    Some((best_padding, _)) if best_padding <= padding => {}
                    _ => best = Some((padding, info)),
                }
            }
        }
    }

    best.map(|(_, info)| info)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum BerLen {
    Definite(usize),
    Indefinite,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct BerTag {
    tag_byte: u8,
    constructed: bool,
}

fn ber_parse_tag(bytes: &[u8]) -> Option<(BerTag, usize /* tag len */)> {
    let b0 = *bytes.first()?;
    let constructed = b0 & 0x20 != 0;
    let mut idx = 1;

    // High-tag-number form: base-128 tag number across subsequent bytes.
    if b0 & 0x1F == 0x1F {
        // Limit tag length to avoid pathological inputs; CMS structures should never hit this.
        loop {
            let b = *bytes.get(idx)?;
            idx += 1;
            if b & 0x80 == 0 {
                break;
            }
            if idx > 6 {
                return None;
            }
        }
    }

    Some((BerTag { tag_byte: b0, constructed }, idx))
}

fn ber_parse_len(bytes: &[u8]) -> Option<(BerLen, usize /* len len */)> {
    let b0 = *bytes.first()?;
    if b0 < 0x80 {
        return Some((BerLen::Definite(b0 as usize), 1));
    }
    if b0 == 0x80 {
        return Some((BerLen::Indefinite, 1));
    }

    let n = (b0 & 0x7F) as usize;
    if n == 0 || n > 8 {
        return None;
    }
    if bytes.len() < 1 + n {
        return None;
    }
    let mut len: usize = 0;
    for &b in &bytes[1..1 + n] {
        len = len.checked_shl(8)?;
        len = len.checked_add(b as usize)?;
    }
    Some((BerLen::Definite(len), 1 + n))
}

fn ber_header(bytes: &[u8]) -> Option<(BerTag, BerLen, usize /* header len */)> {
    let (tag, tag_len) = ber_parse_tag(bytes)?;
    let (len, len_len) = ber_parse_len(bytes.get(tag_len..)?)?;
    Some((tag, len, tag_len + len_len))
}

fn ber_skip_any(bytes: &[u8]) -> Option<&[u8]> {
    let (tag, len, hdr_len) = ber_header(bytes)?;
    let rest = bytes.get(hdr_len..)?;
    match len {
        BerLen::Definite(l) => rest.get(l..),
        BerLen::Indefinite => {
            if !tag.constructed {
                return None;
            }
            let mut cur = rest;
            loop {
                if cur.len() < 2 {
                    return None;
                }
                if cur[0] == 0x00 && cur[1] == 0x00 {
                    return cur.get(2..);
                }
                cur = ber_skip_any(cur)?;
            }
        }
    }
}

fn ber_total_len(bytes: &[u8]) -> Option<usize> {
    let rem = ber_skip_any(bytes)?;
    Some(bytes.len().saturating_sub(rem.len()))
}

fn looks_like_pkcs7_signed_data(bytes: &[u8]) -> bool {
    // ContentInfo ::= SEQUENCE { contentType OID, content [0] EXPLICIT ANY OPTIONAL }
    // For SignedData, contentType == 1.2.840.113549.1.7.2
    const SIGNED_DATA_OID: &[u8] = b"\x2A\x86\x48\x86\xF7\x0D\x01\x07\x02";

    let (tag, _len, hdr_len) = match ber_header(bytes) {
        Some(v) => v,
        None => return false,
    };
    if tag.tag_byte != 0x30 {
        return false;
    }

    let rest = &bytes[hdr_len..];
    let (tag2, oid_len, hdr2_len) = match ber_header(rest) {
        Some(v) => v,
        None => return false,
    };
    if tag2.tag_byte != 0x06 {
        return false;
    }
    let BerLen::Definite(oid_len) = oid_len else {
        return false;
    };
    let oid_bytes = match rest.get(hdr2_len..hdr2_len + oid_len) {
        Some(v) => v,
        None => return false,
    };
    oid_bytes == SIGNED_DATA_OID
}

#[cfg(test)]
mod tests {
    use super::*;

    #[cfg(not(target_arch = "wasm32"))]
    fn make_pkcs7_signed_message(data: &[u8]) -> Vec<u8> {
        // Keep the test self-contained: generate a deterministic PKCS#7 SignedData with an
        // embedded self-signed certificate.
        use openssl::asn1::Asn1Time;
        use openssl::hash::MessageDigest;
        use openssl::pkey::PKey;
        use openssl::pkcs7::{Pkcs7, Pkcs7Flags};
        use openssl::rsa::Rsa;
        use openssl::stack::Stack;
        use openssl::x509::X509NameBuilder;
        use openssl::x509::{X509Builder, X509};

        let rsa = Rsa::generate(2048).expect("generate RSA key");
        let pkey = PKey::from_rsa(rsa).expect("build pkey");

        let mut name = X509NameBuilder::new().expect("name builder");
        name.append_entry_by_text("CN", "formula-vba-test")
            .expect("CN");
        let name = name.build();

        let mut builder = X509Builder::new().expect("x509 builder");
        builder.set_version(2).expect("version");
        builder.set_subject_name(&name).expect("subject");
        builder.set_issuer_name(&name).expect("issuer");
        builder.set_pubkey(&pkey).expect("pubkey");
        builder
            .set_not_before(&Asn1Time::days_from_now(0).unwrap())
            .unwrap();
        builder
            .set_not_after(&Asn1Time::days_from_now(1).unwrap())
            .unwrap();
        builder
            .sign(&pkey, MessageDigest::sha256())
            .expect("sign cert");
        let cert: X509 = builder.build();

        let extra_certs = Stack::new().expect("cert stack");
        let pkcs7 = Pkcs7::sign(
            &cert,
            &pkey,
            &extra_certs,
            data,
            // NOATTR keeps output deterministic enough for testing.
            Pkcs7Flags::BINARY | Pkcs7Flags::NOATTR,
        )
        .expect("pkcs7 sign");
        pkcs7.to_der().expect("pkcs7 der")
    }

    #[test]
    #[cfg(not(target_arch = "wasm32"))]
    fn parses_pkcs7_payload_from_digsig_info_serialized_wrapper() {
        use openssl::pkcs7::Pkcs7;

        let pkcs7 = make_pkcs7_signed_message(b"formula-vba-test");

        // Synthetic DigSigInfoSerialized-like stream:
        // [cbSignature, cbSigningCertStore, cchProjectName] (LE u32)
        // [projectName UTF-16LE] [certStore bytes] [signature bytes]
        let project_name_utf16: Vec<u16> = "VBAProject\0".encode_utf16().collect();
        let mut project_name_bytes = Vec::new();
        for ch in &project_name_utf16 {
            project_name_bytes.extend_from_slice(&ch.to_le_bytes());
        }
        let cert_store = vec![0xAA, 0xBB, 0xCC, 0xDD];

        let cb_signature = pkcs7.len() as u32;
        let cb_cert_store = cert_store.len() as u32;
        let cch_project = project_name_utf16.len() as u32;

        let mut stream = Vec::new();
        stream.extend_from_slice(&cb_signature.to_le_bytes());
        stream.extend_from_slice(&cb_cert_store.to_le_bytes());
        stream.extend_from_slice(&cch_project.to_le_bytes());
        stream.extend_from_slice(&project_name_bytes);
        stream.extend_from_slice(&cert_store);
        stream.extend_from_slice(&pkcs7);

        let parsed = parse_digsig_info_serialized(&stream).expect("should parse");
        let expected_offset = 12 + project_name_bytes.len() + cert_store.len();
        assert_eq!(parsed.pkcs7_offset, expected_offset);
        assert_eq!(parsed.pkcs7_len, pkcs7.len());

        let pkcs7_slice = &stream[parsed.pkcs7_offset..parsed.pkcs7_offset + parsed.pkcs7_len];
        Pkcs7::from_der(pkcs7_slice).expect("openssl should parse extracted pkcs7");
    }

    #[test]
    fn parses_ber_indefinite_pkcs7_payload_from_digsig_info_serialized_wrapper() {
        // This fixture is a CMS/PKCS#7 SignedData blob emitted by OpenSSL with `cms -stream`,
        // which uses BER indefinite-length encodings.
        let pkcs7 = include_bytes!("../tests/fixtures/cms_indefinite.der");

        // Synthetic DigSigInfoSerialized-like stream:
        // [cbSignature, cbSigningCertStore, cchProjectName] (LE u32)
        // [projectName UTF-16LE] [certStore bytes] [signature bytes]
        let project_name_utf16: Vec<u16> = "VBAProject\0".encode_utf16().collect();
        let mut project_name_bytes = Vec::new();
        for ch in &project_name_utf16 {
            project_name_bytes.extend_from_slice(&ch.to_le_bytes());
        }
        let cert_store = vec![0xAA, 0xBB, 0xCC, 0xDD];

        let cb_signature = pkcs7.len() as u32;
        let cb_cert_store = cert_store.len() as u32;
        let cch_project = project_name_utf16.len() as u32;

        let mut stream = Vec::new();
        stream.extend_from_slice(&cb_signature.to_le_bytes());
        stream.extend_from_slice(&cb_cert_store.to_le_bytes());
        stream.extend_from_slice(&cch_project.to_le_bytes());
        stream.extend_from_slice(&project_name_bytes);
        stream.extend_from_slice(&cert_store);
        stream.extend_from_slice(pkcs7);

        let parsed = parse_digsig_info_serialized(&stream).expect("should parse");
        let expected_offset = 12 + project_name_bytes.len() + cert_store.len();
        assert_eq!(parsed.pkcs7_offset, expected_offset);
        assert_eq!(parsed.pkcs7_len, pkcs7.len());
    }
}
