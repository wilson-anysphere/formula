use thiserror::Error;

/// Digest extracted from the signed Authenticode `SpcIndirectDataContent`.
///
/// In MS-OVBA terms, this corresponds to the "project digest" binding value
/// stored inside the VBA digital signature stream.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct VbaSignedDigest {
    /// Digest algorithm OID (e.g. SHA1 `1.3.14.3.2.26`, SHA256 `2.16.840.1.101.3.4.2.1`).
    pub digest_algorithm_oid: String,
    /// Digest bytes.
    pub digest: Vec<u8>,
}

#[derive(Debug, Error)]
pub enum VbaSignatureSignedDigestError {
    #[error("ASN.1 parse error: {0}")]
    Der(String),
    #[error("PKCS#7 SignedData is detached, but no detached content was found")]
    DetachedContentMissing,
}

// PKCS#7 OID value-encoding bytes (no tag/length)
const OID_PKCS7_SIGNED_DATA: &[u8] = b"\x2A\x86\x48\x86\xF7\x0D\x01\x07\x02"; // 1.2.840.113549.1.7.2

/// Extract the signed Authenticode file digest (the `DigestInfo` inside
/// `SpcIndirectDataContent`) from a raw VBA `\x05DigitalSignature*` stream.
///
/// This is a best-effort parser intended for binding verification (MS-OVBA "project digest").
///
/// Returns:
/// - `Ok(Some(_))` if a PKCS#7/CMS SignedData blob and `SpcIndirectDataContent` were found and parsed.
/// - `Ok(None)` if no PKCS#7 SignedData could be located in the stream.
///
/// Notes:
/// - Supports both strict DER and BER with indefinite-length encodings (OpenSSL `cms -stream`).
/// - Handles detached signatures by treating any stream prefix (before the CMS blob) as the
///   detached content.
pub fn extract_vba_signature_signed_digest(
    signature_stream: &[u8],
) -> Result<Option<VbaSignedDigest>, VbaSignatureSignedDigestError> {
    let mut candidates = Vec::new();

    // Prefer a deterministic MS-OFFCRYPTO DigSigInfoSerialized location when present.
    if let Some(info) = crate::offcrypto::parse_digsig_info_serialized(signature_stream) {
        let end = info.pkcs7_offset.saturating_add(info.pkcs7_len);
        if end <= signature_stream.len() {
            candidates.push(Pkcs7Location {
                der: &signature_stream[info.pkcs7_offset..end],
                offset: info.pkcs7_offset,
            });
        }
    }

    // Fast path: raw ContentInfo at the start.
    if signature_stream.first() == Some(&0x30)
        && looks_like_pkcs7_signed_data_content_info(signature_stream)
        && !candidates.iter().any(|c| c.offset == 0)
    {
        candidates.push(Pkcs7Location {
            der: signature_stream,
            offset: 0,
        });
    }

    // Fallback: scan for embedded SignedData ContentInfo sequences. This is best-effort: signature
    // streams can contain *multiple* SignedData blobs (e.g. certificate stores + signature), so we
    // keep searching until we find one whose signed content parses as Authenticode
    // `SpcIndirectDataContent`.
    for offset in 0..signature_stream.len() {
        if signature_stream[offset] != 0x30 {
            continue;
        }
        if candidates.iter().any(|c| c.offset == offset) {
            continue;
        }
        if looks_like_pkcs7_signed_data_content_info(&signature_stream[offset..]) {
            candidates.push(Pkcs7Location {
                der: &signature_stream[offset..],
                offset,
            });
        }
    }

    if candidates.is_empty() {
        return Ok(None);
    }

    let mut last_err = None;
    for pkcs7 in candidates {
        match extract_signed_digest_from_pkcs7_location(signature_stream, pkcs7) {
            Ok(digest) => return Ok(Some(digest)),
            Err(err) => last_err = Some(err),
        }
    }

    Err(last_err.unwrap_or_else(|| VbaSignatureSignedDigestError::Der(
        "no SpcIndirectDataContent digest found in PKCS#7 SignedData candidates".to_owned(),
    )))
}

/// Locate the BER/DER-encoded PKCS#7/CMS SignedData `ContentInfo` inside a VBA signature stream.
///
/// Returns `(offset, len)` where `offset` is the byte offset from the start of the stream and `len`
/// is the total length of the ASN.1 TLV (including tag/length/EOC for indefinite encodings).
#[cfg(not(target_arch = "wasm32"))]
pub(crate) fn locate_pkcs7_signed_data_bounds(signature_stream: &[u8]) -> Option<(usize, usize)> {
    if let Some(info) = crate::offcrypto::parse_digsig_info_serialized(signature_stream) {
        return Some((info.pkcs7_offset, info.pkcs7_len));
    }

    // When scanning, prefer the *last* plausible SignedData ContentInfo in the stream.
    //
    // Real-world VBA signature streams can contain multiple PKCS#7 blobs (notably a PKCS#7
    // certificate store followed by the actual signature). The signature payload typically comes
    // last, so selecting the final candidate avoids treating the cert store as the signature.
    let mut best: Option<(usize, usize)> = None;
    for offset in 0..signature_stream.len() {
        if signature_stream[offset] != 0x30 {
            continue;
        }
        let slice = &signature_stream[offset..];
        if looks_like_pkcs7_signed_data_content_info(slice) {
            let rem = skip_element(slice).ok()?;
            let len = slice.len().saturating_sub(rem.len());
            best = Some((offset, len));
        }
    }

    best
}

#[derive(Debug, Clone, Copy)]
struct Pkcs7Location<'a> {
    der: &'a [u8],
    offset: usize,
}

#[derive(Debug, Clone)]
struct Pkcs7EncapsulatedContent {
    #[allow(dead_code)]
    econtent_type_oid: String,
    econtent: Option<Vec<u8>>,
}

fn extract_signed_digest_from_pkcs7_location(
    signature_stream: &[u8],
    pkcs7: Pkcs7Location<'_>,
) -> Result<VbaSignedDigest, VbaSignatureSignedDigestError> {
    // `pkcs7.der` may include trailing bytes (e.g. when scanning through a larger signature stream).
    // Trim it to exactly one ASN.1 element so BER indefinite-length encodings and appended data don't
    // interfere with parsing.
    let der = pkcs7.der;
    let consumed = der.len().saturating_sub(skip_element(der)?.len());
    let der = der.get(..consumed).unwrap_or(der);

    let encap = parse_pkcs7_signed_data_encap_content(der)?;

    let signed_content = if let Some(econtent) = encap.econtent {
        econtent
    } else if pkcs7.offset > 0 {
        signature_stream[..pkcs7.offset].to_vec()
    } else {
        return Err(VbaSignatureSignedDigestError::DetachedContentMissing);
    };

    parse_spc_indirect_data_content(&signed_content)
}

fn looks_like_pkcs7_signed_data_content_info(bytes: &[u8]) -> bool {
    // ContentInfo ::= SEQUENCE { contentType OID, content [0] EXPLICIT ANY OPTIONAL }
    let Ok((tag, len, rest)) = parse_tag_and_length(bytes) else {
        return false;
    };
    if tag.class != Asn1Class::Universal || !tag.constructed || tag.number != 16 {
        return false;
    }

    let Ok(content) = slice_constructed_contents(rest, len) else {
        return false;
    };
    let mut cur = content;

    let Ok((oid, after_oid)) = parse_oid(cur) else {
        return false;
    };
    if oid != OID_PKCS7_SIGNED_DATA {
        return false;
    }
    cur = after_oid;

    // ContentInfo.content is [0] EXPLICIT for SignedData.
    let Ok((tag2, _len2, rest2)) = parse_tag_and_length(cur) else {
        return false;
    };
    if tag2.class != Asn1Class::ContextSpecific || !tag2.constructed || tag2.number != 0 {
        return false;
    }

    // Sanity check that the explicit content begins with SignedData ::= SEQUENCE { version INTEGER,
    // digestAlgorithms SET, encapContentInfo SEQUENCE, ... }.
    let Ok(signed_data_wrapper) = slice_constructed_contents(rest2, _len2) else {
        return false;
    };
    let Ok((sd_tag, sd_len, sd_rest)) = parse_tag_and_length(signed_data_wrapper) else {
        return false;
    };
    if sd_tag.class != Asn1Class::Universal || !sd_tag.constructed || sd_tag.number != 16 {
        return false;
    }
    let Ok(sd_content) = slice_constructed_contents(sd_rest, sd_len) else {
        return false;
    };
    let mut sd_cur = sd_content;

    // version INTEGER
    let Ok((ver_tag, _ver_len, _ver_rest)) = parse_tag_and_length(sd_cur) else {
        return false;
    };
    if ver_tag.class != Asn1Class::Universal || ver_tag.constructed || ver_tag.number != 2 {
        return false;
    }
    sd_cur = match skip_element(sd_cur) {
        Ok(v) => v,
        Err(_) => return false,
    };

    // digestAlgorithms SET
    let Ok((dig_tag, _dig_len, _dig_rest)) = parse_tag_and_length(sd_cur) else {
        return false;
    };
    if dig_tag.class != Asn1Class::Universal || !dig_tag.constructed || dig_tag.number != 17 {
        return false;
    }
    sd_cur = match skip_element(sd_cur) {
        Ok(v) => v,
        Err(_) => return false,
    };

    // encapContentInfo SEQUENCE
    let Ok((enc_tag, _enc_len, _enc_rest)) = parse_tag_and_length(sd_cur) else {
        return false;
    };
    enc_tag.class == Asn1Class::Universal && enc_tag.constructed && enc_tag.number == 16
}

fn parse_pkcs7_signed_data_encap_content(
    pkcs7_bytes: &[u8],
) -> Result<Pkcs7EncapsulatedContent, VbaSignatureSignedDigestError> {
    // ContentInfo
    let (tag, len, rest) = parse_tag_and_length(pkcs7_bytes)?;
    if tag.class != Asn1Class::Universal || !tag.constructed || tag.number != 16 {
        return Err(der_err("expected ContentInfo SEQUENCE"));
    }
    let content = slice_constructed_contents(rest, len)?;
    let mut cur = content;

    let (content_type, after_oid) = parse_oid(cur)?;
    if content_type != OID_PKCS7_SIGNED_DATA {
        return Err(der_err(format!(
            "expected PKCS#7 signedData ContentInfo ({}), got {}",
            "1.2.840.113549.1.7.2",
            oid_to_string(content_type).unwrap_or_else(|| "<invalid-oid>".to_string())
        )));
    }
    cur = after_oid;

    // ContentInfo.content [0] EXPLICIT
    let signed_data_wrapper = parse_context_specific_constructed(cur, 0)?;

    // SignedData
    let (tag, len, rest) = parse_tag_and_length(signed_data_wrapper)?;
    if tag.class != Asn1Class::Universal || !tag.constructed || tag.number != 16 {
        return Err(der_err("expected SignedData SEQUENCE"));
    }
    let sd_content = slice_constructed_contents(rest, len)?;
    let mut sd_cur = sd_content;

    // version INTEGER
    sd_cur = skip_element(sd_cur)?;
    // digestAlgorithms SET OF AlgorithmIdentifier
    sd_cur = skip_element(sd_cur)?;

    // encapContentInfo
    let (tag, len, rest) = parse_tag_and_length(sd_cur)?;
    if tag.class != Asn1Class::Universal || !tag.constructed || tag.number != 16 {
        return Err(der_err("expected EncapsulatedContentInfo SEQUENCE"));
    }
    let encap_content = slice_constructed_contents(rest, len)?;
    let mut encap_cur = encap_content;

    let (econtent_type, after_encap_oid) = parse_oid(encap_cur)?;
    let econtent_type_oid = oid_to_string(econtent_type).unwrap_or_else(|| "<invalid-oid>".to_string());
    encap_cur = after_encap_oid;

    // eContent [0] EXPLICIT OCTET STRING OPTIONAL
    let econtent = if encap_cur.is_empty() || is_eoc(encap_cur) {
        None
    } else {
        let (tag, len, rest) = parse_tag_and_length(encap_cur)?;
        if tag.class != Asn1Class::ContextSpecific || !tag.constructed || tag.number != 0 {
            return Err(der_err(format!(
                "unexpected EncapsulatedContentInfo field tag class={:?} constructed={} number={}",
                tag.class, tag.constructed, tag.number
            )));
        }
        let wrapper_content = slice_constructed_contents(rest, len)?;
        let (octets, _after_octets) = parse_octet_string(wrapper_content)?;
        Some(octets)
    };

    Ok(Pkcs7EncapsulatedContent {
        econtent_type_oid,
        econtent,
    })
}

fn parse_spc_indirect_data_content(
    bytes: &[u8],
) -> Result<VbaSignedDigest, VbaSignatureSignedDigestError> {
    // SpcIndirectDataContent ::= SEQUENCE { data, messageDigest DigestInfo }
    let (tag, len, rest) = parse_tag_and_length(bytes)?;
    if tag.class != Asn1Class::Universal || !tag.constructed || tag.number != 16 {
        return Err(der_err("expected SpcIndirectDataContent SEQUENCE"));
    }
    let content = slice_constructed_contents(rest, len)?;
    let mut cur = content;

    // data (ignored)
    cur = skip_element(cur)?;

    // DigestInfo
    let (tag, len, rest) = parse_tag_and_length(cur)?;
    if tag.class != Asn1Class::Universal || !tag.constructed || tag.number != 16 {
        return Err(der_err("expected DigestInfo SEQUENCE"));
    }
    let digest_info = slice_constructed_contents(rest, len)?;
    let mut di_cur = digest_info;

    // digestAlgorithm AlgorithmIdentifier
    let (tag, len, rest) = parse_tag_and_length(di_cur)?;
    if tag.class != Asn1Class::Universal || !tag.constructed || tag.number != 16 {
        return Err(der_err("expected AlgorithmIdentifier SEQUENCE"));
    }
    let alg_content = slice_constructed_contents(rest, len)?;
    let (alg_oid, _after_alg_oid) = parse_oid(alg_content)?;
    let digest_algorithm_oid = oid_to_string(alg_oid).unwrap_or_else(|| "<invalid-oid>".to_string());

    // Skip over AlgorithmIdentifier to reach digest OCTET STRING.
    di_cur = skip_element(di_cur)?;

    let (digest, _after_digest) = parse_octet_string(di_cur)?;

    Ok(VbaSignedDigest {
        digest_algorithm_oid,
        digest,
    })
}

fn is_eoc(bytes: &[u8]) -> bool {
    bytes.len() >= 2 && bytes[0] == 0x00 && bytes[1] == 0x00
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Asn1Class {
    Universal,
    Application,
    ContextSpecific,
    Private,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct Tag {
    class: Asn1Class,
    constructed: bool,
    number: u32,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Length {
    Definite(usize),
    Indefinite,
}

fn parse_tag_and_length(
    input: &[u8],
) -> Result<(Tag, Length, &[u8]), VbaSignatureSignedDigestError> {
    let (tag, tag_len) = parse_tag(input)?;
    let (len, len_len) = parse_length(
        input
            .get(tag_len..)
            .ok_or_else(|| der_err("unexpected EOF"))?,
    )?;
    let header_len = tag_len + len_len;
    let rest = input
        .get(header_len..)
        .ok_or_else(|| der_err("unexpected EOF"))?;
    Ok((tag, len, rest))
}

fn parse_tag(input: &[u8]) -> Result<(Tag, usize), VbaSignatureSignedDigestError> {
    let b0 = *input.first().ok_or_else(|| der_err("unexpected EOF"))?;

    let class = match b0 >> 6 {
        0 => Asn1Class::Universal,
        1 => Asn1Class::Application,
        2 => Asn1Class::ContextSpecific,
        3 => Asn1Class::Private,
        _ => return Err(der_err("invalid tag class")),
    };
    let constructed = b0 & 0x20 != 0;
    let mut number: u32 = (b0 & 0x1F) as u32;
    let mut idx = 1;

    if number == 0x1F {
        // High-tag-number form (base-128).
        number = 0;
        loop {
            let b = *input.get(idx).ok_or_else(|| der_err("unexpected EOF"))?;
            idx += 1;
            let v = (b & 0x7F) as u32;
            number = number
                .checked_shl(7)
                .ok_or_else(|| der_err("tag number overflow"))?;
            number |= v;
            if b & 0x80 == 0 {
                break;
            }
            if idx > 6 {
                return Err(der_err("tag number too large"));
            }
        }
    }

    Ok((
        Tag {
            class,
            constructed,
            number,
        },
        idx,
    ))
}

fn parse_length(input: &[u8]) -> Result<(Length, usize), VbaSignatureSignedDigestError> {
    let b0 = *input.first().ok_or_else(|| der_err("unexpected EOF"))?;
    if b0 < 0x80 {
        return Ok((Length::Definite(b0 as usize), 1));
    }
    if b0 == 0x80 {
        return Ok((Length::Indefinite, 1));
    }

    let count = (b0 & 0x7F) as usize;
    if count == 0 || count > 8 {
        return Err(der_err("invalid length"));
    }
    if input.len() < 1 + count {
        return Err(der_err("unexpected EOF parsing length"));
    }

    let mut len: usize = 0;
    for &b in &input[1..1 + count] {
        len = len
            .checked_shl(8)
            .ok_or_else(|| der_err("length overflow"))?;
        len |= b as usize;
    }
    Ok((Length::Definite(len), 1 + count))
}

fn slice_constructed_contents<'a>(
    rest_after_header: &'a [u8],
    len: Length,
) -> Result<&'a [u8], VbaSignatureSignedDigestError> {
    match len {
        Length::Definite(l) => rest_after_header
            .get(..l)
            .ok_or_else(|| der_err("length exceeds input")),
        Length::Indefinite => Ok(rest_after_header),
    }
}

fn parse_context_specific_constructed<'a>(
    input: &'a [u8],
    tag_number: u32,
) -> Result<&'a [u8], VbaSignatureSignedDigestError> {
    let (tag, len, rest) = parse_tag_and_length(input)?;
    if tag.class != Asn1Class::ContextSpecific || !tag.constructed || tag.number != tag_number {
        return Err(der_err("unexpected tag"));
    }
    slice_constructed_contents(rest, len)
}

fn parse_oid<'a>(
    input: &'a [u8],
) -> Result<(&'a [u8], &'a [u8]), VbaSignatureSignedDigestError> {
    let (tag, len, rest) = parse_tag_and_length(input)?;
    if tag.class != Asn1Class::Universal || tag.constructed || tag.number != 6 {
        return Err(der_err("expected OBJECT IDENTIFIER"));
    }
    let Length::Definite(l) = len else {
        return Err(der_err("OID uses indefinite length"));
    };
    let val = rest
        .get(..l)
        .ok_or_else(|| der_err("OID length exceeds input"))?;
    let after = rest
        .get(l..)
        .ok_or_else(|| der_err("unexpected EOF"))?;
    Ok((val, after))
}

fn parse_octet_string(
    input: &[u8],
) -> Result<(Vec<u8>, &[u8]), VbaSignatureSignedDigestError> {
    let (tag, len, rest) = parse_tag_and_length(input)?;
    if tag.class != Asn1Class::Universal || tag.number != 4 {
        return Err(der_err("expected OCTET STRING"));
    }

    if !tag.constructed {
        let Length::Definite(l) = len else {
            return Err(der_err("primitive OCTET STRING uses indefinite length"));
        };
        let val = rest
            .get(..l)
            .ok_or_else(|| der_err("OCTET STRING length exceeds input"))?;
        let after = rest
            .get(l..)
            .ok_or_else(|| der_err("unexpected EOF"))?;
        Ok((val.to_vec(), after))
    } else {
        // BER constructed OCTET STRING: concatenate child OCTET STRING values.
        let content = slice_constructed_contents(rest, len)?;
        let mut cur = content;
        let mut out = Vec::new();

        loop {
            if cur.is_empty() {
                break;
            }
            if cur.len() >= 2 && cur[0] == 0x00 && cur[1] == 0x00 {
                // End-of-contents for indefinite length.
                break;
            }
            let (seg, rest2) = parse_octet_string(cur)?;
            out.extend_from_slice(&seg);
            cur = rest2;
        }

        let after = skip_element(input)?;
        Ok((out, after))
    }
}

fn skip_element(input: &[u8]) -> Result<&[u8], VbaSignatureSignedDigestError> {
    let (tag, len, rest) = parse_tag_and_length(input)?;
    match len {
        Length::Definite(l) => rest
            .get(l..)
            .ok_or_else(|| der_err("length exceeds input")),
        Length::Indefinite => {
            if !tag.constructed {
                return Err(der_err("indefinite length used with primitive tag"));
            }
            let mut cur = rest;
            loop {
                if cur.len() < 2 {
                    return Err(der_err("unexpected EOF"));
                }
                if cur[0] == 0x00 && cur[1] == 0x00 {
                    return Ok(&cur[2..]);
                }
                cur = skip_element(cur)?;
            }
        }
    }
}

fn oid_to_string(oid: &[u8]) -> Option<String> {
    if oid.is_empty() {
        return None;
    }

    let first = oid[0];
    let (a, b) = if first < 40 {
        (0u32, first as u32)
    } else if first < 80 {
        (1u32, (first - 40) as u32)
    } else {
        (2u32, (first - 80) as u32)
    };

    let mut parts = vec![a.to_string(), b.to_string()];
    let mut cur: u64 = 0;
    let mut in_arc = false;

    for &byte in &oid[1..] {
        in_arc = true;
        cur = (cur << 7) | u64::from(byte & 0x7F);
        if byte & 0x80 == 0 {
            parts.push(cur.to_string());
            cur = 0;
            in_arc = false;
        }
    }

    if in_arc {
        return None;
    }

    Some(parts.join("."))
}

fn der_err(msg: impl Into<String>) -> VbaSignatureSignedDigestError {
    VbaSignatureSignedDigestError::Der(msg.into())
}
