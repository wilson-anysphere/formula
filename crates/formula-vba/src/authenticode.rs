use thiserror::Error;

/// Digest extracted from the signed Authenticode / MS-OSHARED signature binding payload.
///
/// Office can embed the VBA signature binding digest (MS-OVBA "Contents Hash") in either:
/// - classic Authenticode `SpcIndirectDataContent` (`DigestInfo.digest`), or
/// - MS-OSHARED `SpcIndirectDataContentV2` (`SigDataV1Serialized.sourceHash`).
///
/// In MS-OVBA terms, this corresponds to the "Contents Hash" binding value stored inside the VBA
/// digital signature stream.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct VbaSignedDigest {
    /// Algorithm OID from the signed digest structure.
    ///
    /// Note: for VBA signatures this OID is not authoritative for binding:
    /// - v1/v2 (`\x05DigitalSignature` / `\x05DigitalSignatureEx`): digest bytes are always
    ///   **16-byte MD5** per MS-OSHARED §4.3 even when this OID indicates SHA-256.
    /// - v3 (`\x05DigitalSignatureExt`): digest bytes are the MS-OVBA v3 binding digest over the
    ///   v3 transcript. In the wild this is commonly a 32-byte SHA-256, but producers can vary and
    ///   some emit inconsistent OIDs.
    ///
    /// This field is surfaced for debugging/UI display only.
    ///
    /// https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oshared/40c8dab3-e8db-4c66-a6be-8cec06351b1e
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
const OID_MD5_STR: &str = "1.2.840.113549.2.5";

/// Extract the signed VBA signature binding digest from a raw VBA `\x05DigitalSignature*` stream.
///
/// This supports both:
/// - classic Authenticode `SpcIndirectDataContent` (extracts `DigestInfo`), and
/// - MS-OSHARED `SpcIndirectDataContentV2` (extracts `SigDataV1Serialized.sourceHash`).
///
/// This is a best-effort parser intended for binding verification (MS-OVBA "Contents Hash").
///
/// Returns:
/// - `Ok(Some(_))` if a PKCS#7/CMS SignedData blob was found and its signed content parsed as either
///   `SpcIndirectDataContent` or `SpcIndirectDataContentV2`.
/// - `Ok(None)` if no PKCS#7 SignedData could be located in the stream.
///
/// Notes:
/// - Supports both strict DER and BER with indefinite-length encodings (OpenSSL `cms -stream`).
/// - Handles detached signatures by treating any stream prefix (before the CMS blob) as the
///   detached content.
pub fn extract_vba_signature_signed_digest(
    signature_stream: &[u8],
) -> Result<Option<VbaSignedDigest>, VbaSignatureSignedDigestError> {
    // Track which offsets we've already attempted so we don't retry them during scanning.
    let mut attempted_offsets = [None::<usize>; 3];
    let mut attempted_count = 0usize;
    let mut any_candidate = false;
    let mut last_err = None;

    // Prefer a deterministic MS-OSHARED WordSigBlob location when present.
    if let Some(info) = crate::offcrypto::parse_wordsig_blob(signature_stream) {
        let end = info.pkcs7_offset.saturating_add(info.pkcs7_len);
        if end <= signature_stream.len() {
            any_candidate = true;
            if attempted_count < attempted_offsets.len() {
                attempted_offsets[attempted_count] = Some(info.pkcs7_offset);
                attempted_count += 1;
            }
            match extract_signed_digest_from_pkcs7_location(
                signature_stream,
                Pkcs7Location {
                    der: &signature_stream[info.pkcs7_offset..end],
                    offset: info.pkcs7_offset,
                },
            ) {
                Ok(digest) => return Ok(Some(digest)),
                Err(err) => last_err = Some(err),
            }
        }
    }

    // Prefer a deterministic MS-OSHARED DigSigBlob location when present.
    if let Some(info) = crate::offcrypto::parse_digsig_blob(signature_stream) {
        let end = info.pkcs7_offset.saturating_add(info.pkcs7_len);
        if end <= signature_stream.len() {
            any_candidate = true;
            if attempted_count < attempted_offsets.len() {
                attempted_offsets[attempted_count] = Some(info.pkcs7_offset);
                attempted_count += 1;
            }
            match extract_signed_digest_from_pkcs7_location(
                signature_stream,
                Pkcs7Location {
                    der: &signature_stream[info.pkcs7_offset..end],
                    offset: info.pkcs7_offset,
                },
            ) {
                Ok(digest) => return Ok(Some(digest)),
                Err(err) => last_err = Some(err),
            }
        }
    }

    // Prefer a deterministic DigSigInfoSerialized-like prefix location when present (the
    // length-prefixed wrapper commonly found at the start of `\x05DigitalSignature*` streams).
    if let Some(info) = crate::offcrypto::parse_digsig_info_serialized(signature_stream) {
        let end = info.pkcs7_offset.saturating_add(info.pkcs7_len);
        if end <= signature_stream.len() {
            any_candidate = true;
            if attempted_count < attempted_offsets.len() {
                attempted_offsets[attempted_count] = Some(info.pkcs7_offset);
                attempted_count += 1;
            }
            match extract_signed_digest_from_pkcs7_location(
                signature_stream,
                Pkcs7Location {
                    der: &signature_stream[info.pkcs7_offset..end],
                    offset: info.pkcs7_offset,
                },
            ) {
                Ok(digest) => return Ok(Some(digest)),
                Err(err) => last_err = Some(err),
            }
        }
    }

    // Fast path: raw ContentInfo at the start.
    if signature_stream.first() == Some(&0x30)
        && looks_like_pkcs7_signed_data_content_info(signature_stream)
        && !attempted_offsets[..attempted_count].contains(&Some(0))
    {
        any_candidate = true;
        if attempted_count < attempted_offsets.len() {
            attempted_offsets[attempted_count] = Some(0);
            attempted_count += 1;
        }
        match extract_signed_digest_from_pkcs7_location(
            signature_stream,
            Pkcs7Location {
                der: signature_stream,
                offset: 0,
            },
        ) {
            Ok(digest) => return Ok(Some(digest)),
            Err(err) => last_err = Some(err),
        }
    }

    // Fallback: scan for embedded SignedData ContentInfo sequences. This is best-effort: signature
    // streams can contain *multiple* SignedData blobs (e.g. certificate stores + signature), so we
    // keep searching until we find one whose signed content parses as an Authenticode/MS-OSHARED
    // binding payload (`SpcIndirectDataContent` or `SpcIndirectDataContentV2`).
    //
    // Prefer later candidates: the actual signature payload is typically stored last.
    for offset in (0..signature_stream.len()).rev() {
        if signature_stream[offset] != 0x30 {
            continue;
        }
        if attempted_offsets[..attempted_count].contains(&Some(offset)) {
            continue;
        }
        if !looks_like_pkcs7_signed_data_content_info(&signature_stream[offset..]) {
            continue;
        }

        any_candidate = true;
        match extract_signed_digest_from_pkcs7_location(
            signature_stream,
            Pkcs7Location {
                der: &signature_stream[offset..],
                offset,
            },
        ) {
            Ok(digest) => return Ok(Some(digest)),
            Err(err) => last_err = Some(err),
        }
    }

    if !any_candidate {
        return Ok(None);
    }
    Err(last_err.unwrap_or_else(|| {
        VbaSignatureSignedDigestError::Der(
            "no VBA signed digest found in PKCS#7 SignedData candidates".to_owned(),
        )
    }))
}

/// Locate the BER/DER-encoded PKCS#7/CMS SignedData `ContentInfo` inside a VBA signature stream.
///
/// Returns `(offset, len)` where `offset` is the byte offset from the start of the stream and `len`
/// is the total length of the ASN.1 TLV (including tag/length/EOC for indefinite encodings).
#[cfg(not(target_arch = "wasm32"))]
pub(crate) fn locate_pkcs7_signed_data_bounds(signature_stream: &[u8]) -> Option<(usize, usize)> {
    if let Some(info) = crate::offcrypto::parse_wordsig_blob(signature_stream) {
        return Some((info.pkcs7_offset, info.pkcs7_len));
    }
    if let Some(info) = crate::offcrypto::parse_digsig_blob(signature_stream) {
        return Some((info.pkcs7_offset, info.pkcs7_len));
    }
    if let Some(info) = crate::offcrypto::parse_digsig_info_serialized(signature_stream) {
        return Some((info.pkcs7_offset, info.pkcs7_len));
    }

    // When scanning, prefer the *last* plausible SignedData ContentInfo in the stream.
    //
    // Real-world VBA signature streams can contain multiple PKCS#7 blobs (notably a PKCS#7
    // certificate store followed by the actual signature). The signature payload typically comes
    // last, so selecting the final candidate avoids treating the cert store as the signature.
    let mut best: Option<(usize, usize)> = None;

    if signature_stream.first() == Some(&0x30)
        && looks_like_pkcs7_signed_data_content_info(signature_stream)
    {
        let rem = skip_element(signature_stream).ok()?;
        let len = signature_stream.len().saturating_sub(rem.len());
        best = Some((0, len));
    }
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

    match parse_spc_indirect_data_content(&signed_content) {
        Ok(v) => Ok(v),
        Err(err1) => match parse_spc_indirect_data_content_v2(&signed_content) {
            Ok(v) => Ok(v),
            Err(err2) => Err(der_err(format!(
                "failed to parse signed content as SpcIndirectDataContent ({err1}) or SpcIndirectDataContentV2 ({err2})"
            ))),
        },
    }
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
    let econtent_type_oid =
        oid_to_string(econtent_type).unwrap_or_else(|| "<invalid-oid>".to_string());
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
    let digest_algorithm_oid =
        oid_to_string(alg_oid).unwrap_or_else(|| "<invalid-oid>".to_string());

    // Skip over AlgorithmIdentifier to reach digest OCTET STRING.
    di_cur = skip_element(di_cur)?;

    let (digest_raw, _after_digest) = parse_octet_string(di_cur)?;

    // Some producers store a serialized `SigDataV1Serialized` blob inside `DigestInfo.digest`
    // instead of raw hash bytes. Detect the common cases and extract the 16-byte "source hash"
    // (MD5 per MS-OSHARED §4.3):
    // - If the digest bytes look like a self-contained DER SEQUENCE, treat it as a serialized
    //   SigData structure.
    // - Otherwise, only attempt SigData parsing when the digest length isn't a standard hash length.
    let maybe_sigdata = if digest_raw.first() == Some(&0x30) {
        matches!(skip_element(&digest_raw), Ok(rest) if rest.is_empty())
    } else {
        !matches!(digest_raw.len(), 16 | 20 | 32)
    };
    let digest = if maybe_sigdata {
        match extract_source_hash_from_sig_data_v1_serialized(&digest_raw)? {
            Some(hash) => hash,
            None => {
                // Some producers (and some legacy tests) wrap the 16-byte source hash in an ASN.1
                // SEQUENCE without the expected SigData version INTEGER. As a best-effort fallback,
                // allow extracting any 16-byte OCTET STRING contained in the DER element, but only
                // when the digest bytes look like a self-contained DER element.
                if digest_raw.first() == Some(&0x30) {
                    scan_asn1_for_octet_string_len(&digest_raw, 16)?.unwrap_or(digest_raw)
                } else {
                    digest_raw
                }
            }
        }
    } else {
        digest_raw
    };

    Ok(VbaSignedDigest {
        digest_algorithm_oid,
        digest,
    })
}

fn parse_spc_indirect_data_content_v2(
    bytes: &[u8],
) -> Result<VbaSignedDigest, VbaSignatureSignedDigestError> {
    // [MS-OSHARED] SpcIndirectDataContentV2 is a newer signature binding format used by Office.
    //
    // Unlike the classic Authenticode `SpcIndirectDataContent` (which stores the digest directly
    // in a `DigestInfo`), the VBA project hash is stored in `SigDataV1Serialized.sourceHash`.
    //
    // This parser is intentionally minimal: it traverses the ASN.1 payload looking for a
    // SigDataV1Serialized-like blob and extracts a 16-byte "source hash" (MD5 per MS-OSHARED §4.3).
    let (tag, len, rest) = parse_tag_and_length(bytes)?;
    if tag.class != Asn1Class::Universal || !tag.constructed || tag.number != 16 {
        return Err(der_err("expected SpcIndirectDataContentV2 SEQUENCE"));
    }
    let content = slice_constructed_contents(rest, len)?;

    // The spec indicates the source hash is carried in `SigDataV1Serialized.sourceHash`. Try a
    // structure-ish parse first (skip `data`, parse the next element as SigDataV1Serialized),
    // then fall back to more permissive scanning.
    if let Some(hash) = try_parse_sigdata_source_hash_from_spc_v2_contents(content)? {
        return Ok(VbaSignedDigest {
            digest_algorithm_oid: OID_MD5_STR.to_owned(),
            digest: hash,
        });
    }

    // Fallback: permissively scan for an embedded SigDataV1Serialized blob elsewhere in the
    // structure. We intentionally do **not** fall back to "any 16-byte OCTET STRING" because
    // signature binding is security-sensitive and we prefer false negatives over false positives.
    if let Some(hash) = scan_asn1_for_sigdata_source_hash(content)? {
        return Ok(VbaSignedDigest {
            digest_algorithm_oid: OID_MD5_STR.to_owned(),
            digest: hash,
        });
    }

    Err(der_err(
        "no SigDataV1Serialized.sourceHash found in SpcIndirectDataContentV2".to_owned(),
    ))
}

fn is_eoc(bytes: &[u8]) -> bool {
    bytes.len() >= 2 && bytes[0] == 0x00 && bytes[1] == 0x00
}

fn try_parse_sigdata_source_hash_from_spc_v2_contents(
    spc_v2_contents: &[u8],
) -> Result<Option<Vec<u8>>, VbaSignatureSignedDigestError> {
    // SpcIndirectDataContentV2 is expected to be:
    //   SEQUENCE { data ANY, sigData SigDataV1Serialized, ... }
    //
    // We intentionally only parse what we need:
    // - Skip the first element (`data`)
    // - Treat the next element as SigDataV1Serialized (either an OCTET STRING wrapping a binary
    //   blob, or an embedded ASN.1 element).
    let mut cur = spc_v2_contents;
    cur = skip_element(cur)?;
    if cur.is_empty() || is_eoc(cur) {
        return Ok(None);
    }

    let after = skip_element(cur)?;
    let consumed = cur.len().saturating_sub(after.len());
    let sigdata_tlv = &cur[..consumed];

    let (tag, _len, _rest) = parse_tag_and_length(sigdata_tlv)?;
    if tag.class == Asn1Class::Universal && tag.number == 4 {
        // SigDataV1Serialized stored as an OCTET STRING (common).
        let (octets, _after) = parse_octet_string(sigdata_tlv)?;
        return extract_source_hash_from_sig_data_v1_serialized(&octets);
    }

    // Otherwise, SigDataV1Serialized may be embedded as ASN.1 (e.g. a SEQUENCE) or wrapped in
    // another constructed element. Try parsing the element itself as SigData first when it's a
    // SEQUENCE, then fall back to scanning within it for an OCTET STRING that contains SigData.
    if tag.class == Asn1Class::Universal && tag.constructed && tag.number == 16 {
        if let Some(hash) = extract_source_hash_from_sig_data_v1_serialized(sigdata_tlv)? {
            return Ok(Some(hash));
        }
    }
    scan_asn1_for_sigdata_source_hash(sigdata_tlv)
}

fn scan_asn1_for_sigdata_source_hash(
    mut cur: &[u8],
) -> Result<Option<Vec<u8>>, VbaSignatureSignedDigestError> {
    while !cur.is_empty() {
        if is_eoc(cur) {
            break;
        }

        if let Some(hash) = extract_sigdata_source_hash_from_asn1_element(cur)? {
            return Ok(Some(hash));
        }
        cur = skip_element(cur)?;
    }
    Ok(None)
}

fn extract_sigdata_source_hash_from_asn1_element(
    element: &[u8],
) -> Result<Option<Vec<u8>>, VbaSignatureSignedDigestError> {
    let (tag, len, rest) = parse_tag_and_length(element)?;

    // SEQUENCE: some producers may encode SigDataV1Serialized directly as ASN.1. Try treating this
    // element as SigData before scanning its children.
    if tag.class == Asn1Class::Universal && tag.constructed && tag.number == 16 {
        if let Some(hash) = extract_source_hash_from_sig_data_v1_serialized(element)? {
            return Ok(Some(hash));
        }
    }

    // OCTET STRING: treat the value bytes as a candidate SigDataV1Serialized blob.
    if tag.class == Asn1Class::Universal && tag.number == 4 {
        let (octets, _after) = parse_octet_string(element)?;
        if let Some(hash) = extract_source_hash_from_sig_data_v1_serialized(&octets)? {
            return Ok(Some(hash));
        }
        return Ok(None);
    }

    // Constructed types: recursively scan their contents.
    if tag.constructed {
        let content = slice_constructed_contents(rest, len)?;
        return scan_asn1_for_sigdata_source_hash(content);
    }

    Ok(None)
}

fn scan_asn1_for_octet_string_len(
    mut cur: &[u8],
    desired_len: usize,
) -> Result<Option<Vec<u8>>, VbaSignatureSignedDigestError> {
    while !cur.is_empty() {
        if is_eoc(cur) {
            break;
        }
        let (tag, _len, _rest) = parse_tag_and_length(cur)?;
        if tag.class == Asn1Class::Universal && tag.number == 4 {
            let (octets, _after) = parse_octet_string(cur)?;
            if octets.len() == desired_len {
                return Ok(Some(octets));
            }
        } else if tag.constructed {
            let (_tag2, len2, rest2) = parse_tag_and_length(cur)?;
            let content = slice_constructed_contents(rest2, len2)?;
            if let Some(found) = scan_asn1_for_octet_string_len(content, desired_len)? {
                return Ok(Some(found));
            }
        }
        cur = skip_element(cur)?;
    }
    Ok(None)
}

fn extract_source_hash_from_sig_data_v1_serialized(
    bytes: &[u8],
) -> Result<Option<Vec<u8>>, VbaSignatureSignedDigestError> {
    // Some producers store `SigDataV1Serialized` as an embedded ASN.1 SEQUENCE (often inside an
    // OCTET STRING). Prefer a structure-aware parse to avoid accidentally interpreting unrelated
    // 16-byte OCTET STRINGs as the VBA project hash.
    if bytes.first() == Some(&0x30) {
        if let Some(hash) = extract_source_hash_from_sig_data_v1_serialized_asn1(bytes)? {
            return Ok(Some(hash));
        }
        return Ok(None);
    }

    // Otherwise treat it as an MS-OSHARED serialized binary structure and heuristically extract
    // the `sourceHash` BLOB. MS-OSHARED §4.3 specifies this is always a 16-byte MD5 digest.
    Ok(extract_source_hash_from_sig_data_v1_serialized_binary(bytes))
}

fn extract_source_hash_from_sig_data_v1_serialized_asn1(
    bytes: &[u8],
) -> Result<Option<Vec<u8>>, VbaSignatureSignedDigestError> {
    // SigDataV1Serialized (ASN.1 form; best-effort) has multiple observed shapes.
    //
    // Real-world producers appear to encode either:
    // 1) `SEQUENCE { version INTEGER, ... , sourceHash OCTET STRING }`, or
    // 2) `SEQUENCE { algorithmId AlgorithmIdentifier, sourceHash OCTET STRING }`.
    //
    // We handle both, but keep the parser conservative to avoid accidentally treating unrelated
    // 16-byte OCTET STRINGs as the VBA project hash.
    let Ok((tag, len, rest)) = parse_tag_and_length(bytes) else {
        return Ok(None);
    };
    if tag.class != Asn1Class::Universal || !tag.constructed || tag.number != 16 {
        return Ok(None);
    }
    let Ok(content) = slice_constructed_contents(rest, len) else {
        return Ok(None);
    };
    let mut cur = content;

    // Pattern (1): version INTEGER, then scan for a 16-byte OCTET STRING in the remainder.
    if let Some((version, after_ver)) = parse_integer_u32(cur)? {
        if (1..=0x100).contains(&version) {
            if let Some(hash) = scan_asn1_for_octet_string_len(after_ver, 16)? {
                return Ok(Some(hash));
            }
        }
    }

    // Pattern (2): algorithmId AlgorithmIdentifier, then `sourceHash` OCTET STRING (16 bytes).
    //
    // Validate that the first element looks like an AlgorithmIdentifier by requiring it to be a
    // SEQUENCE whose first child is an OBJECT IDENTIFIER. This keeps the match conservative.
    let (first_tag, first_len, first_rest) = match parse_tag_and_length(cur) {
        Ok(v) => v,
        Err(_) => return Ok(None),
    };
    if first_tag.class == Asn1Class::Universal && first_tag.constructed && first_tag.number == 16 {
        let alg_content = slice_constructed_contents(first_rest, first_len)?;
        if parse_oid(alg_content).is_ok() {
            cur = skip_element(cur)?;

            // Prefer the common direct-field encoding where `sourceHash` is the next element.
            if let Ok((tag, _len, _rest)) = parse_tag_and_length(cur) {
                if tag.class == Asn1Class::Universal && tag.number == 4 {
                    let (hash, _after) = parse_octet_string(cur)?;
                    if hash.len() == 16 {
                        return Ok(Some(hash));
                    }
                }
            }

            // Fallback: scan for an embedded 16-byte OCTET STRING once we've established the
            // AlgorithmIdentifier marker. Avoid scanning when we don't have *any* SigData markers.
            if let Some(hash) = scan_asn1_for_octet_string_len(cur, 16)? {
                return Ok(Some(hash));
            }
        }
    }

    // Pattern (3): MS-OSHARED SigDataV1Serialized ASN.1 as defined in §2.3.2.4.3.2:
    //
    // SigDataV1Serialized ::= SEQUENCE {
    //   algorithmIdSize INTEGER,
    //   compiledHashSize INTEGER,
    //   sourceHashSize INTEGER,
    //   algorithmIdOffset INTEGER,
    //   compiledHashOffset INTEGER,
    //   sourceHashOffset INTEGER,
    //   algorithmId OBJECT IDENTIFIER,
    //   compiledHash OCTET STRING,
    //   sourceHash OCTET STRING
    // }
    //
    // This is the normative MS-OSHARED structure used by `SpcIndirectDataContentV2`, where the VBA
    // signature binding digest bytes live in `sourceHash` (MD5, 16 bytes).
    let mut cur = content;
    for _ in 0..6 {
        let (tag, _len, _rest) = match parse_tag_and_length(cur) {
            Ok(v) => v,
            Err(_) => return Ok(None),
        };
        if tag.class != Asn1Class::Universal || tag.constructed || tag.number != 2 {
            return Ok(None);
        }
        cur = match skip_element(cur) {
            Ok(v) => v,
            Err(_) => return Ok(None),
        };
    }
    // algorithmId OBJECT IDENTIFIER
    let (_alg_oid, after_oid) = match parse_oid(cur) {
        Ok(v) => v,
        Err(_) => return Ok(None),
    };
    cur = after_oid;

    // compiledHash OCTET STRING (often empty)
    let (tag, _len, _rest) = match parse_tag_and_length(cur) {
        Ok(v) => v,
        Err(_) => return Ok(None),
    };
    if tag.class != Asn1Class::Universal || tag.number != 4 {
        return Ok(None);
    }
    cur = parse_octet_string(cur)?.1;

    // sourceHash OCTET STRING (MD5 bytes)
    let (tag, _len, _rest) = match parse_tag_and_length(cur) {
        Ok(v) => v,
        Err(_) => return Ok(None),
    };
    if tag.class != Asn1Class::Universal || tag.number != 4 {
        return Ok(None);
    }
    let (hash, _after) = parse_octet_string(cur)?;
    if hash.len() == 16 {
        return Ok(Some(hash));
    }

    Ok(None)
}

fn extract_source_hash_from_sig_data_v1_serialized_binary(bytes: &[u8]) -> Option<Vec<u8>> {
    fn read_u32_le(bytes: &[u8], offset: usize) -> Option<u32> {
        let end = offset.checked_add(4)?;
        let b = bytes.get(offset..end)?;
        Some(u32::from_le_bytes([b[0], b[1], b[2], b[3]]))
    }

    // SigDataV1Serialized is expected to begin with a small version field. Keep this strict to
    // avoid misidentifying unrelated octet strings as SigData.
    let version = read_u32_le(bytes, 0)?;
    if !(1..=0x100).contains(&version) {
        return None;
    }

    // Common (test + observed) pattern:
    //   [version u32][cbSourceHash u32][sourceHash bytes]
    if bytes.len() >= 8 {
        if let Some(len) = read_u32_le(bytes, 4).map(|n| n as usize) {
            let start = 8usize;
            let end = start.checked_add(16)?;
            if len == 16 && bytes.len() >= end {
                return bytes.get(start..end).map(|b| b.to_vec());
            }
        }
    }

    // Generic length-prefixed blob scan:
    // try interpreting the payload as a sequence of `[u32 len][len bytes]` blobs, optionally
    // preceded by a 4-byte version field.
    let mut offset = 4usize;
    let mut candidate = None;
    let mut steps = 0usize;
    while offset.checked_add(4).is_some_and(|end| end <= bytes.len()) && steps < 64 {
        let Some(len) = read_u32_le(bytes, offset).map(|n| n as usize) else {
            break;
        };
        offset += 4;
        let Some(end) = offset.checked_add(len) else {
            break;
        };
        if end > bytes.len() {
            break;
        }
        if len == 16 {
            let end16 = offset.checked_add(16)?;
            candidate = bytes.get(offset..end16).map(|b| b.to_vec());
        }
        offset += len;
        steps += 1;
    }
    if candidate.is_some() {
        return candidate;
    }

    None
}

fn parse_integer_u32(
    input: &[u8],
) -> Result<Option<(u32, &[u8])>, VbaSignatureSignedDigestError> {
    let (tag, len, rest) = match parse_tag_and_length(input) {
        Ok(v) => v,
        Err(_) => return Ok(None),
    };
    if tag.class != Asn1Class::Universal || tag.constructed || tag.number != 2 {
        return Ok(None);
    }
    let Length::Definite(l) = len else {
        return Err(der_err("INTEGER uses indefinite length"));
    };
    let val = match rest.get(..l) {
        Some(v) => v,
        None => return Err(der_err("INTEGER length exceeds input")),
    };
    let after = rest.get(l..).ok_or_else(|| der_err("unexpected EOF"))?;

    if val.is_empty() {
        return Ok(None);
    }
    // Ignore a single leading 0x00 used to force a positive sign bit.
    let val = if val.len() > 1 && val[0] == 0x00 {
        &val[1..]
    } else {
        val
    };
    if val.len() > 4 {
        return Ok(None);
    }
    if val.first().is_some_and(|b| b & 0x80 != 0) {
        // Negative integer; not a plausible version field for our use.
        return Ok(None);
    }
    let mut out: u32 = 0;
    for &b in val {
        out = out
            .checked_shl(8)
            .ok_or_else(|| der_err("INTEGER overflow"))?;
        out |= b as u32;
    }
    Ok(Some((out, after)))
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
    let bytes = input.get(1..1 + count).ok_or_else(|| der_err("unexpected EOF parsing length"))?;
    for &b in bytes {
        len = len
            .checked_shl(8)
            .ok_or_else(|| der_err("length overflow"))?;
        len |= b as usize;
    }
    Ok((Length::Definite(len), 1 + count))
}

fn slice_constructed_contents(
    rest_after_header: &[u8],
    len: Length,
) -> Result<&[u8], VbaSignatureSignedDigestError> {
    match len {
        Length::Definite(l) => rest_after_header
            .get(..l)
            .ok_or_else(|| der_err("length exceeds input")),
        Length::Indefinite => Ok(rest_after_header),
    }
}

fn parse_context_specific_constructed(
    input: &[u8],
    tag_number: u32,
) -> Result<&[u8], VbaSignatureSignedDigestError> {
    let (tag, len, rest) = parse_tag_and_length(input)?;
    if tag.class != Asn1Class::ContextSpecific || !tag.constructed || tag.number != tag_number {
        return Err(der_err("unexpected tag"));
    }
    slice_constructed_contents(rest, len)
}

fn parse_oid(input: &[u8]) -> Result<(&[u8], &[u8]), VbaSignatureSignedDigestError> {
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
    let after = rest.get(l..).ok_or_else(|| der_err("unexpected EOF"))?;
    Ok((val, after))
}

fn parse_octet_string(input: &[u8]) -> Result<(Vec<u8>, &[u8]), VbaSignatureSignedDigestError> {
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
        let after = rest.get(l..).ok_or_else(|| der_err("unexpected EOF"))?;
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
        Length::Definite(l) => rest.get(l..).ok_or_else(|| der_err("length exceeds input")),
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
