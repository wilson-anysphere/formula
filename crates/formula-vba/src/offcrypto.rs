//! Parsers for Office digital-signature wrapper structures used by VBA project signatures.
//!
//! Excel stores a signed VBA project in an OLE stream named `\x05DigitalSignature*`.
//! The stream payload is usually a PKCS#7/CMS `SignedData` `ContentInfo`, optionally wrapped in an
//! Office-specific header that includes size/offset metadata.
//!
//! Office produces (at least) two related wrapper shapes:
//!
//! - `[MS-OSHARED] DigSigBlob` (§2.3.2.2), which contains a `[MS-OSHARED] DigSigInfoSerialized`
//!   (§2.3.2.1) pointing at the PKCS#7 buffer via offsets.
//! - `[MS-OSHARED] WordSigBlob` (§2.3.2.3), which wraps the same `DigSigInfoSerialized` payload
//!   with a UTF-16 length prefix.
//! - A shorter *length-prefixed* DigSigInfoSerialized-like header (commonly seen in the wild) that
//!   starts with `cbSignature`, `cbSigningCertStore`, and a project-name length/count, followed by
//!   variable metadata blobs. This variant does **not** match the MS-OSHARED DigSigInfoSerialized
//!   layout, and is sometimes attributed to `[MS-OFFCRYPTO]` in older references.
//!
//! We implement best-effort parsers for both, primarily to locate the embedded PKCS#7 bytes
//! deterministically (instead of scanning for ASN.1 `SEQUENCE` tags) and to support
//! BER/indefinite-length encodings.
//!
//! MS-OSHARED reference:
//! - https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oshared/
//! - DigSigInfoSerialized (§2.3.2.1): https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oshared/30a00273-dbee-422f-b488-f4b8430ae046
//! - DigSigBlob (§2.3.2.2): https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oshared/bc21c922-b7ae-4736-90aa-86afb6403462

use md5::Md5;
use sha1::Sha1;
use sha2::Digest as _;

/// Parsed information from the length-prefixed DigSigInfoSerialized-like prefix.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct DigSigInfoSerialized {
    /// Offset (from the start of the stream) where the PKCS#7 `ContentInfo` begins.
    pub(crate) pkcs7_offset: usize,
    /// Length (in bytes) of the PKCS#7 `ContentInfo` TLV.
    pub(crate) pkcs7_len: usize,
    /// Best-effort version field when present.
    pub(crate) version: Option<u32>,
}

fn read_u32_le(bytes: &[u8], offset: usize) -> Option<u32> {
    let b = bytes.get(offset..offset + 4)?;
    Some(u32::from_le_bytes([b[0], b[1], b[2], b[3]]))
}

fn read_u16_le(bytes: &[u8], offset: usize) -> Option<u16> {
    let b = bytes.get(offset..offset + 2)?;
    Some(u16::from_le_bytes([b[0], b[1]]))
}

/// Parsed PKCS#7 location information from a `[MS-OSHARED] DigSigBlob` wrapper.
///
/// Office sometimes wraps the `\x05DigitalSignature*` stream contents in a DigSigBlob structure
/// that contains offsets to the embedded PKCS#7 buffer (`pbSignatureBuffer`). Parsing this is more
/// deterministic than scanning the whole stream for DER `SEQUENCE` tags.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct DigSigBlob {
    /// Offset (from the start of the stream) where the DER/BER-encoded PKCS#7 `ContentInfo` begins.
    pub(crate) pkcs7_offset: usize,
    /// Length (in bytes) of the PKCS#7 `ContentInfo` (best-effort; derived from BER length).
    pub(crate) pkcs7_len: usize,
}

/// Best-effort parse of a `[MS-OSHARED] DigSigBlob` wrapper around a PKCS#7 signature.
///
/// Returns `None` if the stream does not appear to be a DigSigBlob or if the referenced signature
/// buffer does not look like a PKCS#7 `SignedData` ContentInfo.
pub(crate) fn parse_digsig_blob(stream: &[u8]) -> Option<DigSigBlob> {
    // DigSigBlob header is at least two DWORDs: cb + serializedPointer.
    if stream.len() < 16 {
        return None;
    }

    let cb = read_u32_le(stream, 0)? as usize;
    let serialized_pointer = read_u32_le(stream, 4)? as usize;

    // MS-OSHARED examples use a fixed pointer to the serialized DigSigInfoSerialized at offset 8.
    // Requiring this keeps the heuristic conservative and avoids mis-detecting the other, distinct
    // length-prefixed DigSigInfoSerialized-like wrapper format.
    if serialized_pointer != 8 {
        return None;
    }

    // Ensure the declared signatureInfo region fits inside the stream when present. Some producers
    // may set `cb` to the length of the signatureInfo header only (excluding the signature bytes),
    // so we don't use it to bound `signatureOffset`.
    if serialized_pointer.checked_add(cb).is_none() || serialized_pointer + cb > stream.len() {
        return None;
    }

    // DigSigInfoSerialized (MS-OSHARED) begins at serializedPointer and starts with:
    //   DWORD cbSignature;
    //   DWORD signatureOffset;
    let cb_signature = read_u32_le(stream, serialized_pointer)? as usize;
    let signature_offset = read_u32_le(stream, serialized_pointer + 4)? as usize;
    if cb_signature == 0 {
        return None;
    }

    let sig_end = signature_offset.checked_add(cb_signature)?;
    if signature_offset >= stream.len() || sig_end > stream.len() {
        return None;
    }

    let sig_slice = &stream[signature_offset..sig_end];
    let pkcs7_len = pkcs7_signed_data_len(sig_slice)?;

    Some(DigSigBlob {
        pkcs7_offset: signature_offset,
        pkcs7_len,
    })
}

/// Best-effort parse of a `[MS-OSHARED] WordSigBlob` wrapper around a PKCS#7 signature.
///
/// WordSigBlob is similar to DigSigBlob but uses a UTF-16-length-prefixed wrapper:
/// - The first field is `cch: u16` (half the byte count of the remainder of the structure).
/// - Offsets in the embedded `DigSigInfoSerialized` are relative to the start of the `cbSigInfo`
///   field (at byte offset 2), not the start of the structure.
///
/// Returns `None` if the stream does not appear to be a WordSigBlob or if the inferred signature
/// buffer does not look like a PKCS#7 `SignedData` ContentInfo.
pub(crate) fn parse_wordsig_blob(stream: &[u8]) -> Option<DigSigBlob> {
    // Need at least: cch (u16) + cbSigInfo (u32) + serializedPointer (u32) + cbSignature+offset.
    if stream.len() < 2 + 4 + 4 + 8 {
        return None;
    }

    let cch = read_u16_le(stream, 0)? as usize;
    let remainder_len = cch.checked_mul(2)?;
    let total_len = 2usize.checked_add(remainder_len)?;
    if total_len > stream.len() {
        return None;
    }

    let cb_siginfo = read_u32_le(stream, 2)? as usize;
    let serialized_pointer = read_u32_le(stream, 6)? as usize;
    // MS-OSHARED requires a fixed pointer of 8 bytes from `cbSigInfo` to the DigSigInfoSerialized.
    if serialized_pointer != 8 {
        return None;
    }

    // Validate the `cch` formula from MS-OSHARED §2.3.2.3 to keep the heuristic conservative:
    // cch = (cbSigInfo + (cbSigInfo mod 2) + 8) / 2
    let expected_cch =
        (cb_siginfo.checked_add(cb_siginfo % 2)?.checked_add(8)?) / 2usize;
    if expected_cch != cch {
        return None;
    }

    // WordSigBlob offsets are relative to the start of `cbSigInfo` (byte offset 2).
    let base = 2usize;
    let siginfo_offset = base.checked_add(serialized_pointer)?;
    let siginfo_end = siginfo_offset.checked_add(cb_siginfo)?;
    if siginfo_offset + 8 > total_len || siginfo_end > total_len {
        return None;
    }

    // DigSigInfoSerialized starts with: DWORD cbSignature; DWORD signatureOffset;
    let cb_signature = read_u32_le(stream, siginfo_offset)? as usize;
    let signature_offset_rel = read_u32_le(stream, siginfo_offset + 4)? as usize;
    if cb_signature == 0 {
        return None;
    }

    let signature_offset = base.checked_add(signature_offset_rel)?;
    let sig_end = signature_offset.checked_add(cb_signature)?;
    if signature_offset >= total_len || sig_end > total_len {
        return None;
    }

    let sig_slice = &stream[signature_offset..sig_end];
    let pkcs7_len = pkcs7_signed_data_len(sig_slice)?;

    Some(DigSigBlob {
        pkcs7_offset: signature_offset,
        pkcs7_len,
    })
}

/// Best-effort parse of the *length-prefixed* DigSigInfoSerialized-like wrapper used by some
/// `\x05DigitalSignature*` streams.
///
/// Spec note: MS-OSHARED defines a different DigSigInfoSerialized structure (§2.3.2.1) used inside
/// DigSigBlob (§2.3.2.2) with offset fields (`signatureOffset`, `certStoreOffset`, ...). The
/// length-prefixed wrapper parsed here is a separate, commonly-seen on-disk shape.
///
/// Returns `None` if the stream does not look like a length-prefixed DigSigInfoSerialized-like
/// wrapper around a PKCS#7 payload.
///
/// Notes:
/// - This DigSigInfoSerialized-like wrapper contains several length-prefixed metadata blobs
///   (project name, certificate store, etc.). The order varies across producers/versions, so we try
///   a small set of deterministic layouts and validate by checking for a well-formed PKCS#7
///   `SignedData` `ContentInfo` at the computed offset.
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

    // Build a small set of header candidates. The structure uses little-endian DWORD fields.
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
        //
        // Some producers also appear to omit the terminating NUL from the count, so include +2 byte
        // variants (for UTF-16LE NUL) as well.
        for proj_bytes in [
            Some(header.proj_len),
            header.proj_len.checked_add(2),
            header.proj_len.checked_mul(2),
            header
                .proj_len
                .checked_mul(2)
                .and_then(|n| n.checked_add(2)),
        ]
        .into_iter()
        .flatten()
        {
            // The signature can appear at a small number of offsets depending on the ordering of
            // the (project name, cert store, signature) blobs.
            // Some producers include additional unknown blobs/fields before the signature bytes.
            // When the signature is stored as the final blob in the structure, `cbSignature` lets
            // us locate it by counting back from the end of the stream. Include that as an
            // additional candidate offset.
            let candidate_offsets = [
                header.header_size,                                 // sig first
                header.header_size.saturating_add(header.cert_len), // cert then sig
                header.header_size.saturating_add(proj_bytes),      // project then sig
                header
                    .header_size
                    .saturating_add(header.cert_len)
                    .saturating_add(proj_bytes), // project+cert then sig (or cert+project then sig)
                stream
                    .len()
                    .checked_sub(header.sig_len)
                    .filter(|&off| off >= header.header_size)
                    .unwrap_or(usize::MAX),
            ];

            for &pkcs7_offset in &candidate_offsets {
                if pkcs7_offset == usize::MAX {
                    continue;
                }
                let sig_end = match pkcs7_offset.checked_add(header.sig_len) {
                    Some(end) => end,
                    None => continue,
                };
                if sig_end > stream.len() {
                    continue;
                }

                let sig_slice = &stream[pkcs7_offset..sig_end];
                let Some(pkcs7_len) = pkcs7_signed_data_len(sig_slice) else {
                    continue;
                };

                let padding = header.sig_len.saturating_sub(pkcs7_len);
                let info = DigSigInfoSerialized {
                    pkcs7_offset,
                    pkcs7_len,
                    version: header.version,
                };

                match best {
                    Some((best_padding, best_info)) => {
                        if padding < best_padding
                            || (padding == best_padding
                                && info.pkcs7_offset > best_info.pkcs7_offset)
                        {
                            best = Some((padding, info));
                        }
                    }
                    None => best = Some((padding, info)),
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

    Some((
        BerTag {
            tag_byte: b0,
            constructed,
        },
        idx,
    ))
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

/// If `bytes` starts with a PKCS#7/CMS `ContentInfo` for `signedData`, return the total BER/DER
/// length of that object (including the tag/length header).
pub(crate) fn pkcs7_signed_data_len(bytes: &[u8]) -> Option<usize> {
    if !looks_like_pkcs7_signed_data(bytes) {
        return None;
    }
    ber_total_len(bytes)
}

fn looks_like_pkcs7_signed_data(bytes: &[u8]) -> bool {
    // ContentInfo ::= SEQUENCE { contentType OID, content [0] EXPLICIT ANY OPTIONAL }
    // For SignedData, contentType == 1.2.840.113549.1.7.2
    const SIGNED_DATA_OID: &[u8] = b"\x2A\x86\x48\x86\xF7\x0D\x01\x07\x02";

    let (tag, seq_len, hdr_len) = match ber_header(bytes) {
        Some(v) => v,
        None => return false,
    };
    if tag.tag_byte != 0x30 || !tag.constructed {
        return false;
    }
    let rest = match bytes.get(hdr_len..) {
        Some(v) => v,
        None => return false,
    };
    let rest = match seq_len {
        BerLen::Definite(l) => match rest.get(..l) {
            Some(v) => v,
            None => return false,
        },
        BerLen::Indefinite => rest,
    };

    // ContentInfo.contentType OBJECT IDENTIFIER
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
    if oid_bytes != SIGNED_DATA_OID {
        return false;
    }
    let mut cur = match rest.get(hdr2_len + oid_len..) {
        Some(v) => v,
        None => return false,
    };

    // ContentInfo.content [0] EXPLICIT
    let (tag3, explicit_len, hdr3_len) = match ber_header(cur) {
        Some(v) => v,
        None => return false,
    };
    if tag3.tag_byte != 0xA0 || !tag3.constructed {
        return false;
    }
    cur = match cur.get(hdr3_len..) {
        Some(v) => v,
        None => return false,
    };
    let cur = match explicit_len {
        BerLen::Definite(l) => match cur.get(..l) {
            Some(v) => v,
            None => return false,
        },
        BerLen::Indefinite => cur,
    };

    // SignedData ::= SEQUENCE { version INTEGER, digestAlgorithms SET, encapContentInfo SEQUENCE, ... }
    let (sd_tag, sd_len, sd_hdr_len) = match ber_header(cur) {
        Some(v) => v,
        None => return false,
    };
    if sd_tag.tag_byte != 0x30 || !sd_tag.constructed {
        return false;
    }
    let sd_rest = match cur.get(sd_hdr_len..) {
        Some(v) => v,
        None => return false,
    };
    let sd_rest = match sd_len {
        BerLen::Definite(l) => match sd_rest.get(..l) {
            Some(v) => v,
            None => return false,
        },
        BerLen::Indefinite => sd_rest,
    };

    // version INTEGER
    let (ver_tag, _ver_len, _ver_hdr) = match ber_header(sd_rest) {
        Some(v) => v,
        None => return false,
    };
    if ver_tag.tag_byte != 0x02 || ver_tag.constructed {
        return false;
    }
    let sd_rest = match ber_skip_any(sd_rest) {
        Some(v) => v,
        None => return false,
    };

    // digestAlgorithms SET
    let (dig_tag, _dig_len, _dig_hdr) = match ber_header(sd_rest) {
        Some(v) => v,
        None => return false,
    };
    if dig_tag.tag_byte != 0x31 || !dig_tag.constructed {
        return false;
    }
    let sd_rest = match ber_skip_any(sd_rest) {
        Some(v) => v,
        None => return false,
    };

    // encapContentInfo SEQUENCE
    let (enc_tag, _enc_len, _enc_hdr) = match ber_header(sd_rest) {
        Some(v) => v,
        None => return false,
    };
    enc_tag.tag_byte == 0x30 && enc_tag.constructed
}

pub(crate) const CALG_MD5: u32 = 0x8003;
pub(crate) const CALG_SHA1: u32 = 0x8004;

#[allow(dead_code)]
pub(crate) fn standard_cryptoapi_rc4_block_key(
    alg_id_hash: u32,
    password: &str,
    salt: &[u8],
    spin_count: u32,
    block: u32,
    key_size_bits: usize,
) -> Option<Vec<u8>> {
    // MS-OFFCRYPTO specifies that for Standard/CryptoAPI RC4, `EncryptionHeader.keySize == 0` MUST
    // be interpreted as 40-bit.
    let key_size_bits = if key_size_bits == 0 { 40 } else { key_size_bits };
    if !key_size_bits.is_multiple_of(8) {
        return None;
    }
    let key_size_bytes = key_size_bits / 8;

    let mut password_utf16le = Vec::with_capacity(password.len().saturating_mul(2));
    for ch in password.encode_utf16() {
        password_utf16le.extend_from_slice(&ch.to_le_bytes());
    }

    let mut h = cryptoapi_hash2(alg_id_hash, salt, &password_utf16le)?;
    for i in 0..spin_count {
        let i_bytes = (i as u32).to_le_bytes();
        h = cryptoapi_hash2(alg_id_hash, &i_bytes, &h)?;
    }

    let block_bytes = block.to_le_bytes();
    let mut key = cryptoapi_hash2(alg_id_hash, &h, &block_bytes)?;
    if key_size_bytes > key.len() {
        return None;
    }
    key.truncate(key_size_bytes);
    Some(key)
}

fn cryptoapi_hash2(alg_id_hash: u32, a: &[u8], b: &[u8]) -> Option<Vec<u8>> {
    match alg_id_hash {
        CALG_MD5 => {
            let mut hasher = Md5::new();
            hasher.update(a);
            hasher.update(b);
            Some(hasher.finalize().to_vec())
        }
        CALG_SHA1 => {
            let mut hasher = Sha1::new();
            hasher.update(a);
            hasher.update(b);
            Some(hasher.finalize().to_vec())
        }
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn hex_lower(bytes: &[u8]) -> String {
        bytes.iter().map(|b| format!("{b:02x}")).collect()
    }

    #[cfg(not(target_arch = "wasm32"))]
    fn make_pkcs7_signed_message(data: &[u8]) -> Vec<u8> {
        // Keep the test self-contained: generate a deterministic PKCS#7 SignedData with an
        // embedded self-signed certificate.
        use openssl::asn1::Asn1Time;
        use openssl::hash::MessageDigest;
        use openssl::pkcs7::{Pkcs7, Pkcs7Flags};
        use openssl::pkey::PKey;
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

    #[test]
    fn parses_digsig_info_serialized_with_version_prefix() {
        let pkcs7 = include_bytes!("../tests/fixtures/cms_indefinite.der");

        // Same as the previous test, but with an explicit version DWORD prefix:
        // [version, cbSignature, cbSigningCertStore, cchProjectName]
        let version = 1u32;

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
        stream.extend_from_slice(&version.to_le_bytes());
        stream.extend_from_slice(&cb_signature.to_le_bytes());
        stream.extend_from_slice(&cb_cert_store.to_le_bytes());
        stream.extend_from_slice(&cch_project.to_le_bytes());
        stream.extend_from_slice(&project_name_bytes);
        stream.extend_from_slice(&cert_store);
        stream.extend_from_slice(pkcs7);

        let parsed = parse_digsig_info_serialized(&stream).expect("should parse");
        let expected_offset = 16 + project_name_bytes.len() + cert_store.len();
        assert_eq!(parsed.pkcs7_offset, expected_offset);
        assert_eq!(parsed.pkcs7_len, pkcs7.len());
        assert_eq!(parsed.version, Some(version));
    }

    #[test]
    fn parses_digsig_info_serialized_with_reserved_and_byte_len_project_name() {
        let pkcs7 = include_bytes!("../tests/fixtures/cms_indefinite.der");

        // Variant:
        // [version, reserved, cbSignature, cbSigningCertStore, cchProjectName]
        // with cchProjectName interpreted as *byte* length rather than UTF-16 code unit count.
        let version = 1u32;
        let reserved = 0u32;

        let project_name_utf16: Vec<u16> = "VBAProject\0".encode_utf16().collect();
        let mut project_name_bytes = Vec::new();
        for ch in &project_name_utf16 {
            project_name_bytes.extend_from_slice(&ch.to_le_bytes());
        }
        let cert_store = vec![0xAA, 0xBB, 0xCC, 0xDD];

        let cb_signature = pkcs7.len() as u32;
        let cb_cert_store = cert_store.len() as u32;
        let cch_project_bytes = project_name_bytes.len() as u32;

        let mut stream = Vec::new();
        stream.extend_from_slice(&version.to_le_bytes());
        stream.extend_from_slice(&reserved.to_le_bytes());
        stream.extend_from_slice(&cb_signature.to_le_bytes());
        stream.extend_from_slice(&cb_cert_store.to_le_bytes());
        stream.extend_from_slice(&cch_project_bytes.to_le_bytes());
        stream.extend_from_slice(&project_name_bytes);
        stream.extend_from_slice(&cert_store);
        stream.extend_from_slice(pkcs7);

        let parsed = parse_digsig_info_serialized(&stream).expect("should parse");
        let expected_offset = 20 + project_name_bytes.len() + cert_store.len();
        assert_eq!(parsed.pkcs7_offset, expected_offset);
        assert_eq!(parsed.pkcs7_len, pkcs7.len());
        assert_eq!(parsed.version, Some(version));
    }

    #[test]
    fn parses_digsig_info_serialized_when_unknown_bytes_shift_signature_offset() {
        let pkcs7 = include_bytes!("../tests/fixtures/cms_indefinite.der");

        // Synthetic DigSigInfoSerialized-like stream with an extra unknown blob inserted between the
        // project name and cert store. This is not part of the "common" layout, but some producers
        // appear to include extra fields. We should still be able to locate the PKCS#7 payload using
        // cbSignature (counting backwards from the end).
        let project_name_utf16: Vec<u16> = "VBAProject\0".encode_utf16().collect();
        let mut project_name_bytes = Vec::new();
        for ch in &project_name_utf16 {
            project_name_bytes.extend_from_slice(&ch.to_le_bytes());
        }
        let unknown = vec![0x11, 0x22, 0x33];
        let cert_store = vec![0xAA, 0xBB, 0xCC, 0xDD];

        let cb_signature = pkcs7.len() as u32;
        let cb_cert_store = cert_store.len() as u32;
        let cch_project = project_name_utf16.len() as u32;

        let mut stream = Vec::new();
        stream.extend_from_slice(&cb_signature.to_le_bytes());
        stream.extend_from_slice(&cb_cert_store.to_le_bytes());
        stream.extend_from_slice(&cch_project.to_le_bytes());
        stream.extend_from_slice(&project_name_bytes);
        stream.extend_from_slice(&unknown);
        stream.extend_from_slice(&cert_store);
        stream.extend_from_slice(pkcs7);

        let parsed = parse_digsig_info_serialized(&stream).expect("should parse");
        let expected_offset = 12 + project_name_bytes.len() + unknown.len() + cert_store.len();
        assert_eq!(parsed.pkcs7_offset, expected_offset);
        assert_eq!(parsed.pkcs7_len, pkcs7.len());
    }

    #[test]
    fn parses_digsig_info_serialized_even_when_cert_store_len_is_inconsistent() {
        let pkcs7 = include_bytes!("../tests/fixtures/cms_indefinite.der");

        // Build a stream where `cbSigningCertStore` is clearly bogus (too large), but cbSignature is
        // accurate and the PKCS#7 blob is appended at the end of the stream.
        //
        // A permissive parser should still be able to find the PKCS#7 payload by counting back from
        // the end of the stream using cbSignature.
        let project_name_utf16: Vec<u16> = "VBAProject\0".encode_utf16().collect();
        let mut project_name_bytes = Vec::new();
        for ch in &project_name_utf16 {
            project_name_bytes.extend_from_slice(&ch.to_le_bytes());
        }
        let cert_store = vec![0xAA, 0xBB];

        let cb_signature = pkcs7.len() as u32;
        let cb_cert_store = 0xFFFF_FFFFu32; // intentionally inconsistent
        let cch_project = project_name_utf16.len() as u32;

        let mut stream = Vec::new();
        stream.extend_from_slice(&cb_signature.to_le_bytes());
        stream.extend_from_slice(&cb_cert_store.to_le_bytes());
        stream.extend_from_slice(&cch_project.to_le_bytes());
        stream.extend_from_slice(&project_name_bytes);
        stream.extend_from_slice(&cert_store);
        stream.extend_from_slice(pkcs7);

        let parsed = parse_digsig_info_serialized(&stream).expect("should parse");
        let expected_offset = stream.len().saturating_sub(pkcs7.len());
        assert_eq!(parsed.pkcs7_offset, expected_offset);
        assert_eq!(parsed.pkcs7_len, pkcs7.len());
    }

    #[test]
    fn parses_digsig_info_serialized_when_project_name_count_omits_nul() {
        let pkcs7 = include_bytes!("../tests/fixtures/cms_indefinite.der");

        // Variant: cchProjectName is a UTF-16 code unit count *excluding* the terminating NUL, but
        // the stored project name bytes still include the NUL terminator. This requires the parser
        // to consider the `proj_len * 2 + 2` interpretation.
        let project_name_utf16_with_nul: Vec<u16> = "VBAProject\0".encode_utf16().collect();
        let project_name_utf16_no_nul =
            &project_name_utf16_with_nul[..project_name_utf16_with_nul.len() - 1];

        let mut project_name_bytes = Vec::new();
        for ch in &project_name_utf16_with_nul {
            project_name_bytes.extend_from_slice(&ch.to_le_bytes());
        }
        let cert_store = vec![0xAA, 0xBB, 0xCC, 0xDD];

        let cb_signature = pkcs7.len() as u32;
        let cb_cert_store = cert_store.len() as u32;
        let cch_project_no_nul = project_name_utf16_no_nul.len() as u32;

        let mut stream = Vec::new();
        stream.extend_from_slice(&cb_signature.to_le_bytes());
        stream.extend_from_slice(&cb_cert_store.to_le_bytes());
        stream.extend_from_slice(&cch_project_no_nul.to_le_bytes());
        stream.extend_from_slice(&project_name_bytes);
        stream.extend_from_slice(&cert_store);
        stream.extend_from_slice(pkcs7);

        let parsed = parse_digsig_info_serialized(&stream).expect("should parse");
        let expected_offset = 12 + project_name_bytes.len() + cert_store.len();
        assert_eq!(parsed.pkcs7_offset, expected_offset);
        assert_eq!(parsed.pkcs7_len, pkcs7.len());
    }

    #[test]
    fn parses_digsig_info_serialized_when_signature_bytes_come_first() {
        let pkcs7 = include_bytes!("../tests/fixtures/cms_indefinite.der");

        let project_name_utf16: Vec<u16> = "VBAProject\0".encode_utf16().collect();
        let mut project_name_bytes = Vec::new();
        for ch in &project_name_utf16 {
            project_name_bytes.extend_from_slice(&ch.to_le_bytes());
        }
        let cert_store = vec![0xAA, 0xBB, 0xCC, 0xDD];

        let cb_signature = pkcs7.len() as u32;
        let cb_cert_store = cert_store.len() as u32;
        let cch_project = project_name_utf16.len() as u32;

        // Layout: header, signature, project name, cert store.
        let mut stream = Vec::new();
        stream.extend_from_slice(&cb_signature.to_le_bytes());
        stream.extend_from_slice(&cb_cert_store.to_le_bytes());
        stream.extend_from_slice(&cch_project.to_le_bytes());
        stream.extend_from_slice(pkcs7);
        stream.extend_from_slice(&project_name_bytes);
        stream.extend_from_slice(&cert_store);

        let parsed = parse_digsig_info_serialized(&stream).expect("should parse");
        assert_eq!(parsed.pkcs7_offset, 12);
        assert_eq!(parsed.pkcs7_len, pkcs7.len());
    }

    #[test]
    #[cfg(not(target_arch = "wasm32"))]
    fn parses_pkcs7_payload_from_digsig_blob_wrapper() {
        use openssl::pkcs7::Pkcs7;

        let pkcs7 = make_pkcs7_signed_message(b"formula-vba-test");

        // Minimal DigSigBlob:
        // - u32le cb
        // - u32le serializedPointer (=8)
        // - DigSigInfoSerialized (starts at offset 8):
        //     u32le cbSignature
        //     u32le signatureOffset
        //     (remaining fields set to 0)
        // - signature bytes at signatureOffset
        let digsig_blob_header_len = 8usize;
        // DigSigInfoSerialized is 9 DWORDs total in MS-OSHARED:
        // cbSignature, signatureOffset, cbSigningCertStore, certStoreOffset, cbProjectName,
        // projectNameOffset, fTimestamp, cbTimestampUrl, timestampUrlOffset.
        // For this synthetic fixture we only care about the first two, so the remaining fields are
        // set to 0.
        let digsig_info_len = 0x24usize;
        let signature_offset = digsig_blob_header_len + digsig_info_len; // 0x2C

        let mut stream = Vec::new();
        stream.extend_from_slice(&0u32.to_le_bytes()); // cb placeholder
        stream.extend_from_slice(&8u32.to_le_bytes()); // serializedPointer
        stream.extend_from_slice(&(pkcs7.len() as u32).to_le_bytes()); // cbSignature
        stream.extend_from_slice(&(signature_offset as u32).to_le_bytes()); // signatureOffset
        for _ in 0..7 {
            stream.extend_from_slice(&0u32.to_le_bytes());
        }
        assert_eq!(stream.len(), signature_offset);
        stream.extend_from_slice(&pkcs7);

        let cb = (stream.len().saturating_sub(digsig_blob_header_len)) as u32;
        stream[0..4].copy_from_slice(&cb.to_le_bytes());

        let parsed = parse_digsig_blob(&stream).expect("should parse digsig blob");
        assert_eq!(parsed.pkcs7_offset, signature_offset);
        assert_eq!(parsed.pkcs7_len, pkcs7.len());

        let pkcs7_slice = &stream[parsed.pkcs7_offset..parsed.pkcs7_offset + parsed.pkcs7_len];
        Pkcs7::from_der(pkcs7_slice).expect("openssl should parse extracted pkcs7");
    }

    #[test]
    fn parses_ber_indefinite_pkcs7_payload_from_digsig_blob_wrapper() {
        let pkcs7 = include_bytes!("../tests/fixtures/cms_indefinite.der");

        let digsig_blob_header_len = 8usize;
        let digsig_info_len = 0x24usize;
        let signature_offset = digsig_blob_header_len + digsig_info_len; // 0x2C

        let mut stream = Vec::new();
        stream.extend_from_slice(&0u32.to_le_bytes()); // cb placeholder
        stream.extend_from_slice(&8u32.to_le_bytes()); // serializedPointer
        stream.extend_from_slice(&(pkcs7.len() as u32).to_le_bytes()); // cbSignature
        stream.extend_from_slice(&(signature_offset as u32).to_le_bytes()); // signatureOffset
        for _ in 0..7 {
            stream.extend_from_slice(&0u32.to_le_bytes());
        }
        stream.extend_from_slice(pkcs7);

        let cb = (stream.len().saturating_sub(digsig_blob_header_len)) as u32;
        stream[0..4].copy_from_slice(&cb.to_le_bytes());

        let parsed = parse_digsig_blob(&stream).expect("should parse digsig blob");
        assert_eq!(parsed.pkcs7_offset, signature_offset);
        assert_eq!(parsed.pkcs7_len, pkcs7.len());
    }

    #[test]
    #[cfg(not(target_arch = "wasm32"))]
    fn parses_pkcs7_payload_from_wordsig_blob_wrapper() {
        use openssl::pkcs7::Pkcs7;

        let pkcs7 = make_pkcs7_signed_message(b"formula-vba-test");

        // Minimal WordSigBlob:
        // - u16le cch
        // - u32le cbSigInfo
        // - u32le serializedPointer (=8)
        // - DigSigInfoSerialized (starts at offset 10, because offsets are relative to cbSigInfo at 2)
        // - signature bytes at signatureOffset (relative to cbSigInfo)
        let wordsig_header_len = 10usize; // cch + cbSigInfo + serializedPointer
        let digsig_info_len = 0x24usize; // 9 DWORDs
        let signature_offset = wordsig_header_len + digsig_info_len; // 0x2E
        let signature_offset_rel = signature_offset - 2; // relative to cbSigInfo at offset 2

        let cb_siginfo = digsig_info_len + pkcs7.len();
        let cch = (cb_siginfo + (cb_siginfo % 2) + 8) / 2;

        let mut stream = Vec::new();
        stream.extend_from_slice(&(cch as u16).to_le_bytes());
        stream.extend_from_slice(&(cb_siginfo as u32).to_le_bytes());
        stream.extend_from_slice(&8u32.to_le_bytes()); // serializedPointer
        stream.extend_from_slice(&(pkcs7.len() as u32).to_le_bytes()); // cbSignature
        stream.extend_from_slice(&(signature_offset_rel as u32).to_le_bytes()); // signatureOffset (relative)
        for _ in 0..7 {
            stream.extend_from_slice(&0u32.to_le_bytes());
        }
        assert_eq!(stream.len(), signature_offset);
        stream.extend_from_slice(&pkcs7);
        if !cb_siginfo.is_multiple_of(2) {
            stream.push(0);
        }
        assert_eq!(stream.len(), 2 + cch * 2);

        let parsed = parse_wordsig_blob(&stream).expect("should parse wordsig blob");
        assert_eq!(parsed.pkcs7_offset, signature_offset);
        assert_eq!(parsed.pkcs7_len, pkcs7.len());

        let pkcs7_slice = &stream[parsed.pkcs7_offset..parsed.pkcs7_offset + parsed.pkcs7_len];
        Pkcs7::from_der(pkcs7_slice).expect("openssl should parse extracted pkcs7");
    }

    #[test]
    fn parses_ber_indefinite_pkcs7_payload_from_wordsig_blob_wrapper() {
        let pkcs7 = include_bytes!("../tests/fixtures/cms_indefinite.der");

        let wordsig_header_len = 10usize; // cch + cbSigInfo + serializedPointer
        let digsig_info_len = 0x24usize; // 9 DWORDs
        let signature_offset = wordsig_header_len + digsig_info_len; // 0x2E
        let signature_offset_rel = signature_offset - 2;

        let cb_siginfo = digsig_info_len + pkcs7.len();
        let cch = (cb_siginfo + (cb_siginfo % 2) + 8) / 2;

        let mut stream = Vec::new();
        stream.extend_from_slice(&(cch as u16).to_le_bytes());
        stream.extend_from_slice(&(cb_siginfo as u32).to_le_bytes());
        stream.extend_from_slice(&8u32.to_le_bytes()); // serializedPointer
        stream.extend_from_slice(&(pkcs7.len() as u32).to_le_bytes()); // cbSignature
        stream.extend_from_slice(&(signature_offset_rel as u32).to_le_bytes()); // signatureOffset
        for _ in 0..7 {
            stream.extend_from_slice(&0u32.to_le_bytes());
        }
        stream.extend_from_slice(pkcs7);
        if !cb_siginfo.is_multiple_of(2) {
            stream.push(0);
        }
        assert_eq!(stream.len(), 2 + cch * 2);

        let parsed = parse_wordsig_blob(&stream).expect("should parse wordsig blob");
        assert_eq!(parsed.pkcs7_offset, signature_offset);
        assert_eq!(parsed.pkcs7_len, pkcs7.len());
    }

    #[test]
    fn standard_cryptoapi_rc4_key_derivation_md5_vectors() {
        let password = "password";
        let salt: Vec<u8> = (0u8..=0x0F).collect();

        let expected = [
            (0u32, "69badcae244868e209d4e053ccd2a3bc"),
            (1u32, "6f4d502ab37700ffdab5704160455b47"),
            (2u32, "ac69022e396c7750872133f37e2c7afc"),
            (3u32, "1b056e7118ab8d35e9d67adee8b11104"),
        ];

        for (block, expected_hex) in expected {
            let key = standard_cryptoapi_rc4_block_key(CALG_MD5, password, &salt, 50_000, block, 128)
                .expect("should derive md5 rc4 block key");
            assert_eq!(hex_lower(&key), expected_hex, "block={block}");
        }

        let key_40 = standard_cryptoapi_rc4_block_key(CALG_MD5, password, &salt, 50_000, 0, 40)
            .expect("should derive md5 rc4 block key");
        assert_eq!(hex_lower(&key_40), "69badcae24");
        assert_eq!(key_40.len(), 5);
    }

    #[test]
    fn standard_cryptoapi_rc4_key_derivation_sha1_40_bit_truncates_to_5_bytes() {
        // Matches the SHA-1 worked example in `docs/offcrypto-standard-cryptoapi.md`.
        let password = "password";
        let salt: Vec<u8> = (0u8..=0x0F).collect();

        let key_40 = standard_cryptoapi_rc4_block_key(CALG_SHA1, password, &salt, 50_000, 0, 40)
            .expect("should derive sha1 rc4 block key");
        assert_eq!(hex_lower(&key_40), "6ad7dedf2d");
        assert_eq!(key_40.len(), 5);
    }

    #[test]
    fn standard_cryptoapi_rc4_key_derivation_sha1_keysize_0_is_interpreted_as_40bit() {
        // MS-OFFCRYPTO specifies that for Standard/CryptoAPI RC4, `keySize==0` MUST be interpreted
        // as 40-bit.
        let password = "password";
        let salt: Vec<u8> = (0u8..=0x0F).collect();

        let key_0 = standard_cryptoapi_rc4_block_key(CALG_SHA1, password, &salt, 50_000, 0, 0)
            .expect("should derive sha1 rc4 block key");
        assert_eq!(hex_lower(&key_0), "6ad7dedf2d");
        assert_eq!(key_0.len(), 5);
    }

    #[test]
    fn standard_cryptoapi_rc4_key_derivation_sha1_vectors() {
        // Deterministic vectors to lock in MS-OFFCRYPTO Standard/CryptoAPI RC4 block key derivation:
        // - password UTF-16LE (not UTF-8)
        // - H0 = SHA1(salt || password)
        // - spin loop: Hi = SHA1(LE32(i) || H(i-1)) for i in 0..50000
        // - per-block key: key(b) = SHA1(H || LE32(b))[0..keySizeBytes]
        let password = "password";
        let salt: Vec<u8> = (0u8..=0x0F).collect();

        let expected = [
            (0u32, "6ad7dedf2da3514b1d85eabee069d47d"),
            (1u32, "2ed4e8825cd48aa4a47994cda7415b4a"),
            (2u32, "9ce57d0699be3938951f47fa949361db"),
            (3u32, "e65b2643eaba3815a37a61159f137840"),
        ];

        for (block, expected_hex) in expected {
            let key = standard_cryptoapi_rc4_block_key(CALG_SHA1, password, &salt, 50_000, block, 128)
                .expect("should derive sha1 rc4 block key");
            assert_eq!(hex_lower(&key), expected_hex, "block={block}");
        }

        let key_56 = standard_cryptoapi_rc4_block_key(CALG_SHA1, password, &salt, 50_000, 0, 56)
            .expect("should derive sha1 rc4 block key");
        assert_eq!(hex_lower(&key_56), "6ad7dedf2da351");
        assert_eq!(key_56.len(), 7);

        let key_40 = standard_cryptoapi_rc4_block_key(CALG_SHA1, password, &salt, 50_000, 0, 40)
            .expect("should derive sha1 rc4 block key");
        assert_eq!(hex_lower(&key_40), "6ad7dedf2d");
        assert_eq!(key_40.len(), 5);
    }
}
