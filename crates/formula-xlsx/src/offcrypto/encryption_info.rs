//! MS-OFFCRYPTO Agile `EncryptionInfo` XML parsing.
//!
//! Modern Excel "Encrypt with Password" workbooks embed an XML document in the `EncryptionInfo`
//! stream (version 4.4). That XML describes:
//! - the cipher + KDF parameters used to encrypt the package payload
//! - one or more `<keyEncryptor>` entries (password, certificate, ...)
//! - optional integrity metadata
//!
//! This module provides:
//! - a best-effort parser focused on selecting the password key encryptor (and surfacing actionable
//!   errors when the file is certificate-encrypted), and
//! - bounded helpers for extracting the XML payload and decoding base64 attributes safely to avoid
//!   memory DoS on malicious/corrupt inputs.

use base64::engine::general_purpose::{STANDARD, STANDARD_NO_PAD};
use base64::Engine as _;
use std::borrow::Cow;

use super::{OffCryptoError, Result};

/// Parsing limits for MS-OFFCRYPTO Agile `EncryptionInfo` XML descriptors.
///
/// These defaults are intentionally generous for real-world Office files while still bounding
/// memory usage for malicious/corrupt inputs.
#[derive(Debug, Clone, Copy)]
pub struct ParseOptions {
    /// Maximum length (in bytes) of the raw XML payload stored in the `EncryptionInfo` stream
    /// **after** the 8-byte version header.
    ///
    /// This bounds:
    /// - allocation when the stream is read into memory, and
    /// - work performed by the XML parser.
    pub max_encryption_info_xml_len: usize,

    /// Maximum length (in bytes/chars) of a base64-encoded field **after whitespace stripping**.
    ///
    /// This bounds allocation for the intermediate stripped base64 buffer.
    pub max_base64_field_len: usize,

    /// Maximum length (in bytes) of a base64-decoded field.
    ///
    /// This bounds allocation for the decoded output buffer.
    pub max_base64_decoded_len: usize,
}

impl Default for ParseOptions {
    fn default() -> Self {
        // 1 MiB defaults: generous for real-world descriptors, small enough to prevent memory DoS.
        const ONE_MIB: usize = 1024 * 1024;
        Self {
            max_encryption_info_xml_len: ONE_MIB,
            max_base64_field_len: ONE_MIB,
            max_base64_decoded_len: ONE_MIB,
        }
    }
}

/// Extract the XML payload from an `EncryptionInfo` stream and enforce [`ParseOptions`] limits.
///
/// The `EncryptionInfo` stream begins with an 8-byte version header:
/// `majorVersion (u16le)`, `minorVersion (u16le)`, `flags (u32le)`.
///
/// For Agile encryption (`4.4`), the remainder of the stream is an XML document.
///
/// This helper returns the raw XML bytes (without copying) and errors if the payload exceeds
/// `max_encryption_info_xml_len`.
pub fn extract_encryption_info_xml<'a>(
    encryption_info_stream: &'a [u8],
    opts: &ParseOptions,
) -> Result<&'a [u8]> {
    let xml = encryption_info_stream.get(8..).unwrap_or(&[]);
    if xml.len() > opts.max_encryption_info_xml_len {
        return Err(OffCryptoError::EncryptionInfoTooLarge {
            len: xml.len(),
            max: opts.max_encryption_info_xml_len,
        });
    }
    Ok(xml)
}

fn trim_trailing_nul_bytes(mut bytes: &[u8]) -> &[u8] {
    while let Some((&last, rest)) = bytes.split_last() {
        if last == 0 {
            bytes = rest;
        } else {
            break;
        }
    }
    bytes
}

fn trim_start_ascii_whitespace(bytes: &[u8]) -> &[u8] {
    let mut idx = 0usize;
    while idx < bytes.len() {
        if bytes[idx].is_ascii_whitespace() {
            idx += 1;
        } else {
            break;
        }
    }
    &bytes[idx..]
}

fn strip_utf8_bom(bytes: &[u8]) -> &[u8] {
    bytes.strip_prefix(&[0xEF, 0xBB, 0xBF]).unwrap_or(bytes)
}

fn is_nul_heavy(bytes: &[u8]) -> bool {
    if bytes.is_empty() {
        return false;
    }
    let zeros = bytes.iter().filter(|&&b| b == 0).count();
    zeros > bytes.len() / 8
}

fn decode_utf16le_xml(bytes: &[u8]) -> Result<String> {
    let mut bytes = bytes;
    // Trim trailing UTF-16LE NUL terminators / padding.
    while bytes.len() >= 2 {
        let n = bytes.len();
        if bytes[n - 2] == 0 && bytes[n - 1] == 0 {
            bytes = &bytes[..n - 2];
        } else {
            break;
        }
    }

    if bytes.starts_with(&[0xFF, 0xFE]) {
        bytes = &bytes[2..];
    }

    // UTF-16 requires an even number of bytes; ignore a trailing odd byte.
    bytes = &bytes[..bytes.len().saturating_sub(bytes.len() % 2)];

    let mut units: Vec<u16> = Vec::with_capacity(bytes.len() / 2);
    for chunk in bytes.chunks_exact(2) {
        units.push(u16::from_le_bytes([chunk[0], chunk[1]]));
    }

    let mut xml = String::from_utf16(&units)?;
    // Be tolerant of a BOM encoded as U+FEFF.
    if let Some(stripped) = xml.strip_prefix('\u{FEFF}') {
        xml = stripped.to_string();
    }
    while xml.ends_with('\0') {
        xml.pop();
    }
    Ok(xml)
}

fn length_prefixed_slice(payload: &[u8]) -> Option<&[u8]> {
    let len_bytes: [u8; 4] = payload.get(0..4)?.try_into().ok()?;
    let len = u32::from_le_bytes(len_bytes) as usize;
    if len == 0 || len > payload.len().saturating_sub(4) {
        return None;
    }
    let candidate = payload.get(4..4 + len)?;

    // Ensure the candidate *looks* like XML to avoid false positives on arbitrary data.
    let trimmed = strip_utf8_bom(trim_start_ascii_whitespace(candidate));
    let looks_like_utf8 = trimmed.first() == Some(&b'<');
    let looks_like_utf16le = trimmed.starts_with(&[0xFF, 0xFE])
        || (trimmed.len() >= 2 && trimmed[0] == b'<' && trimmed[1] == 0);
    if !(looks_like_utf8 || looks_like_utf16le) {
        return None;
    }

    Some(candidate)
}

fn scan_to_encryption_tag(payload: &[u8]) -> Option<&[u8]> {
    const NEEDLE: &[u8] = b"<encryption";

    // Do not scan if the payload already looks like XML after trimming BOM + leading whitespace.
    let trimmed = trim_start_ascii_whitespace(strip_utf8_bom(payload));
    if trimmed.first() == Some(&b'<') {
        return None;
    }

    let idx = payload
        .windows(NEEDLE.len())
        .position(|w| w.eq_ignore_ascii_case(NEEDLE))?;
    Some(&payload[idx..])
}

/// Decode the XML payload bytes of an Agile `EncryptionInfo` stream into UTF-8 text.
///
/// Real-world Office producers vary in how the XML is wrapped/encoded. This helper supports:
/// - UTF-8 with an optional BOM and/or trailing NUL padding
/// - UTF-16LE (heuristic: NUL-heavy)
/// - a 4-byte little-endian length prefix before the XML
/// - leading junk before the `<encryption ...>` root tag (scan forward)
pub(super) fn decode_encryption_info_xml_text<'a>(payload: &'a [u8]) -> Result<Cow<'a, str>> {
    // Optional: a 4-byte little-endian length prefix before the XML.
    let candidate = length_prefixed_slice(payload)
        // Fallback: scan forward to the `<encryption` tag when the payload has leading bytes.
        .or_else(|| scan_to_encryption_tag(payload))
        .unwrap_or(payload);

    // UTF-16LE fallback heuristic (NUL-heavy buffers, or explicit BOM / `<\0` prefix).
    let looks_like_utf16le = candidate.starts_with(&[0xFF, 0xFE])
        || (candidate.len() >= 2 && candidate[0] == b'<' && candidate[1] == 0);
    if looks_like_utf16le || is_nul_heavy(candidate) {
        return Ok(Cow::Owned(decode_utf16le_xml(candidate)?));
    }

    let candidate = strip_utf8_bom(trim_trailing_nul_bytes(candidate));
    let xml = std::str::from_utf8(candidate)?;
    let xml = xml.strip_prefix('\u{FEFF}').unwrap_or(xml);
    Ok(Cow::Borrowed(xml))
}

fn count_non_ascii_whitespace_bytes(s: &str) -> usize {
    s.as_bytes()
        .iter()
        .filter(|b| !b.is_ascii_whitespace())
        .count()
}

fn base64_max_decoded_len(stripped_len: usize) -> usize {
    // Base64 expands by 4/3; the decoded length is at most 3 bytes for every 4 input chars.
    // Use a conservative upper bound so we can reject huge values before allocating the output.
    ((stripped_len.saturating_add(3)) / 4).saturating_mul(3)
}

/// Decode a base64 field from an Agile `EncryptionInfo` XML descriptor, enforcing size limits.
///
/// This function:
/// 1. Counts the field length after stripping ASCII whitespace (` \t\r\n...`) and rejects it if it
///    exceeds `max_base64_field_len`.
/// 2. Computes an upper bound for the decoded size and rejects it if it exceeds
///    `max_base64_decoded_len`.
/// 3. Allocates buffers sized to the bounded input/output and performs base64 decoding.
///
/// The `attr` name is included verbatim in [`OffCryptoError::FieldTooLarge`] for actionable error
/// reporting.
pub fn decode_base64_field_limited(
    element: &str,
    attr: &'static str,
    value: &str,
    opts: &ParseOptions,
) -> Result<Vec<u8>> {
    let bytes = value.as_bytes();
    let stripped_len = count_non_ascii_whitespace_bytes(value);
    if stripped_len > opts.max_base64_field_len {
        return Err(OffCryptoError::FieldTooLarge {
            field: attr,
            len: stripped_len,
            max: opts.max_base64_field_len,
        });
    }

    let max_decoded = base64_max_decoded_len(stripped_len);
    if max_decoded > opts.max_base64_decoded_len {
        return Err(OffCryptoError::FieldTooLarge {
            field: attr,
            len: max_decoded,
            max: opts.max_base64_decoded_len,
        });
    }

    // Avoid allocating a stripped copy when the attribute value contains no ASCII whitespace.
    let decoded = if stripped_len == bytes.len() {
        STANDARD
            .decode(bytes)
            .or_else(|_| STANDARD_NO_PAD.decode(bytes))
    } else {
        let mut stripped = Vec::with_capacity(stripped_len);
        stripped.extend(bytes.iter().copied().filter(|b| !b.is_ascii_whitespace()));
        STANDARD
            .decode(&stripped)
            .or_else(|_| STANDARD_NO_PAD.decode(&stripped))
    }
    .map_err(|source| OffCryptoError::Base64Decode {
        element: element.to_string(),
        attr: attr.to_string(),
        source,
    })?;

    // Defensive: should be redundant with `max_decoded` check, but keep to be safe if the base64
    // engine behavior changes.
    if decoded.len() > opts.max_base64_decoded_len {
        return Err(OffCryptoError::FieldTooLarge {
            field: attr,
            len: decoded.len(),
            max: opts.max_base64_decoded_len,
        });
    }

    Ok(decoded)
}

/// Password key encryptor URI as used by MS-OFFCRYPTO Agile EncryptionInfo XML.
pub const KEY_ENCRYPTOR_URI_PASSWORD: &str =
    "http://schemas.microsoft.com/office/2006/keyEncryptor/password";
/// Certificate key encryptor URI as used by MS-OFFCRYPTO Agile EncryptionInfo XML.
pub const KEY_ENCRYPTOR_URI_CERTIFICATE: &str =
    "http://schemas.microsoft.com/office/2006/keyEncryptor/certificate";

/// Warnings produced while parsing `EncryptionInfo` XML.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum EncryptionInfoWarning {
    /// Multiple password `<keyEncryptor>` entries were present.
    ///
    /// Parsing is deterministic: the first password key encryptor wins.
    MultiplePasswordKeyEncryptors { count: usize },
}

/// Parsed key encryptor information for password-based encryption.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PasswordKeyEncryptor {
    /// The `uri` attribute of the selected `<keyEncryptor>` element.
    pub uri: String,
}

/// Parsed Agile `EncryptionInfo` XML (best-effort; currently focused on key encryptor selection).
///
/// This is intentionally distinct from the full Agile descriptor parsed by `offcrypto::agile`
/// (`AgileKeyData`, `AgileDataIntegrity`, etc.). The goal of this lightweight type is to provide
/// deterministic, user-friendly diagnostics about key-encryptor selection (password vs certificate)
/// without having to fully parse the cryptographic parameters.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AgileEncryptionInfoXml {
    /// The selected password-based key encryptor.
    pub password_key_encryptor: PasswordKeyEncryptor,
    /// Non-fatal parse warnings (deterministic; suitable for telemetry/corpus triage).
    pub warnings: Vec<EncryptionInfoWarning>,
}

/// Parse the XML payload of an Agile (4.4) `EncryptionInfo` stream.
///
/// The caller is responsible for reading the `EncryptionInfo` stream header and providing only the
/// XML bytes.
pub fn parse_agile_encryption_info_xml(xml: &[u8]) -> Result<AgileEncryptionInfoXml> {
    let xml = decode_encryption_info_xml_text(xml)?;
    let doc = roxmltree::Document::parse(xml.as_ref())?;

    let root = doc.root_element();

    // Detect unsupported cipher chaining modes early.
    //
    // Some producers declare AES with CFB chaining in the Agile descriptor. Formula only supports
    // `ChainingModeCBC`, so fail fast with an actionable error instead of attempting decryption.
    if let Some(key_data) = root
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "keyData")
    {
        if let Some(chaining) = key_data.attribute("cipherChaining") {
            validate_cipher_chaining(chaining)?;
        }
    }

    let key_encryptors = root
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "keyEncryptors")
        .ok_or_else(|| OffCryptoError::MissingRequiredElement {
            element: "keyEncryptors".to_string(),
        })?;

    let mut available_uris: Vec<String> = Vec::new();
    let mut password_uri_count = 0usize;
    let mut selected_password_uri: Option<String> = None;
    let mut selected_password_encryptor: Option<roxmltree::Node<'_, '_>> = None;

    for enc in key_encryptors
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "keyEncryptor")
    {
        let uri = enc
            .attribute("uri")
            .ok_or_else(|| OffCryptoError::MissingRequiredAttribute {
                element: "keyEncryptor".to_string(),
                attr: "uri".to_string(),
            })?;

        // Keep a deterministic list of URIs for error reporting. Prefer unique values but preserve
        // first-seen ordering.
        if !available_uris.iter().any(|u| u == uri) {
            available_uris.push(uri.to_string());
        }

        if uri == KEY_ENCRYPTOR_URI_PASSWORD {
            password_uri_count += 1;
            if selected_password_uri.is_none() {
                selected_password_uri = Some(uri.to_string());
                selected_password_encryptor = Some(enc);
            }
        }
    }

    let Some(uri) = selected_password_uri else {
        let mut msg = String::new();
        msg.push_str("unsupported key encryptor in Agile EncryptionInfo: ");
        msg.push_str("Formula currently supports only password-based encryption. ");

        if available_uris.is_empty() {
            msg.push_str("No `<keyEncryptor>` entries were found.");
        } else {
            msg.push_str("Found keyEncryptor URIs: ");
            msg.push_str(&available_uris.join(", "));
            msg.push('.');
        }

        if available_uris
            .iter()
            .any(|u| u == KEY_ENCRYPTOR_URI_CERTIFICATE)
        {
            msg.push_str(" This file appears to be certificate-encrypted (public/private key) rather than password-encrypted. Re-save the workbook in Excel using “Encrypt with Password”.");
        } else {
            msg.push_str(
                " Re-save the workbook in Excel using “Encrypt with Password” (not certificate-based protection).",
            );
        }

        return Err(OffCryptoError::UnsupportedKeyEncryptor {
            available_uris,
            message: msg,
        });
    };

    if let Some(enc) = selected_password_encryptor {
        // Validate `cipherChaining` on the selected password `<p:encryptedKey>` when present.
        // (Match by local name so the namespace prefix doesn't matter.)
        if let Some(encrypted_key) = enc
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "encryptedKey")
        {
            if let Some(chaining) = encrypted_key.attribute("cipherChaining") {
                validate_cipher_chaining(chaining)?;
            }
        }
    }

    let mut warnings = Vec::new();
    if password_uri_count > 1 {
        warnings.push(EncryptionInfoWarning::MultiplePasswordKeyEncryptors {
            count: password_uri_count,
        });
    }

    Ok(AgileEncryptionInfoXml {
        password_key_encryptor: PasswordKeyEncryptor { uri },
        warnings,
    })
}

fn validate_cipher_chaining(chaining: &str) -> Result<()> {
    let chaining = chaining.trim();
    if chaining.eq_ignore_ascii_case("ChainingModeCBC") {
        Ok(())
    } else {
        Err(OffCryptoError::UnsupportedCipherChaining {
            chaining: chaining.to_string(),
        })
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn selects_password_key_encryptor_when_multiple_present() {
        let xml = r#"
            <encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
                        xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password"
                        xmlns:c="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate">
              <keyEncryptors>
                <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate">
                  <c:encryptedKey/>
                </keyEncryptor>
                <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
                  <p:encryptedKey spinCount="1"/>
                </keyEncryptor>
              </keyEncryptors>
            </encryption>
        "#;

        let info = parse_agile_encryption_info_xml(xml.as_bytes()).expect("parse should succeed");
        assert_eq!(info.password_key_encryptor.uri, KEY_ENCRYPTOR_URI_PASSWORD);
        assert!(info.warnings.is_empty());
    }

    #[test]
    fn errors_when_password_key_encryptor_missing() {
        let xml = r#"
            <encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
                        xmlns:c="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate">
              <keyEncryptors>
                <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate">
                  <c:encryptedKey/>
                </keyEncryptor>
              </keyEncryptors>
            </encryption>
        "#;

        let err = parse_agile_encryption_info_xml(xml.as_bytes()).expect_err("expected error");
        match &err {
            OffCryptoError::UnsupportedKeyEncryptor { available_uris, .. } => {
                assert!(
                    available_uris.iter().any(|u| u == KEY_ENCRYPTOR_URI_CERTIFICATE),
                    "expected certificate URI to be reported, got {available_uris:?}"
                );
            }
            other => panic!("expected UnsupportedKeyEncryptor, got {other:?}"),
        }

        let msg = err.to_string();
        assert!(
            msg.contains(KEY_ENCRYPTOR_URI_CERTIFICATE)
                || msg.to_ascii_lowercase().contains("certificate"),
            "expected error message to mention certificate encryption; got: {msg}"
        );
    }

    #[test]
    fn warns_on_multiple_password_key_encryptors() {
        let xml = r#"
            <encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
                        xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
              <keyEncryptors>
                <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
                  <p:encryptedKey spinCount="1"/>
                </keyEncryptor>
                <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
                  <p:encryptedKey spinCount="2"/>
                </keyEncryptor>
              </keyEncryptors>
            </encryption>
        "#;

        let info = parse_agile_encryption_info_xml(xml.as_bytes()).expect("parse should succeed");
        assert_eq!(info.password_key_encryptor.uri, KEY_ENCRYPTOR_URI_PASSWORD);
        assert_eq!(
            info.warnings,
            vec![EncryptionInfoWarning::MultiplePasswordKeyEncryptors { count: 2 }]
        );
    }

    #[test]
    fn rejects_cfb_cipher_chaining_in_key_data() {
        let xml = r#"
            <encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
                        xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
              <keyData cipherChaining="ChainingModeCFB" />
              <keyEncryptors>
                <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
                  <p:encryptedKey cipherChaining="ChainingModeCBC" />
                </keyEncryptor>
              </keyEncryptors>
            </encryption>
        "#;

        let err = parse_agile_encryption_info_xml(xml.as_bytes()).expect_err("expected error");
        assert!(
            matches!(err, OffCryptoError::UnsupportedCipherChaining { ref chaining } if chaining == "ChainingModeCFB"),
            "unexpected error: {err:?}"
        );
    }

    #[test]
    fn rejects_cfb_cipher_chaining_in_encrypted_key() {
        let xml = r#"
            <encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
                        xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
              <keyData cipherChaining="ChainingModeCBC" />
              <keyEncryptors>
                <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
                  <p:encryptedKey cipherChaining="ChainingModeCFB" />
                </keyEncryptor>
              </keyEncryptors>
            </encryption>
        "#;

        let err = parse_agile_encryption_info_xml(xml.as_bytes()).expect_err("expected error");
        assert!(
            matches!(err, OffCryptoError::UnsupportedCipherChaining { ref chaining } if chaining == "ChainingModeCFB"),
            "unexpected error: {err:?}"
        );

        let msg = err.to_string();
        assert!(
            msg.contains("only") && msg.contains("ChainingModeCBC"),
            "expected message to mention only CBC is supported, got: {msg}"
        );
    }

    #[test]
    fn base64_whitespace_is_stripped_before_counting() {
        let opts = ParseOptions {
            max_base64_field_len: 4,
            max_base64_decoded_len: 1024,
            ..ParseOptions::default()
        };

        // Base64 "AA==" (1 byte) but with whitespace.
        let decoded =
            decode_base64_field_limited("keyData", "saltValue", " A A = = ", &opts).unwrap();
        assert_eq!(decoded, vec![0]);
    }
}
