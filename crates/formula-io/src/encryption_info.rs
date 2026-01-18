use encoding_rs::UTF_16LE;
use thiserror::Error;

/// Errors returned by [`extract_agile_encryption_info_xml`].
#[derive(Debug, Error)]
pub enum EncryptionInfoXmlError {
    #[error("EncryptionInfo stream is truncated: expected at least 8 bytes, got {0}")]
    Truncated(usize),
    #[error("failed to extract Agile EncryptionInfo XML: {0}")]
    InvalidXml(String),
}

/// Extract the Agile `EncryptionInfo` XML payload from an `EncryptionInfo` OLE stream.
///
/// Office-encrypted OOXML containers store encryption metadata in a stream named `EncryptionInfo`.
/// The stream begins with an 8-byte `EncryptionVersionInfo` header, followed by an XML document
/// describing the encryption parameters (Agile encryption).
///
/// Real-world producers vary in how they encode/wrap that XML:
/// - UTF-8 with an optional UTF-8 BOM.
/// - UTF-16LE XML (often without an explicit BOM).
/// - A 4-byte length prefix before the XML.
/// - Leading/trailing padding (e.g. NUL bytes).
///
/// This helper extracts and validates the embedded XML document. It returns the decoded XML string
/// (without BOMs / trailing NUL terminators) and ensures:
/// - It is well-formed XML.
/// - The root element is `encryption` (case-insensitive).
pub fn extract_agile_encryption_info_xml(
    encryption_info_stream: &[u8],
) -> Result<String, EncryptionInfoXmlError> {
    if encryption_info_stream.len() < 8 {
        return Err(EncryptionInfoXmlError::Truncated(
            encryption_info_stream.len(),
        ));
    }

    let payload = &encryption_info_stream[8..];
    // Defensive: bound the amount of work we do scanning/parsing attacker-controlled
    // `EncryptionInfo` streams.
    //
    // Real-world Agile `EncryptionInfo` XML descriptors are tiny (typically a few KB). Allow up to
    // 1 MiB to be generous while preventing pathological memory/CPU usage on corrupt inputs.
    const MAX_XML_LEN: usize = 1024 * 1024;
    if payload.len() > MAX_XML_LEN {
        return Err(EncryptionInfoXmlError::InvalidXml(format!(
            "payload too large: {} bytes (max {MAX_XML_LEN})",
            payload.len()
        )));
    }
    let mut errors: Vec<String> = Vec::new();

    // --- Primary: UTF-8 payload (trim UTF-8 BOM, trim trailing NULs). ---
    match parse_utf8_xml(payload) {
        Ok(xml) => return Ok(xml),
        Err(err) => errors.push(format!("utf-8: {err}")),
    }

    // --- Fallback: UTF-16LE when there are many NUL bytes (ASCII UTF-16 pattern). ---
    if is_nul_heavy(payload) {
        match parse_utf16le_xml(payload) {
            Ok(xml) => return Ok(xml),
            Err(err) => errors.push(format!("utf-16le: {err}")),
        }
    }

    // --- Fallback: length-prefix heuristic (u32 LE) ---
    if let Some(len_slice) = length_prefixed_slice(payload) {
        match parse_utf8_xml(len_slice) {
            Ok(xml) => return Ok(xml),
            Err(err) => errors.push(format!("len+utf-8: {err}")),
        }
        if is_nul_heavy(len_slice) {
            match parse_utf16le_xml(len_slice) {
                Ok(xml) => return Ok(xml),
                Err(err) => errors.push(format!("len+utf-16le: {err}")),
            }
        }
    }

    // --- Fallback: scan for a `<encryption>...</encryption>` blob and ignore any trailing bytes. ---
    if let Some(scanned) = scan_to_encryption_xml_blob(payload) {
        match parse_utf8_xml(scanned) {
            Ok(xml) => return Ok(xml),
            Err(err) => errors.push(format!("blob+utf-8: {err}")),
        }
        if is_nul_heavy(scanned) {
            match parse_utf16le_xml(scanned) {
                Ok(xml) => return Ok(xml),
                Err(err) => errors.push(format!("blob+utf-16le: {err}")),
            }
        }
    }

    // --- Fallback: scan forward to the first `<` when the payload contains `<encryption` later. ---
    if let Some(scanned) = scan_to_first_xml_tag(payload) {
        match parse_utf8_xml(scanned) {
            Ok(xml) => return Ok(xml),
            Err(err) => errors.push(format!("scan+utf-8: {err}")),
        }
        if is_nul_heavy(scanned) {
            match parse_utf16le_xml(scanned) {
                Ok(xml) => return Ok(xml),
                Err(err) => errors.push(format!("scan+utf-16le: {err}")),
            }
        }
    }

    if errors.is_empty() {
        errors.push("no candidates".to_string());
    }

    Err(EncryptionInfoXmlError::InvalidXml(errors.join("; ")))
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

fn trim_trailing_utf16le_nul_units(mut bytes: &[u8]) -> &[u8] {
    while bytes.len() >= 2 {
        let n = bytes.len();
        if bytes[n - 2] == 0 && bytes[n - 1] == 0 {
            bytes = &bytes[..n - 2];
        } else {
            break;
        }
    }
    bytes
}

fn trim_utf8_bom(bytes: &[u8]) -> &[u8] {
    bytes.strip_prefix(&[0xEF, 0xBB, 0xBF]).unwrap_or(bytes)
}

fn trim_start_ascii_whitespace(bytes: &[u8]) -> &[u8] {
    let mut idx = 0usize;
    while idx < bytes.len() {
        if matches!(bytes[idx], b' ' | b'\t' | b'\r' | b'\n') {
            idx += 1;
        } else {
            break;
        }
    }
    &bytes[idx..]
}

fn is_nul_heavy(bytes: &[u8]) -> bool {
    if bytes.is_empty() {
        return false;
    }
    let zeros = bytes.iter().filter(|&&b| b == 0).count();
    zeros > bytes.len() / 8
}

fn validate_agile_encryption_xml(xml: &str) -> Result<(), String> {
    let doc = roxmltree::Document::parse(xml).map_err(|e| e.to_string())?;
    let root = doc.root_element();
    let name = root.tag_name().name();
    if !name.eq_ignore_ascii_case("encryption") {
        return Err(format!(
            "unexpected root element `{name}` (expected `encryption`)"
        ));
    }
    Ok(())
}

fn parse_utf8_xml(bytes: &[u8]) -> Result<String, String> {
    let bytes = trim_trailing_nul_bytes(bytes);
    let bytes = trim_utf8_bom(bytes);
    let xml = std::str::from_utf8(bytes).map_err(|e| e.to_string())?;
    // In case the stream was decoded through a path that preserved U+FEFF.
    let xml = xml.strip_prefix('\u{FEFF}').unwrap_or(xml);
    validate_agile_encryption_xml(xml)?;
    Ok(xml.to_string())
}

fn parse_utf16le_xml(bytes: &[u8]) -> Result<String, String> {
    let mut bytes = trim_trailing_utf16le_nul_units(bytes);
    if bytes.starts_with(&[0xFF, 0xFE]) {
        bytes = &bytes[2..];
    }
    // UTF-16 requires an even number of bytes; ignore a trailing odd byte.
    bytes = &bytes[..bytes.len().saturating_sub(bytes.len() % 2)];

    let (cow, _) = UTF_16LE.decode_without_bom_handling(bytes);
    let mut xml = cow.into_owned();
    if let Some(stripped) = xml.strip_prefix('\u{FEFF}') {
        xml = stripped.to_string();
    }
    while xml.ends_with('\0') {
        xml.pop();
    }
    validate_agile_encryption_xml(&xml)?;
    Ok(xml)
}

fn length_prefixed_slice(payload: &[u8]) -> Option<&[u8]> {
    let len_bytes: [u8; 4] = payload.get(0..4)?.try_into().ok()?;
    let len = u32::from_le_bytes(len_bytes) as usize;
    if len == 0 || len > payload.len().saturating_sub(4) {
        return None;
    }
    let end = 4usize.checked_add(len)?;
    let candidate = payload.get(4..end)?;

    // Ensure the candidate *looks* like XML to avoid false positives on arbitrary data.
    let candidate_trimmed = trim_start_ascii_whitespace(trim_utf8_bom(candidate));
    if candidate_trimmed.first() != Some(&b'<') {
        return None;
    }

    Some(candidate)
}

fn scan_to_first_xml_tag(payload: &[u8]) -> Option<&[u8]> {
    // Be conservative: only scan if we see the expected root tag bytes somewhere later.
    const NEEDLE: &[u8] = b"<encryption";
    if !payload
        .windows(NEEDLE.len())
        .any(|w| w.eq_ignore_ascii_case(NEEDLE))
    {
        return None;
    }

    let payload = trim_utf8_bom(payload);
    let trimmed = trim_start_ascii_whitespace(payload);
    if trimmed.first() == Some(&b'<') {
        return None;
    }

    let idx = payload.iter().position(|&b| b == b'<')?;
    Some(&payload[idx..])
}

fn scan_to_encryption_xml_blob(payload: &[u8]) -> Option<&[u8]> {
    const START: &[u8] = b"<encryption";
    const END: &[u8] = b"</encryption>";

    let start = payload
        .windows(START.len())
        .position(|w| w.eq_ignore_ascii_case(START))?;
    let after_start = start.checked_add(START.len())?;
    let end_rel = payload
        .get(after_start..)?
        .windows(END.len())
        .position(|w| w.eq_ignore_ascii_case(END))?;
    let end = after_start
        .checked_add(end_rel)?
        .checked_add(END.len())?;
    payload.get(start..end)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn make_stream(payload: &[u8]) -> Vec<u8> {
        let mut out = vec![0u8; 8]; // EncryptionVersionInfo (ignored by the extractor)
        out.extend_from_slice(payload);
        out
    }

    #[test]
    fn extracts_utf8_xml_with_bom_and_trailing_nuls() {
        let xml = r#"<encryption><keyData/></encryption>"#;
        let mut payload = Vec::new();
        payload.extend_from_slice(&[0xEF, 0xBB, 0xBF]); // UTF-8 BOM
        payload.extend_from_slice(xml.as_bytes());
        payload.extend_from_slice(&[0, 0, 0]);

        let stream = make_stream(&payload);
        let extracted = extract_agile_encryption_info_xml(&stream).expect("should parse");
        assert_eq!(extracted, xml);
    }

    #[test]
    fn extracts_utf16le_encoded_xml() {
        let xml = r#"<encryption><keyData/></encryption>"#;
        let mut payload = Vec::new();
        payload.extend_from_slice(&[0xFF, 0xFE]); // UTF-16LE BOM
        for unit in xml.encode_utf16() {
            payload.extend_from_slice(&unit.to_le_bytes());
        }
        payload.extend_from_slice(&[0x00, 0x00]); // terminating NUL

        let stream = make_stream(&payload);
        let extracted = extract_agile_encryption_info_xml(&stream).expect("should parse");
        assert_eq!(extracted, xml);
    }

    #[test]
    fn extracts_length_prefixed_utf8_xml_even_with_trailing_garbage() {
        let xml = r#"<encryption><keyData/></encryption>"#;
        let xml_bytes = xml.as_bytes();
        let mut payload = Vec::new();
        payload.extend_from_slice(&(xml_bytes.len() as u32).to_le_bytes());
        payload.extend_from_slice(xml_bytes);
        payload.extend_from_slice(b"GARBAGE"); // non-whitespace bytes to force length-prefix slicing

        let stream = make_stream(&payload);
        let extracted = extract_agile_encryption_info_xml(&stream).expect("should parse");
        assert_eq!(extracted, xml);
    }

    #[test]
    fn extracts_utf8_xml_by_scanning_to_end_tag_when_trailing_bytes_exist() {
        let xml = r#"<encryption><keyData/></encryption>"#;
        let mut payload = Vec::new();
        payload.extend_from_slice(xml.as_bytes());
        payload.extend_from_slice(b"GARBAGE");

        let stream = make_stream(&payload);
        let extracted = extract_agile_encryption_info_xml(&stream).expect("should parse");
        assert_eq!(extracted, xml);
    }

    #[test]
    fn rejects_invalid_xml_payloads() {
        let stream = make_stream(b"not xml at all");
        assert!(extract_agile_encryption_info_xml(&stream).is_err());

        // Well-formed XML but wrong root element.
        let stream = make_stream(b"<foo/>");
        assert!(extract_agile_encryption_info_xml(&stream).is_err());
    }

    #[test]
    fn rejects_truncated_streams() {
        let err = extract_agile_encryption_info_xml(&[0u8; 7]).expect_err("should error");
        assert!(matches!(err, EncryptionInfoXmlError::Truncated(7)));
    }

    #[test]
    fn rejects_overly_large_payloads() {
        // This should fail the size guardrail before any XML parsing is attempted.
        let payload = vec![b'a'; 1024 * 1024 + 1];
        let stream = make_stream(&payload);
        let err = extract_agile_encryption_info_xml(&stream).expect_err("should error");
        match err {
            EncryptionInfoXmlError::InvalidXml(msg) => {
                assert!(
                    msg.contains("payload too large"),
                    "unexpected error message: {msg}"
                );
            }
            other => panic!("unexpected error: {other:?}"),
        }
    }
}
