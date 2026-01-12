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
    #[error("DER parse error: {0}")]
    Der(String),
    #[error("PKCS#7 SignedData is detached, but no detached content was found")]
    DetachedContentMissing,
}

const OID_PKCS7_SIGNED_DATA: &str = "1.2.840.113549.1.7.2";

/// Extract the signed Authenticode file digest (the `DigestInfo` inside
/// `SpcIndirectDataContent`) from a raw VBA `\x05DigitalSignature*` stream.
///
/// Returns:
/// - `Ok(Some(_))` if a PKCS#7/CMS SignedData blob and `SpcIndirectDataContent` were found and parsed.
/// - `Ok(None)` if no PKCS#7 SignedData could be located in the stream.
pub fn extract_vba_signature_signed_digest(
    signature_stream: &[u8],
) -> Result<Option<VbaSignedDigest>, VbaSignatureSignedDigestError> {
    let Some(pkcs7) = locate_pkcs7_signed_data(signature_stream)? else {
        return Ok(None);
    };

    let encap = parse_pkcs7_signed_data_encap_content(pkcs7.der)?;

    let signed_content = if let Some(econtent) = encap.econtent {
        econtent
    } else if pkcs7.offset > 0 {
        signature_stream[..pkcs7.offset].to_vec()
    } else {
        return Err(VbaSignatureSignedDigestError::DetachedContentMissing);
    };

    let digest = parse_spc_indirect_data_content(&signed_content)?;
    Ok(Some(digest))
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

fn locate_pkcs7_signed_data<'a>(
    signature_stream: &'a [u8],
) -> Result<Option<Pkcs7Location<'a>>, VbaSignatureSignedDigestError> {
    // Fast path: begins with a DER SEQUENCE.
    if signature_stream.first() == Some(&0x30) {
        if let Some(loc) = try_locate_pkcs7_at(signature_stream, 0)? {
            return Ok(Some(loc));
        }
    }

    for offset in 0..signature_stream.len() {
        if signature_stream[offset] != 0x30 {
            continue;
        }
        if let Some(loc) = try_locate_pkcs7_at(signature_stream, offset)? {
            return Ok(Some(loc));
        }
    }
    Ok(None)
}

fn try_locate_pkcs7_at<'a>(
    bytes: &'a [u8],
    offset: usize,
) -> Result<Option<Pkcs7Location<'a>>, VbaSignatureSignedDigestError> {
    if offset >= bytes.len() || bytes[offset] != 0x30 {
        return Ok(None);
    }
    if offset + 2 > bytes.len() {
        return Ok(None);
    }

    let (len_len, content_len) = match parse_der_length(&bytes[offset + 1..]) {
        Ok(v) => v,
        Err(_) => return Ok(None),
    };

    let total_len = 1usize
        .checked_add(len_len)
        .and_then(|v| v.checked_add(content_len))
        .ok_or_else(|| VbaSignatureSignedDigestError::Der("length overflow".to_owned()))?;

    let end = match offset.checked_add(total_len) {
        Some(end) => end,
        None => return Ok(None),
    };
    if end > bytes.len() {
        return Ok(None);
    }

    let candidate = &bytes[offset..end];
    match is_pkcs7_signed_data_content_info(candidate) {
        Ok(true) => Ok(Some(Pkcs7Location {
            der: candidate,
            offset,
        })),
        Ok(false) => Ok(None),
        Err(_) => Ok(None),
    }
}

fn is_pkcs7_signed_data_content_info(der: &[u8]) -> Result<bool, VbaSignatureSignedDigestError> {
    let mut top = DerReader::new(der);
    let mut content_info = top.read_sequence()?;
    let oid = content_info.read_oid()?;
    if oid != OID_PKCS7_SIGNED_DATA {
        return Ok(false);
    }

    // ContentInfo.content is [0] EXPLICIT for SignedData.
    let mut explicit = match content_info.read_explicit(0) {
        Ok(v) => v,
        Err(_) => return Ok(false),
    };

    // Validate it looks like SignedData.
    let mut signed_data = match explicit.read_sequence() {
        Ok(v) => v,
        Err(_) => return Ok(false),
    };
    // version INTEGER
    if signed_data.peek_tag() != Some(0x02) {
        return Ok(false);
    }
    signed_data.skip_any()?;
    // digestAlgorithms SET
    if signed_data.peek_tag() != Some(0x31) {
        return Ok(false);
    }
    signed_data.skip_any()?;
    // encapContentInfo SEQUENCE
    if signed_data.peek_tag() != Some(0x30) {
        return Ok(false);
    }
    Ok(true)
}

fn parse_pkcs7_signed_data_encap_content(
    pkcs7_der: &[u8],
) -> Result<Pkcs7EncapsulatedContent, VbaSignatureSignedDigestError> {
    let mut top = DerReader::new(pkcs7_der);
    let mut content_info = top.read_sequence()?;
    let oid = content_info.read_oid()?;
    if oid != OID_PKCS7_SIGNED_DATA {
        return Err(VbaSignatureSignedDigestError::Der(format!(
            "expected PKCS#7 signedData ContentInfo ({}), got {}",
            OID_PKCS7_SIGNED_DATA, oid
        )));
    }

    let mut explicit = content_info.read_explicit(0)?;
    let mut signed_data = explicit.read_sequence()?;

    // version INTEGER
    signed_data.skip_any()?;
    // digestAlgorithms SET OF AlgorithmIdentifier
    signed_data.skip_any()?;

    // encapContentInfo
    let mut encap = signed_data.read_sequence()?;
    let econtent_type_oid = encap.read_oid()?;

    let econtent = if encap.is_empty() {
        None
    } else if encap.peek_tag() == Some(explicit_tag(0)) {
        let mut econtent_explicit = encap.read_explicit(0)?;
        let octets = econtent_explicit.read_octet_string()?;
        Some(octets.to_vec())
    } else {
        return Err(VbaSignatureSignedDigestError::Der(format!(
            "unexpected EncapsulatedContentInfo field tag 0x{:02x}",
            encap.peek_tag().unwrap_or(0)
        )));
    };

    Ok(Pkcs7EncapsulatedContent {
        econtent_type_oid,
        econtent,
    })
}

fn parse_spc_indirect_data_content(
    der: &[u8],
) -> Result<VbaSignedDigest, VbaSignatureSignedDigestError> {
    let mut top = DerReader::new(der);
    let mut seq = top.read_sequence()?;

    // data SpcAttributeTypeAndOptionalValue (ignored)
    seq.skip_any()?;

    // messageDigest DigestInfo
    let mut digest_info = seq.read_sequence()?;
    let mut alg = digest_info.read_sequence()?;
    let digest_algorithm_oid = alg.read_oid()?;
    // Optional parameters (often NULL) - ignore.
    while !alg.is_empty() {
        alg.skip_any()?;
    }
    let digest = digest_info.read_octet_string()?.to_vec();

    Ok(VbaSignedDigest {
        digest_algorithm_oid,
        digest,
    })
}

#[derive(Debug, Clone, Copy)]
struct DerTlv<'a> {
    tag: u8,
    value: &'a [u8],
}

#[derive(Debug, Clone, Copy)]
struct DerReader<'a> {
    bytes: &'a [u8],
    pos: usize,
}

impl<'a> DerReader<'a> {
    fn new(bytes: &'a [u8]) -> Self {
        Self { bytes, pos: 0 }
    }

    fn is_empty(&self) -> bool {
        self.pos >= self.bytes.len()
    }

    fn peek_tag(&self) -> Option<u8> {
        self.bytes.get(self.pos).copied()
    }

    fn read_tlv(&mut self) -> Result<DerTlv<'a>, VbaSignatureSignedDigestError> {
        let tag = *self
            .bytes
            .get(self.pos)
            .ok_or_else(|| VbaSignatureSignedDigestError::Der("unexpected EOF".to_owned()))?;
        if tag & 0x1F == 0x1F {
            return Err(VbaSignatureSignedDigestError::Der(
                "high-tag-number form not supported".to_owned(),
            ));
        }
        self.pos += 1;

        let (len_len, len) = parse_der_length(&self.bytes[self.pos..])?;
        self.pos += len_len;
        let end = self
            .pos
            .checked_add(len)
            .ok_or_else(|| VbaSignatureSignedDigestError::Der("length overflow".to_owned()))?;
        if end > self.bytes.len() {
            return Err(VbaSignatureSignedDigestError::Der(
                "DER length exceeds input".to_owned(),
            ));
        }
        let value = &self.bytes[self.pos..end];
        self.pos = end;
        Ok(DerTlv { tag, value })
    }

    fn read_expected(
        &mut self,
        expected_tag: u8,
    ) -> Result<DerTlv<'a>, VbaSignatureSignedDigestError> {
        let tlv = self.read_tlv()?;
        if tlv.tag != expected_tag {
            return Err(VbaSignatureSignedDigestError::Der(format!(
                "expected tag 0x{:02x}, got 0x{:02x}",
                expected_tag, tlv.tag
            )));
        }
        Ok(tlv)
    }

    fn read_sequence(&mut self) -> Result<DerReader<'a>, VbaSignatureSignedDigestError> {
        let tlv = self.read_expected(0x30)?;
        Ok(DerReader::new(tlv.value))
    }

    fn read_explicit(
        &mut self,
        tag_no: u8,
    ) -> Result<DerReader<'a>, VbaSignatureSignedDigestError> {
        let tlv = self.read_expected(explicit_tag(tag_no))?;
        Ok(DerReader::new(tlv.value))
    }

    fn read_oid(&mut self) -> Result<String, VbaSignatureSignedDigestError> {
        let tlv = self.read_expected(0x06)?;
        decode_der_oid(tlv.value)
    }

    fn read_octet_string(&mut self) -> Result<&'a [u8], VbaSignatureSignedDigestError> {
        let tlv = self.read_expected(0x04)?;
        Ok(tlv.value)
    }

    fn skip_any(&mut self) -> Result<(), VbaSignatureSignedDigestError> {
        let _ = self.read_tlv()?;
        Ok(())
    }
}

fn explicit_tag(tag_no: u8) -> u8 {
    0xA0u8 | (tag_no & 0x1F)
}

fn parse_der_length(input: &[u8]) -> Result<(usize, usize), VbaSignatureSignedDigestError> {
    let first = *input
        .get(0)
        .ok_or_else(|| VbaSignatureSignedDigestError::Der("unexpected EOF".to_owned()))?;
    if first & 0x80 == 0 {
        return Ok((1, first as usize));
    }
    let num_bytes = (first & 0x7F) as usize;
    if num_bytes == 0 {
        return Err(VbaSignatureSignedDigestError::Der(
            "indefinite length form is not supported".to_owned(),
        ));
    }
    if num_bytes > 8 {
        return Err(VbaSignatureSignedDigestError::Der(
            "length too large".to_owned(),
        ));
    }
    if input.len() < 1 + num_bytes {
        return Err(VbaSignatureSignedDigestError::Der(
            "unexpected EOF parsing length".to_owned(),
        ));
    }
    let mut len: usize = 0;
    for &b in &input[1..1 + num_bytes] {
        len = (len << 8) | (b as usize);
    }
    Ok((1 + num_bytes, len))
}

fn decode_der_oid(bytes: &[u8]) -> Result<String, VbaSignatureSignedDigestError> {
    if bytes.is_empty() {
        return Err(VbaSignatureSignedDigestError::Der(
            "OID has empty value".to_owned(),
        ));
    }
    let first = bytes[0];
    let first_arc = (first / 40) as u32;
    let second_arc = (first % 40) as u32;

    let mut arcs = vec![first_arc, second_arc];
    let mut value: u32 = 0;
    let mut in_arc = false;

    for &b in &bytes[1..] {
        in_arc = true;
        value = value
            .checked_shl(7)
            .and_then(|v| v.checked_add((b & 0x7F) as u32))
            .ok_or_else(|| VbaSignatureSignedDigestError::Der("OID arc overflow".to_owned()))?;
        if b & 0x80 == 0 {
            arcs.push(value);
            value = 0;
            in_arc = false;
        }
    }

    if in_arc {
        return Err(VbaSignatureSignedDigestError::Der(
            "OID has truncated base128 arc".to_owned(),
        ));
    }

    Ok(arcs
        .iter()
        .map(|n| n.to_string())
        .collect::<Vec<_>>()
        .join("."))
}
