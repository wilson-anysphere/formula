use thiserror::Error;

use crate::{OleError, OleFile};

/// Result of inspecting a VBA project's OLE structure for a digital signature.
///
/// Excel stores the VBA project signature in one of the `\u{0005}DigitalSignature*`
/// streams (see MS-OVBA). The stream contents are typically an Authenticode-like
/// PKCS#7/CMS structure that embeds the signing certificate.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct VbaDigitalSignature {
    /// OLE stream path the signature was loaded from.
    pub stream_path: String,
    /// Best-effort signer certificate subject (e.g. `CN=...`), if found.
    pub signer_subject: Option<String>,
    /// Raw signature stream bytes.
    pub signature: Vec<u8>,
    /// Verification state (best-effort).
    pub verification: VbaSignatureVerification,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VbaSignatureVerification {
    /// Signature is present but we did not (yet) validate it against the VBA project contents.
    SignedButUnverified,
}

#[derive(Debug, Error)]
pub enum SignatureError {
    #[error("OLE error: {0}")]
    Ole(#[from] OleError),
}

/// Best-effort detection + parsing of a VBA digital signature.
///
/// Returns `Ok(None)` when the project appears unsigned.
pub fn parse_vba_digital_signature(
    vba_project_bin: &[u8],
) -> Result<Option<VbaDigitalSignature>, SignatureError> {
    let mut ole = OleFile::open(vba_project_bin)?;
    let streams = ole.list_streams()?;

    // Excel/VBA signature streams are control-character prefixed in OLE:
    // - "\x05DigitalSignature"
    // - "\x05DigitalSignatureEx"
    // - "\x05DigitalSignatureExt"
    //
    // They may appear as either a stream or a storage containing a stream; we
    // match on any path component to be tolerant.
    let mut candidates = streams
        .into_iter()
        .filter(|path| path.split('/').any(is_signature_component))
        .collect::<Vec<_>>();

    if candidates.is_empty() {
        return Ok(None);
    }

    // Prefer the canonical stream name when present.
    candidates.sort_by(|a, b| signature_path_rank(a).cmp(&signature_path_rank(b)).then(a.cmp(b)));
    let chosen = candidates
        .into_iter()
        .next()
        .expect("candidates non-empty");
    let signature = ole
        .read_stream_opt(&chosen)?
        .unwrap_or_else(|| Vec::new());

    let signer_subject = extract_first_certificate_subject(&signature);

    Ok(Some(VbaDigitalSignature {
        stream_path: chosen,
        signer_subject,
        signature,
        verification: VbaSignatureVerification::SignedButUnverified,
    }))
}

fn is_signature_component(component: &str) -> bool {
    let trimmed = component.trim_start_matches(|c: char| c <= '\u{001F}');
    matches!(
        trimmed,
        "DigitalSignature" | "DigitalSignatureEx" | "DigitalSignatureExt"
    )
}

fn signature_path_rank(path: &str) -> u8 {
    // Lower = higher priority.
    let last = path.split('/').last().unwrap_or(path);
    let trimmed = last.trim_start_matches(|c: char| c <= '\u{001F}');
    match trimmed {
        "DigitalSignature" => 0,
        "DigitalSignatureEx" => 1,
        "DigitalSignatureExt" => 2,
        _ => 3,
    }
}

fn extract_first_certificate_subject(bytes: &[u8]) -> Option<String> {
    use x509_parser::prelude::parse_x509_certificate;

    // Try parsing from the beginning first (some producers store a raw cert).
    if let Ok((_, cert)) = parse_x509_certificate(bytes) {
        return Some(cert.subject().to_string());
    }

    // Otherwise, scan for embedded certificates inside a CMS/PKCS#7 blob. This
    // is a best-effort heuristic: certificates are DER-encoded and begin with a
    // SEQUENCE (0x30) tag.
    for start in 0..bytes.len() {
        if bytes[start] != 0x30 {
            continue;
        }
        if let Ok((_, cert)) = parse_x509_certificate(&bytes[start..]) {
            return Some(cert.subject().to_string());
        }
    }

    None
}

