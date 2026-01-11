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
    /// Signature is present and the PKCS#7/CMS blob verifies successfully.
    ///
    /// Note: this verifies the CMS structure is internally consistent (the signature matches the
    /// embedded content / signed attributes). It does **not** currently verify that the signature
    /// content corresponds to the rest of the VBA project per MS-OVBA.
    SignedVerified,
    /// Signature stream exists and parses as CMS/PKCS#7, but verification failed.
    SignedInvalid,
    /// Signature stream exists, but does not parse as CMS/PKCS#7.
    SignedParseError,
    /// Signature is present but we did not validate it (legacy / reserved for future use).
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

/// Parse and cryptographically verify a VBA digital signature (if present).
///
/// Returns `Ok(None)` when the project appears unsigned.
///
/// Verification is "internal" CMS/PKCS#7 verification only: we validate that the signature blob
/// is well-formed and that the signature matches the signed attributes / embedded content. We do
/// **not** currently validate that the signature is bound to the rest of the VBA project streams
/// as Excel does per MS-OVBA.
pub fn verify_vba_digital_signature(
    vba_project_bin: &[u8],
) -> Result<Option<VbaDigitalSignature>, SignatureError> {
    let mut sig = match parse_vba_digital_signature(vba_project_bin)? {
        Some(sig) => sig,
        None => return Ok(None),
    };

    sig.verification = verify_signature_blob(&sig.signature);
    Ok(Some(sig))
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

fn verify_signature_blob(signature: &[u8]) -> VbaSignatureVerification {
    #[cfg(target_arch = "wasm32")]
    {
        let _ = signature;
        // `openssl` isn't available on wasm targets. Keep the signature blob available to callers,
        // but don't treat it as verified.
        return VbaSignatureVerification::SignedButUnverified;
    }

    #[cfg(not(target_arch = "wasm32"))]
    use openssl::pkcs7::{Pkcs7, Pkcs7Flags};
    #[cfg(not(target_arch = "wasm32"))]
    use openssl::stack::Stack;
    #[cfg(not(target_arch = "wasm32"))]
    use openssl::x509::store::X509StoreBuilder;

    #[cfg(not(target_arch = "wasm32"))]
    let pkcs7 = match Pkcs7::from_der(signature) {
        Ok(pkcs7) => pkcs7,
        Err(_) => return VbaSignatureVerification::SignedParseError,
    };

    #[cfg(not(target_arch = "wasm32"))]
    let store = match X509StoreBuilder::new() {
        Ok(builder) => builder.build(),
        Err(_) => return VbaSignatureVerification::SignedInvalid,
    };
    #[cfg(not(target_arch = "wasm32"))]
    let extra_certs = match Stack::new() {
        Ok(stack) => stack,
        Err(_) => return VbaSignatureVerification::SignedInvalid,
    };

    #[cfg(not(target_arch = "wasm32"))]
    {
        // NOVERIFY skips certificate chain verification. We still validate the signature itself and
        // any messageDigest attributes over the embedded content.
        let flags = Pkcs7Flags::NOVERIFY;
        match pkcs7.verify(&extra_certs, &store, None, None, flags) {
            Ok(()) => VbaSignatureVerification::SignedVerified,
            Err(_) => VbaSignatureVerification::SignedInvalid,
        }
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
