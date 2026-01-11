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
///
/// If multiple signature streams are present, we return the first one (by Excel's preferred stream
/// name ordering) that verifies successfully, falling back to the first candidate if none verify.
pub fn verify_vba_digital_signature(
    vba_project_bin: &[u8],
) -> Result<Option<VbaDigitalSignature>, SignatureError> {
    let mut ole = OleFile::open(vba_project_bin)?;
    let streams = ole.list_streams()?;

    let mut candidates = streams
        .into_iter()
        .filter(|path| path.split('/').any(is_signature_component))
        .collect::<Vec<_>>();

    if candidates.is_empty() {
        return Ok(None);
    }

    candidates.sort_by(|a, b| signature_path_rank(a).cmp(&signature_path_rank(b)).then(a.cmp(b)));

    let mut first: Option<VbaDigitalSignature> = None;
    for path in candidates {
        let signature = ole.read_stream_opt(&path)?.unwrap_or_default();
        let signer_subject = extract_first_certificate_subject(&signature);
        let verification = verify_signature_blob(&signature);
        let sig = VbaDigitalSignature {
            stream_path: path,
            signer_subject,
            signature,
            verification,
        };
        if sig.verification == VbaSignatureVerification::SignedVerified {
            return Ok(Some(sig));
        }
        if first.is_none() {
            first = Some(sig);
        }
    }

    Ok(first)
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
    path.split('/')
        .map(|component| {
            let trimmed = component.trim_start_matches(|c: char| c <= '\u{001F}');
            match trimmed {
                "DigitalSignature" => 0,
                "DigitalSignatureEx" => 1,
                "DigitalSignatureExt" => 2,
                _ => 3,
            }
        })
        .min()
        .unwrap_or(3)
}

#[cfg(not(target_arch = "wasm32"))]
fn parse_pkcs7_with_offset(signature: &[u8]) -> Option<(openssl::pkcs7::Pkcs7, usize)> {
    use openssl::pkcs7::Pkcs7;

    // Some producers include a small header before the DER-encoded PKCS#7 payload. Try parsing
    // from the start first, then scan for an embedded DER SEQUENCE that parses as PKCS#7.
    if let Ok(pkcs7) = Pkcs7::from_der(signature) {
        return Some((pkcs7, 0));
    }

    for start in 0..signature.len() {
        if signature[start] != 0x30 {
            continue;
        }
        if let Ok(pkcs7) = Pkcs7::from_der(&signature[start..]) {
            return Some((pkcs7, start));
        }
    }

    None
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
    use openssl::pkcs7::Pkcs7Flags;
    #[cfg(not(target_arch = "wasm32"))]
    use openssl::stack::Stack;
    #[cfg(not(target_arch = "wasm32"))]
    use openssl::x509::X509;
    #[cfg(not(target_arch = "wasm32"))]
    use openssl::x509::store::X509StoreBuilder;

    #[cfg(not(target_arch = "wasm32"))]
    let Some((pkcs7, pkcs7_offset)) = parse_pkcs7_with_offset(signature) else {
        return VbaSignatureVerification::SignedParseError;
    };

    #[cfg(not(target_arch = "wasm32"))]
    let store = match X509StoreBuilder::new() {
        Ok(builder) => builder.build(),
        Err(_) => return VbaSignatureVerification::SignedInvalid,
    };
    #[cfg(not(target_arch = "wasm32"))]
    let empty_certs = match Stack::<X509>::new() {
        Ok(stack) => stack,
        Err(_) => return VbaSignatureVerification::SignedInvalid,
    };
    #[cfg(not(target_arch = "wasm32"))]
    let certs = pkcs7
        .signed()
        .and_then(|s| s.certificates())
        .unwrap_or(&empty_certs);

    #[cfg(not(target_arch = "wasm32"))]
    {
        // NOVERIFY skips certificate chain verification. We still validate the signature itself and
        // any messageDigest attributes over the embedded content.
        let flags = Pkcs7Flags::NOVERIFY;

        // First try verifying as a "normal" PKCS#7 blob with embedded content.
        if pkcs7.verify(certs, &store, None, None, flags).is_ok() {
            return VbaSignatureVerification::SignedVerified;
        }

        // If the PKCS#7 blob was found after a prefix/header, try treating the prefix as detached
        // content. This matches common patterns where the stream contains signed bytes followed by
        // a detached PKCS#7 signature over those bytes.
        if pkcs7_offset > 0 {
            let prefix = &signature[..pkcs7_offset];
            if pkcs7
                .verify(
                    certs,
                    &store,
                    Some(prefix),
                    None,
                    flags | Pkcs7Flags::DETACHED,
                )
                .is_ok()
                || pkcs7.verify(certs, &store, Some(prefix), None, flags).is_ok()
            {
                return VbaSignatureVerification::SignedVerified;
            }
        }

        VbaSignatureVerification::SignedInvalid
    }
}

#[cfg(not(target_arch = "wasm32"))]
fn extract_signer_subject_from_pkcs7(bytes: &[u8]) -> Option<String> {
    use openssl::pkcs7::Pkcs7Flags;
    use openssl::stack::Stack;
    use openssl::x509::X509;
    use x509_parser::prelude::parse_x509_certificate;

    let (pkcs7, _) = parse_pkcs7_with_offset(bytes)?;
    let signers = if let Some(certs) = pkcs7.signed().and_then(|s| s.certificates()) {
        pkcs7.signers(certs, Pkcs7Flags::empty()).ok()
    } else {
        let empty = Stack::<X509>::new().ok()?;
        pkcs7.signers(&empty, Pkcs7Flags::empty()).ok()
    }?;

    let signer_cert = signers.get(0)?;
    let der = signer_cert.to_der().ok()?;
    let (_, cert) = parse_x509_certificate(&der).ok()?;
    Some(cert.subject().to_string())
}

fn extract_first_certificate_subject(bytes: &[u8]) -> Option<String> {
    use x509_parser::prelude::parse_x509_certificate;

    // Try parsing from the beginning first (some producers store a raw cert).
    if let Ok((_, cert)) = parse_x509_certificate(bytes) {
        return Some(cert.subject().to_string());
    }

    #[cfg(not(target_arch = "wasm32"))]
    if let Some(subject) = extract_signer_subject_from_pkcs7(bytes) {
        return Some(subject);
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
