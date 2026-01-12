use thiserror::Error;

use crate::{
    authenticode::extract_vba_signature_signed_digest,
    project_digest::compute_vba_project_digest,
    project_digest::DigestAlg,
    OleError, OleFile,
};

/// Metadata extracted from the signer certificate embedded in a VBA digital signature.
///
/// This is intended for UI display (e.g. Trust Center) and is **best-effort**:
/// callers should treat it as untrusted metadata unless they separately validate
/// the signature and certificate chain.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct VbaSignerCertificateInfo {
    /// Certificate subject (e.g. `CN=...`).
    pub subject: String,
    /// Certificate issuer (e.g. `CN=...`).
    pub issuer: String,
    /// Serial number encoded as lowercase hexadecimal (no `0x` prefix).
    pub serial_hex: String,
    /// SHA-256 fingerprint of the DER-encoded certificate, lowercase hex.
    pub sha256_fingerprint_hex: String,
    /// Validity period start time (best-effort).
    pub not_before: Option<String>,
    /// Validity period end time (best-effort).
    pub not_after: Option<String>,
}

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
    /// Whether the signature is bound to the VBA project streams via the MS-OVBA "project digest"
    /// mechanism.
    pub binding: VbaSignatureBinding,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VbaSignatureBinding {
    /// The signature's signed digest matches the computed digest of the project streams.
    Bound,
    /// We extracted the signed digest and computed a project digest, but they do not match.
    NotBound,
    /// Binding could not be verified (unsupported/unknown format, missing data, or signature not
    /// cryptographically verified).
    Unknown,
}

/// Result of inspecting an individual VBA digital signature stream.
///
/// Unlike [`VbaDigitalSignature`] (which is used by the single-signature helper APIs),
/// this type is designed for enumerating *all* signature streams that may be present in
/// a `vbaProject.bin`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct VbaDigitalSignatureStream {
    /// OLE stream path the signature was loaded from.
    pub stream_path: String,
    /// Best-effort signer certificate subject (e.g. `CN=...`), if found.
    pub signer_subject: Option<String>,
    /// Raw signature stream bytes.
    pub signature: Vec<u8>,
    /// Verification state (best-effort).
    pub verification: VbaSignatureVerification,
    /// Best-effort DigestInfo algorithm OID extracted from Authenticode's
    /// `SpcIndirectDataContent` (if present).
    pub signed_digest_algorithm_oid: Option<String>,
    /// Best-effort DigestInfo digest extracted from Authenticode's
    /// `SpcIndirectDataContent` (if present).
    pub signed_digest: Option<Vec<u8>>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VbaSignatureVerification {
    /// Signature is present and the PKCS#7/CMS blob verifies successfully.
    ///
    /// Note: this verifies the CMS structure is internally consistent (the signature matches the
    /// embedded content / signed attributes). Use [`VbaDigitalSignature::binding`] to determine
    /// whether the signature is bound to the VBA project streams per MS-OVBA.
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

/// Extract metadata from the signer certificate embedded in a VBA signature blob.
///
/// This is a best-effort helper intended for display:
/// - Returns `None` if we can't find an embedded certificate.
/// - Does **not** validate the certificate chain.
/// - On non-wasm targets we prefer extracting the actual signer cert via OpenSSL PKCS#7 parsing.
/// - On wasm (or as a fallback) we heuristically scan for an embedded DER certificate using
///   `x509-parser`.
pub fn extract_signer_certificate_info(signature_blob: &[u8]) -> Option<VbaSignerCertificateInfo> {
    #[cfg(not(target_arch = "wasm32"))]
    if let Some(der) = extract_signer_certificate_der_from_pkcs7(signature_blob) {
        if let Some(info) = signer_certificate_info_from_der(&der) {
            return Some(info);
        }
    }

    // Some producers store a raw DER certificate in the signature stream.
    if let Some(der) = parse_first_embedded_der_certificate(signature_blob) {
        return signer_certificate_info_from_der(der);
    }

    // Fallback: scan for embedded certificates inside a CMS/PKCS#7 blob.
    for der in scan_for_embedded_der_certificates(signature_blob) {
        if let Some(info) = signer_certificate_info_from_der(der) {
            return Some(info);
        }
    }

    None
}

/// Enumerate and inspect *all* VBA digital signature streams found in a `vbaProject.bin`.
///
/// Excel stores VBA project signatures in one of the `\u{0005}DigitalSignature*` streams
/// (see MS-OVBA). Some files may contain multiple signature streams (e.g. legacy and
/// SHA-2-era formats). This API exposes each stream independently.
///
/// Streams are returned in deterministic Excel-like order (newest stream first; see
/// `signature_path_rank`).
pub fn list_vba_digital_signatures(
    vba_project_bin: &[u8],
) -> Result<Vec<VbaDigitalSignatureStream>, SignatureError> {
    let mut ole = OleFile::open(vba_project_bin)?;
    let streams = ole.list_streams()?;

    let mut candidates = streams
        .into_iter()
        .filter(|path| path.split('/').any(is_signature_component))
        .collect::<Vec<_>>();

    candidates.sort_by(|a, b| signature_path_rank(a).cmp(&signature_path_rank(b)).then(a.cmp(b)));

    let mut out = Vec::new();
    for path in candidates {
        let signature = ole.read_stream_opt(&path)?.unwrap_or_default();
        let signer_subject = extract_first_certificate_subject(&signature);
        let verification = verify_signature_blob(&signature);

        let (signed_digest_algorithm_oid, signed_digest) =
            match crate::authenticode::extract_vba_signature_signed_digest(&signature) {
                Ok(Some(digest)) => (Some(digest.digest_algorithm_oid), Some(digest.digest)),
                Ok(None) | Err(_) => (None, None),
            };

        out.push(VbaDigitalSignatureStream {
            stream_path: path,
            signer_subject,
            signature,
            verification,
            signed_digest_algorithm_oid,
            signed_digest,
        });
    }

    Ok(out)
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

    // Prefer the signature stream that Excel would treat as authoritative when more than one
    // signature stream exists (see `signature_path_rank` for the MS-OVBA-defined ordering).
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
        binding: VbaSignatureBinding::Unknown,
    }))
}

/// Parse and cryptographically verify a VBA digital signature (if present).
///
/// Returns `Ok(None)` when the project appears unsigned.
///
/// Verification is "internal" CMS/PKCS#7 verification only: we validate that the signature blob
/// is well-formed and that the signature matches the signed attributes / embedded content. We do
/// best-effort binding verification by extracting the signed project digest (Authenticode-style
/// `SpcIndirectDataContent`) and comparing it to a freshly computed digest over the project's OLE
/// streams (excluding any signature streams).
///
/// If multiple signature streams are present, we return the first one (by Excel's preferred stream
/// name ordering; see `signature_path_rank`) that verifies successfully, falling back to the first
/// candidate if none verify.
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
        let binding = match verification {
            VbaSignatureVerification::SignedVerified => {
                verify_signature_binding(vba_project_bin, &signature)
            }
            _ => VbaSignatureBinding::Unknown,
        };
        let sig = VbaDigitalSignature {
            stream_path: path,
            signer_subject,
            signature,
            verification,
            binding,
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

fn verify_signature_binding(vba_project_bin: &[u8], signature: &[u8]) -> VbaSignatureBinding {
    #[cfg(target_arch = "wasm32")]
    {
        let _ = (vba_project_bin, signature);
        return VbaSignatureBinding::Unknown;
    }

    #[cfg(not(target_arch = "wasm32"))]
    {
        let signed = match extract_vba_signature_signed_digest(signature) {
            Ok(Some(v)) => v,
            _ => return VbaSignatureBinding::Unknown,
        };

        let Some(alg) = digest_alg_from_oid_str(&signed.digest_algorithm_oid) else {
            return VbaSignatureBinding::Unknown;
        };

        let Ok(computed) = compute_vba_project_digest(vba_project_bin, alg) else {
            return VbaSignatureBinding::Unknown;
        };

        if computed == signed.digest {
            VbaSignatureBinding::Bound
        } else {
            VbaSignatureBinding::NotBound
        }
    }
}

fn digest_alg_from_oid_str(oid: &str) -> Option<DigestAlg> {
    match oid {
        "1.3.14.3.2.26" => Some(DigestAlg::Sha1),
        "2.16.840.1.101.3.4.2.1" => Some(DigestAlg::Sha256),
        _ => None,
    }
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
    //
    // MS-OVBA defines three possible signature streams that can appear in a VBA project storage:
    // - "\x05DigitalSignature"     (legacy; SHA-1 era)
    // - "\x05DigitalSignatureEx"   (extended; SHA-2 era)
    // - "\x05DigitalSignatureExt"  (extension; newest format)
    //
    // When more than one exists, Office apps (including Excel) prefer the newest stream.
    //
    // Spec: MS-OVBA "Project Signature" streams (§2.3.4.1–2.3.4.3).
    // https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/
    path.split('/')
        .map(|component| {
            let trimmed = component.trim_start_matches(|c: char| c <= '\u{001F}');
            match trimmed {
                "DigitalSignatureExt" => 0,
                "DigitalSignatureEx" => 1,
                "DigitalSignature" => 2,
                _ => 3,
            }
        })
        .min()
        .unwrap_or(3)
}

fn bytes_to_lower_hex(bytes: &[u8]) -> String {
    let mut out = String::with_capacity(bytes.len() * 2);
    for b in bytes {
        use std::fmt::Write;
        // Safe to unwrap: writing to a String cannot fail.
        write!(&mut out, "{:02x}", b).expect("writing to string is infallible");
    }
    out
}

fn signer_certificate_info_from_der(der: &[u8]) -> Option<VbaSignerCertificateInfo> {
    use sha2::{Digest, Sha256};
    use x509_parser::prelude::parse_x509_certificate;

    let (_, cert) = parse_x509_certificate(der).ok()?;

    let subject = cert.subject().to_string();
    let issuer = cert.issuer().to_string();

    // Prefer the raw serial bytes (stable, includes leading zeros if present).
    // If the parser ever returns an empty serial, normalize to "00".
    let serial_bytes = cert.raw_serial();
    let serial_hex = if serial_bytes.is_empty() {
        "00".to_owned()
    } else {
        bytes_to_lower_hex(serial_bytes)
    };

    let sha256 = Sha256::digest(der);
    let sha256_fingerprint_hex = bytes_to_lower_hex(&sha256);

    let validity = cert.validity();
    let not_before = Some(validity.not_before.to_string());
    let not_after = Some(validity.not_after.to_string());

    Some(VbaSignerCertificateInfo {
        subject,
        issuer,
        serial_hex,
        sha256_fingerprint_hex,
        not_before,
        not_after,
    })
}

fn parse_first_embedded_der_certificate(bytes: &[u8]) -> Option<&[u8]> {
    use x509_parser::prelude::parse_x509_certificate;

    let (rem, _) = parse_x509_certificate(bytes).ok()?;
    let consumed_len = bytes.len().saturating_sub(rem.len());
    Some(&bytes[..consumed_len])
}

fn scan_for_embedded_der_certificates(bytes: &[u8]) -> impl Iterator<Item = &[u8]> {
    use x509_parser::prelude::parse_x509_certificate;

    // Heuristic: certificates are DER-encoded and begin with a SEQUENCE (0x30) tag.
    // Yield each candidate DER slice that parses as a certificate.
    (0..bytes.len()).filter_map(move |start| {
        if bytes[start] != 0x30 {
            return None;
        }
        let slice = &bytes[start..];
        let (rem, _) = parse_x509_certificate(slice).ok()?;
        let consumed_len = slice.len().saturating_sub(rem.len());
        if consumed_len == 0 {
            return None;
        }
        Some(&slice[..consumed_len])
    })
}

#[cfg(not(target_arch = "wasm32"))]
fn parse_pkcs7_with_offset(signature: &[u8]) -> Option<(openssl::pkcs7::Pkcs7, usize)> {
    use openssl::pkcs7::Pkcs7;

    // Some producers include a small header before the DER-encoded PKCS#7 payload. Try parsing
    // from the start first, then scan for an embedded DER SEQUENCE that parses as PKCS#7.
    if let Ok(pkcs7) = Pkcs7::from_der(signature) {
        return Some((pkcs7, 0));
    }

    // Office commonly wraps the PKCS#7 blob in a [MS-OFFCRYPTO] DigSigInfoSerialized structure.
    // Parsing the header is deterministic and avoids the worst-case behavior of scanning/parsing
    // from every 0x30 offset.
    if let Some(info) = crate::offcrypto::parse_digsig_info_serialized(signature) {
        let end = info.pkcs7_offset.saturating_add(info.pkcs7_len);
        if end <= signature.len() {
            let slice = &signature[info.pkcs7_offset..end];
            if let Ok(pkcs7) = Pkcs7::from_der(slice) {
                return Some((pkcs7, info.pkcs7_offset));
            }
        }
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

#[cfg(not(target_arch = "wasm32"))]
fn extract_signer_certificate_der_from_pkcs7(bytes: &[u8]) -> Option<Vec<u8>> {
    use openssl::pkcs7::Pkcs7Flags;
    use openssl::stack::Stack;
    use openssl::x509::X509;

    let (pkcs7, _) = parse_pkcs7_with_offset(bytes)?;

    // Prefer the actual signer cert when possible.
    if let Some(certs) = pkcs7.signed().and_then(|s| s.certificates()) {
        if let Ok(signers) = pkcs7.signers(certs, Pkcs7Flags::empty()) {
            if let Some(signer_cert) = signers.get(0) {
                return signer_cert.to_der().ok();
            }
        }

        // Fallback: first embedded certificate.
        if let Some(cert) = certs.get(0) {
            return cert.to_der().ok();
        }
    }

    // Some PKCS#7 blobs may omit the certificates stack or require a different lookup. As a
    // last attempt, run `signers` with an empty stack (OpenSSL will try to resolve signer info).
    let empty = Stack::<X509>::new().ok()?;
    let signers = pkcs7.signers(&empty, Pkcs7Flags::empty()).ok()?;
    let signer_cert = signers.get(0)?;
    signer_cert.to_der().ok()
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
        // BINARY avoids any canonicalization (e.g. newline conversions) when verifying.
        let flags = Pkcs7Flags::NOVERIFY | Pkcs7Flags::BINARY;

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
