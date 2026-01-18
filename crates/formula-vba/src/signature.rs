use thiserror::Error;

use crate::{
    authenticode::extract_vba_signature_signed_digest,
    contents_hash::content_normalized_data,
    normalized_data::forms_normalized_data,
    OleError,
    OleFile,
};

use md5::{Digest as _, Md5};
/// Identifies which `\x05DigitalSignature*` stream/storage variant a signature was loaded from.
///
/// Excel stores VBA project signatures in one of three known variants:
/// - `\x05DigitalSignature` (legacy)
/// - `\x05DigitalSignatureEx` (extended, SHA-2-era)
/// - `\x05DigitalSignatureExt` (extension; newest / used for `ContentsHashV3` binding)
///
/// Some producers store the signature stream inside a storage (e.g. `\x05DigitalSignatureEx/sig`);
/// detection therefore inspects all OLE path components.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VbaSignatureStreamKind {
    DigitalSignature,
    DigitalSignatureEx,
    DigitalSignatureExt,
    /// The stream path looks signature-related (e.g. component begins with `DigitalSignature` after
    /// trimming leading C0 control chars) but doesn't exactly match a known variant name.
    Unknown,
}

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
/// Excel stores the VBA project signature in one of the `\u{0005}DigitalSignature*` streams.
///
/// The stream contents are typically an Authenticode-like PKCS#7/CMS structure that embeds the
/// signing certificate and a signed digest binding the signature to the VBA project's OLE streams.
/// MS-OVBA specifies the digest computation ("Contents Hash" §2.4.2), but does not normatively
/// specify the stream naming/precedence rules between the `DigitalSignature*` variants.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct VbaDigitalSignature {
    /// OLE stream path the signature was loaded from.
    pub stream_path: String,
    /// Which `DigitalSignature*` variant this signature stream corresponds to.
    pub stream_kind: VbaSignatureStreamKind,
    /// Best-effort signer certificate subject (e.g. `CN=...`), if found.
    pub signer_subject: Option<String>,
    /// Raw signature stream bytes.
    pub signature: Vec<u8>,
    /// Verification state (best-effort).
    pub verification: VbaSignatureVerification,
    /// Whether the signature is bound to the VBA project via the MS-OVBA digest ("Contents Hash")
    /// mechanism (`ContentsHashV3` for `DigitalSignatureExt`).
    ///
    /// Note: per MS-OSHARED §4.3, Office uses **16-byte MD5** binding digest bytes for legacy VBA
    /// signature streams (`\x05DigitalSignature` / `\x05DigitalSignatureEx`) even when
    /// `DigestInfo.digestAlgorithm.algorithm` indicates SHA-256.
    ///
    /// For the v3 `\x05DigitalSignatureExt` variant, binding uses the MS-OVBA v3 content-hash
    /// transcript (MS-OVBA §2.4.2.7). In the wild, the signed digest bytes are commonly 32-byte
    /// SHA-256, but producers can vary.
    ///
    /// The `DigestInfo` algorithm OID is not authoritative for binding (some producers emit
    /// inconsistent OIDs); `formula-vba` compares digest bytes to [`crate::contents_hash_v3`].
    pub binding: VbaSignatureBinding,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VbaSignatureBinding {
    /// The signature's signed digest matches the computed MS-OVBA Contents Hash for the project
    /// (Content Hash / Agile Content Hash, or `ContentsHashV3` for `DigitalSignatureExt`).
    Bound,
    /// We extracted the signed digest and computed the relevant MS-OVBA digests, but they do not
    /// match.
    NotBound,
    /// Binding could not be verified (unsupported/unknown format, missing data, or signature not
    /// cryptographically verified).
    Unknown,
}

/// Result of verifying a VBA digital signature *and* verifying that the signature is bound to the
/// current VBA project via the MS-OVBA Contents Hash (`ContentsHashV3` for `DigitalSignatureExt`).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct VbaDigitalSignatureBound {
    /// The parsed signature stream, including PKCS#7 verification result.
    pub signature: VbaDigitalSignature,
    /// Contents Hash binding verification result.
    pub binding: VbaProjectBindingVerification,
}

/// Best-effort debug information for MS-OVBA Contents Hash binding verification.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
    pub struct VbaProjectDigestDebugInfo {
    /// OID from `DigestInfo.digestAlgorithm.algorithm` (if found).
    ///
    /// Note: for legacy VBA signature streams (`DigitalSignature` / `DigitalSignatureEx`), this OID
    /// is not authoritative for binding; Office uses 16-byte MD5 digest bytes per MS-OSHARED §4.3
    /// even when the OID indicates SHA-256. For v3 (`DigitalSignatureExt`), the OID is surfaced for
    /// debugging/UI display but is not authoritative for binding (some producers emit inconsistent
    /// OIDs).
    pub hash_algorithm_oid: Option<String>,
    /// Human-readable name for the hash algorithm (best-effort).
    pub hash_algorithm_name: Option<String>,
    /// Digest bytes extracted from the signed Authenticode `SpcIndirectDataContent` (if found).
    pub signed_digest: Option<Vec<u8>>,
    /// Digest bytes computed from the current `vbaProject.bin` (if computed).
    pub computed_digest: Option<Vec<u8>>,
}

/// Binding verification result for the MS-OVBA Contents Hash.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum VbaProjectBindingVerification {
    /// Signature is present and the signed digest matches the digest computed from the current
    /// project contents.
    BoundVerified(VbaProjectDigestDebugInfo),
    /// Signature is present and parses, but the signed digest does not match the current project
    /// contents.
    BoundMismatch(VbaProjectDigestDebugInfo),
    /// Binding could not be verified (unsupported algorithm, missing digest structure, etc).
    BoundUnknown(VbaProjectDigestDebugInfo),
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
    /// Which `DigitalSignature*` variant this signature stream corresponds to.
    pub stream_kind: VbaSignatureStreamKind,
    /// Best-effort signer certificate subject (e.g. `CN=...`), if found.
    pub signer_subject: Option<String>,
    /// Raw signature stream bytes.
    pub signature: Vec<u8>,
    /// Verification state (best-effort).
    pub verification: VbaSignatureVerification,
    /// If the signature stream starts with a *length-prefixed* DigSigInfoSerialized-like header
    /// (commonly seen in the wild; distinct from the MS-OSHARED DigSigBlob/offset-based wrapper) and
    /// the header includes a version DWORD, this is the parsed value.
    pub digsig_info_version: Option<u32>,
    /// If the signature stream includes a wrapper that provides deterministic offsets to the
    /// PKCS#7/CMS bytes (e.g. a length-prefixed DigSigInfoSerialized-like header or an MS-OSHARED
    /// `DigSigBlob` offset table), this is the byte offset (from the start of the stream) where the
    /// PKCS#7/CMS `ContentInfo` begins.
    pub pkcs7_offset: Option<usize>,
    /// If the signature stream includes a wrapper that provides deterministic offsets to the
    /// PKCS#7/CMS bytes (e.g. a length-prefixed DigSigInfoSerialized-like header or an MS-OSHARED
    /// `DigSigBlob` offset table), this is the length (in bytes) of the PKCS#7/CMS `ContentInfo` TLV
    /// (supports both strict DER and BER/indefinite-length encodings).
    pub pkcs7_len: Option<usize>,
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

/// Best-effort trust evaluation state for a VBA signature's signing certificate.
///
/// This is intentionally separate from [`VbaSignatureVerification`]:
/// - [`VbaSignatureVerification`] describes whether the PKCS#7/CMS blob is internally valid.
/// - [`VbaCertificateTrust`] describes whether the signing certificate chains to a caller-provided
///   trust anchor set.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VbaCertificateTrust {
    /// Signing certificate chains to a trusted root in the provided store.
    Trusted,
    /// Signing certificate is well-formed but does not chain to a trusted root in the provided
    /// store.
    Untrusted,
    /// Trust was not evaluated (e.g. no trust anchors provided, verification not supported on the
    /// current platform, or the signature wasn't internally valid).
    Unknown,
}

/// Options controlling optional certificate trust evaluation for VBA signatures.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct VbaSignatureTrustOptions {
    /// Trusted root certificates (DER-encoded) to use when evaluating "trusted publisher" policy.
    ///
    /// When empty, certificate trust is not evaluated and the result will be
    /// [`VbaCertificateTrust::Unknown`].
    pub trusted_root_certs_der: Vec<Vec<u8>>,
}

/// A VBA digital signature along with the result of optional publisher trust evaluation.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct VbaDigitalSignatureTrusted {
    pub signature: VbaDigitalSignature,
    pub cert_trust: VbaCertificateTrust,
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
/// Excel stores VBA project signatures in one of the `\u{0005}DigitalSignature*` streams.
/// Some files may contain multiple signature streams (e.g. legacy and newer formats). This API
/// exposes each stream independently.
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

    candidates.sort_by(|a, b| {
        signature_path_rank(a)
            .cmp(&signature_path_rank(b))
            .then(a.cmp(b))
    });

    let mut out = Vec::new();
    for path in candidates {
        let stream_kind =
            signature_path_stream_kind(&path).unwrap_or(VbaSignatureStreamKind::Unknown);
        let signature = ole.read_stream_opt(&path)?.unwrap_or_default();
        let signer_subject = extract_first_certificate_subject(&signature);
        let verification = verify_signature_blob(&signature);
        // Prefer deterministic MS-OSHARED DigSigBlob/WordSigBlob offsets when present.
        let (digsig_info_version, pkcs7_offset, pkcs7_len) =
            if let Some(info) = crate::offcrypto::parse_wordsig_blob(&signature)
                .or_else(|| crate::offcrypto::parse_digsig_blob(&signature))
            {
                (None, Some(info.pkcs7_offset), Some(info.pkcs7_len))
            } else if let Some(info) = crate::offcrypto::parse_digsig_info_serialized(&signature) {
                (info.version, Some(info.pkcs7_offset), Some(info.pkcs7_len))
            } else {
                (None, None, None)
            };

        let (signed_digest_algorithm_oid, signed_digest) =
            match crate::authenticode::extract_vba_signature_signed_digest(&signature) {
                Ok(Some(digest)) => (Some(digest.digest_algorithm_oid), Some(digest.digest)),
                Ok(None) | Err(_) => (None, None),
            };

        out.push(VbaDigitalSignatureStream {
            stream_path: path,
            stream_kind,
            signer_subject,
            signature,
            verification,
            digsig_info_version,
            pkcs7_offset,
            pkcs7_len,
            signed_digest_algorithm_oid,
            signed_digest,
        });
    }

    Ok(out)
}

/// Parsed + verified metadata for a raw VBA signature blob (PKCS#7/CMS).
///
/// Some XLSM producers store the VBA signature payload in a dedicated OOXML part
/// (commonly `xl/vbaProjectSignature.bin`) as raw bytes rather than as an OLE
/// compound file containing `\u{0005}DigitalSignature*` streams.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct VbaSignatureBlobInfo {
    /// Best-effort signer certificate subject (e.g. `CN=...`), if found.
    pub signer_subject: Option<String>,
    /// Raw signature bytes.
    pub signature: Vec<u8>,
    /// Verification state (best-effort).
    pub verification: VbaSignatureVerification,
}

/// Parse and cryptographically verify a raw VBA signature blob (PKCS#7/CMS).
///
/// This is intended for signature payloads that are stored outside the VBA
/// project's OLE container (for example in the OOXML part
/// `xl/vbaProjectSignature.bin`).
pub fn parse_and_verify_vba_signature_blob(signature_blob: &[u8]) -> VbaSignatureBlobInfo {
    let signer_subject = extract_first_certificate_subject(signature_blob);
    let verification = verify_signature_blob(signature_blob);
    VbaSignatureBlobInfo {
        signer_subject,
        signature: signature_blob.to_vec(),
        verification,
    }
}

/// Cryptographically verify a raw VBA signature blob (PKCS#7/CMS).
///
/// Returns `(verification, signer_subject)`.
///
/// On wasm targets `openssl` is unavailable; in that case this always returns
/// [`VbaSignatureVerification::SignedButUnverified`] but still attempts to
/// extract a signer subject via a best-effort X.509 scan.
pub fn verify_vba_signature_blob(
    signature_blob: &[u8],
) -> (VbaSignatureVerification, Option<String>) {
    let signer_subject = extract_first_certificate_subject(signature_blob);
    let verification = verify_signature_blob(signature_blob);
    (verification, signer_subject)
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
    // signature stream exists (see `signature_path_rank` for the Excel-like stream-name ordering).
    candidates.sort_by(|a, b| {
        signature_path_rank(a)
            .cmp(&signature_path_rank(b))
            .then(a.cmp(b))
    });
    let Some(chosen) = candidates.into_iter().next() else {
        return Ok(None);
    };
    let stream_kind =
        signature_path_stream_kind(&chosen).unwrap_or(VbaSignatureStreamKind::Unknown);
    let signature = ole.read_stream_opt(&chosen)?.unwrap_or_default();

    let signer_subject = extract_first_certificate_subject(&signature);

    Ok(Some(VbaDigitalSignature {
        stream_path: chosen,
        stream_kind,
        signer_subject,
        signature,
        verification: VbaSignatureVerification::SignedButUnverified,
        binding: VbaSignatureBinding::Unknown,
    }))
}

/// Parse and verify a VBA digital signature (if present) and return additional Contents Hash
/// binding verification details.
///
/// This is a convenience wrapper around [`verify_vba_digital_signature`] that returns a richer
/// binding status enum with optional debug information (hash algorithm OID/name and the signed vs
/// computed digest bytes).
///
/// Returns `Ok(None)` when the project appears unsigned.
pub fn verify_vba_digital_signature_bound(
    vba_project_bin: &[u8],
) -> Result<Option<VbaDigitalSignatureBound>, SignatureError> {
    let Some(signature) = verify_vba_digital_signature(vba_project_bin)? else {
        return Ok(None);
    };

    // Best-effort debug info for callers.
    let mut debug = VbaProjectDigestDebugInfo::default();

    let signed = match extract_vba_signature_signed_digest(&signature.signature) {
        Ok(Some(v)) => Some(v),
        _ => None,
    };

    if let Some(signed) = signed {
        debug.hash_algorithm_oid = Some(signed.digest_algorithm_oid.clone());
        debug.signed_digest = Some(signed.digest.clone());

        debug.hash_algorithm_name =
            digest_name_from_oid_str(&signed.digest_algorithm_oid).map(str::to_owned);

        match signature.stream_kind {
            VbaSignatureStreamKind::DigitalSignatureExt => {
                // Best-effort: compute the v3 binding digest using the SHA-256 helper and compare it
                // to the signed digest bytes.
                if let Ok(computed) = crate::contents_hash_v3(vba_project_bin) {
                    debug.computed_digest = Some(computed.clone());
                    if signature.verification == VbaSignatureVerification::SignedVerified {
                        let binding = if signed.digest.as_slice() == computed.as_slice() {
                            VbaProjectBindingVerification::BoundVerified(debug)
                        } else {
                            VbaProjectBindingVerification::BoundMismatch(debug)
                        };
                        return Ok(Some(VbaDigitalSignatureBound { signature, binding }));
                    }
                }
            }

            VbaSignatureStreamKind::DigitalSignature => {
                // ContentsHashV1 (legacy): MD5(ContentNormalizedData).
                let Ok(content_normalized) = content_normalized_data(vba_project_bin) else {
                    return Ok(Some(VbaDigitalSignatureBound {
                        signature,
                        binding: VbaProjectBindingVerification::BoundUnknown(debug),
                    }));
                };
                let content_hash: [u8; 16] = Md5::digest(&content_normalized).into();
                debug.computed_digest = Some(content_hash.to_vec());

                if signature.verification == VbaSignatureVerification::SignedVerified {
                    let binding = if signed.digest.as_slice() == content_hash.as_slice() {
                        VbaProjectBindingVerification::BoundVerified(debug)
                    } else {
                        VbaProjectBindingVerification::BoundMismatch(debug)
                    };
                    return Ok(Some(VbaDigitalSignatureBound { signature, binding }));
                }
            }

            VbaSignatureStreamKind::DigitalSignatureEx => {
                // ContentsHashV2 (Agile): MD5(ContentNormalizedData || FormsNormalizedData).
                let Ok(content_normalized) = content_normalized_data(vba_project_bin) else {
                    return Ok(Some(VbaDigitalSignatureBound {
                        signature,
                        binding: VbaProjectBindingVerification::BoundUnknown(debug),
                    }));
                };
                let content_hash: [u8; 16] = Md5::digest(&content_normalized).into();

                let Ok(forms) = forms_normalized_data(vba_project_bin) else {
                    // We can't compute the Agile hash. Still surface the legacy ContentHash as a
                    // best-effort debug value.
                    debug.computed_digest = Some(content_hash.to_vec());
                    return Ok(Some(VbaDigitalSignatureBound {
                        signature,
                        binding: VbaProjectBindingVerification::BoundUnknown(debug),
                    }));
                };

                let mut h = Md5::new();
                h.update(&content_normalized);
                h.update(&forms);
                let agile_hash: [u8; 16] = h.finalize().into();
                debug.computed_digest = Some(agile_hash.to_vec());

                if signature.verification == VbaSignatureVerification::SignedVerified {
                    let binding = if signed.digest.as_slice() == agile_hash.as_slice() {
                        VbaProjectBindingVerification::BoundVerified(debug)
                    } else {
                        VbaProjectBindingVerification::BoundMismatch(debug)
                    };
                    return Ok(Some(VbaDigitalSignatureBound { signature, binding }));
                }
            }

            VbaSignatureStreamKind::Unknown => {
                // Unknown stream variant: best-effort fallback.
                //
                // If we can compute *all* plausible contents-hash candidates derived from the signed
                // digest length:
                // - any match => BoundVerified
                // - no match  => BoundMismatch
                let signed_digest = signed.digest.as_slice();
                let mut match_count = 0usize;
                let mut missing_candidate = false;
                let mut first_computed: Option<Vec<u8>> = None;
                let mut matching_digest: Option<Vec<u8>> = None;

                let want_md5 = signed_digest.len() == 16;
                let mut content_normalized: Option<Vec<u8>> = None;
                if want_md5 {
                    match content_normalized_data(vba_project_bin) {
                        Ok(v) => content_normalized = Some(v),
                        Err(_) => missing_candidate = true,
                    }
                }

                if want_md5 {
                    if let Some(content_normalized) = content_normalized.as_deref() {
                        let content_hash: [u8; 16] = Md5::digest(content_normalized).into();
                        let content_hash_vec = content_hash.to_vec();
                        if first_computed.is_none() {
                            first_computed = Some(content_hash_vec.clone());
                        }
                        if signed_digest == content_hash.as_slice() {
                            match_count += 1;
                            matching_digest = Some(content_hash_vec.clone());
                        }

                        match forms_normalized_data(vba_project_bin) {
                            Ok(forms) => {
                                let mut h = Md5::new();
                                h.update(content_normalized);
                                h.update(&forms);
                                let agile_hash: [u8; 16] = h.finalize().into();
                                let agile_vec = agile_hash.to_vec();
                                if first_computed.is_none() {
                                    first_computed = Some(agile_vec.clone());
                                }
                                if signed_digest == agile_hash.as_slice() {
                                    match_count += 1;
                                    matching_digest = Some(agile_vec);
                                }
                            }
                            Err(_) => missing_candidate = true,
                        }
                    }
                }

                // In this crate, `contents_hash_v3` is a 32-byte SHA-256 digest. If the signed
                // digest is 32 bytes, attempt v3 binding comparison.
                if signed_digest.len() == 32 {
                    match crate::contents_hash_v3(vba_project_bin) {
                        Ok(v3) => {
                            if first_computed.is_none() {
                                first_computed = Some(v3.clone());
                            }
                            if signed_digest == v3.as_slice() {
                                match_count += 1;
                                matching_digest = Some(v3);
                            }
                        }
                        Err(_) => missing_candidate = true,
                    }
                }

                debug.computed_digest = matching_digest.or(first_computed);

                if signature.verification == VbaSignatureVerification::SignedVerified && !missing_candidate {
                    if match_count > 0 {
                        return Ok(Some(VbaDigitalSignatureBound {
                            signature,
                            binding: VbaProjectBindingVerification::BoundVerified(debug),
                        }));
                    }
                    // Only treat this as a definite mismatch when we computed at least one plausible
                    // candidate digest (e.g. MD5/SHA-256-sized digests).
                    if debug.computed_digest.is_some() {
                        return Ok(Some(VbaDigitalSignatureBound {
                            signature,
                            binding: VbaProjectBindingVerification::BoundMismatch(debug),
                        }));
                    }
                }
            }
        }
    }

    Ok(Some(VbaDigitalSignatureBound {
        signature,
        binding: VbaProjectBindingVerification::BoundUnknown(debug),
    }))
}

/// Parse and cryptographically verify a VBA digital signature (if present).
///
/// Returns `Ok(None)` when the project appears unsigned.
///
/// Verification is "internal" CMS/PKCS#7 verification only: we validate that the signature blob
/// is well-formed and that the signature matches the signed attributes / embedded content.
///
/// We also perform a best-effort MS-OVBA binding check by extracting the signed binding digest
/// (Authenticode `SpcIndirectDataContent.messageDigest`) and comparing it against the relevant
/// MS-OVBA Contents Hash transcript (Content Hash / Agile Content Hash, or `ContentsHashV3` for
/// `DigitalSignatureExt`).
///
/// Notes:
/// - For legacy signature streams (`\x05DigitalSignature` / `\x05DigitalSignatureEx`), the embedded
///   digest bytes are always a 16-byte MD5 even when `DigestInfo.digestAlgorithm.algorithm` indicates
///   SHA-256 (MS-OSHARED §4.3).
/// - For v3 (`DigitalSignatureExt`), binding uses the MS-OVBA §2.4.2 v3 content-hash transcript.
///   In the wild, the signed digest bytes are commonly 32-byte SHA-256, but the MS-OVBA v3
///   pseudocode is written in terms of a generic hash function over:
///   `ContentBuffer = V3ContentNormalizedData || ProjectNormalizedData`.
///   `formula-vba` currently verifies v3 binding by comparing the signed digest bytes to
///   [`crate::contents_hash_v3`] (a SHA-256 helper over `project_normalized_data_v3_transcript`; see
///   `docs/vba-digital-signatures.md` for spec vs implementation notes).
///
/// If multiple signature streams are present, we prefer:
/// 1) The first signature stream (by Excel-like stream-name ordering; see `signature_path_rank`)
///    that is both cryptographically verified *and* bound to this project (`binding == Bound`).
/// 2) Otherwise, the first cryptographically verified signature stream (even if not bound).
/// 3) Otherwise, the first signature stream candidate (even if invalid/unparseable).
pub fn verify_vba_digital_signature(
    vba_project_bin: &[u8],
) -> Result<Option<VbaDigitalSignature>, SignatureError> {
    verify_vba_digital_signature_with_project(vba_project_bin, vba_project_bin)
}

/// Verify a VBA digital signature when the signature streams are stored separately from the VBA
/// project streams.
///
/// Some producers (notably XLSM packages) store the `\x05DigitalSignature*` streams in a dedicated
/// OLE part (`xl/vbaProjectSignature.bin`) instead of embedding them inside `xl/vbaProject.bin`.
///
/// In that situation:
/// - the signature stream bytes live in `signature_container_bin`, but
/// - the MS-OVBA Contents Hash binding must be computed over `vba_project_bin`.
///
/// This helper verifies signature streams found in `signature_container_bin` and computes binding
/// against `vba_project_bin`.
pub fn verify_vba_digital_signature_with_project(
    vba_project_bin: &[u8],
    signature_container_bin: &[u8],
) -> Result<Option<VbaDigitalSignature>, SignatureError> {
    let mut ole = OleFile::open(signature_container_bin)?;
    let streams = ole.list_streams()?;

    let mut candidates = streams
        .into_iter()
        .filter(|path| path.split('/').any(is_signature_component))
        .collect::<Vec<_>>();

    if candidates.is_empty() {
        return Ok(None);
    }

    candidates.sort_by(|a, b| {
        signature_path_rank(a)
            .cmp(&signature_path_rank(b))
            .then(a.cmp(b))
    });

    let mut first: Option<VbaDigitalSignature> = None;
    let mut first_verified: Option<VbaDigitalSignature> = None;
    for path in candidates {
        let stream_kind =
            signature_path_stream_kind(&path).unwrap_or(VbaSignatureStreamKind::Unknown);
        let signature = ole.read_stream_opt(&path)?.unwrap_or_default();
        let signer_subject = extract_first_certificate_subject(&signature);
        let verification = verify_signature_blob(&signature);
        let binding = match verification {
            VbaSignatureVerification::SignedVerified => {
                verify_vba_signature_binding_with_stream_path(vba_project_bin, &path, &signature)
            }
            _ => VbaSignatureBinding::Unknown,
        };
        let sig = VbaDigitalSignature {
            stream_path: path,
            stream_kind,
            signer_subject,
            signature,
            verification,
            binding,
        };

        if sig.verification == VbaSignatureVerification::SignedVerified {
            if sig.binding == VbaSignatureBinding::Bound {
                return Ok(Some(sig));
            }
            if first_verified.is_none() {
                first_verified = Some(sig.clone());
            }
        }
        if first.is_none() {
            first = Some(sig);
        }
    }

    Ok(first_verified.or(first))
}

/// Parse and verify a VBA digital signature, optionally evaluating publisher trust.
///
/// This is an opt-in extension of [`verify_vba_digital_signature`]:
/// - First, we perform the same "internal" PKCS#7/CMS verification as
///   [`verify_vba_digital_signature`]. This does *not* validate the certificate chain.
/// - If that internal verification succeeds and the caller provides one or more trusted root
///   certificates in `options`, we then re-run verification with certificate-chain validation
///   enabled.
///
/// Returns `Ok(None)` when the project appears unsigned.
pub fn verify_vba_digital_signature_with_trust(
    vba_project_bin: &[u8],
    options: &VbaSignatureTrustOptions,
) -> Result<Option<VbaDigitalSignatureTrusted>, SignatureError> {
    let mut ole = OleFile::open(vba_project_bin)?;
    let streams = ole.list_streams()?;

    let mut candidates = streams
        .into_iter()
        .filter(|path| path.split('/').any(is_signature_component))
        .collect::<Vec<_>>();

    if candidates.is_empty() {
        return Ok(None);
    }

    candidates.sort_by(|a, b| {
        signature_path_rank(a)
            .cmp(&signature_path_rank(b))
            .then(a.cmp(b))
    });

    let mut first: Option<VbaDigitalSignatureTrusted> = None;
    let mut first_verified: Option<VbaDigitalSignatureTrusted> = None;
    for path in candidates {
        let stream_kind =
            signature_path_stream_kind(&path).unwrap_or(VbaSignatureStreamKind::Unknown);
        let signature_bytes = ole.read_stream_opt(&path)?.unwrap_or_default();
        let signer_subject = extract_first_certificate_subject(&signature_bytes);

        // Internal signature verification (equivalent to `verify_vba_digital_signature`).
        let verification = verify_signature_blob(&signature_bytes);
        let binding = match verification {
            VbaSignatureVerification::SignedVerified => {
                verify_vba_signature_binding_with_stream_path(
                    vba_project_bin,
                    &path,
                    &signature_bytes,
                )
            }
            _ => VbaSignatureBinding::Unknown,
        };

        let cert_trust = if verification == VbaSignatureVerification::SignedVerified
            && !options.trusted_root_certs_der.is_empty()
        {
            verify_pkcs7_trust(&signature_bytes, &options.trusted_root_certs_der)
        } else {
            VbaCertificateTrust::Unknown
        };

        let sig = VbaDigitalSignature {
            stream_path: path,
            stream_kind,
            signer_subject,
            signature: signature_bytes,
            verification,
            binding,
        };
        let sig = VbaDigitalSignatureTrusted {
            signature: sig,
            cert_trust,
        };

        if sig.signature.verification == VbaSignatureVerification::SignedVerified {
            if sig.signature.binding == VbaSignatureBinding::Bound {
                return Ok(Some(sig));
            }
            if first_verified.is_none() {
                first_verified = Some(sig.clone());
            }
        }
        if first.is_none() {
            first = Some(sig);
        }
    }

    Ok(first_verified.or(first))
}

/// Verify whether a VBA signature blob is bound to the given `vbaProject.bin` payload via the
/// MS-OVBA "Contents Hash" mechanism.
///
/// Excel associates each `\x05DigitalSignature*` stream name with a specific contents-hash
/// transcript:
/// - `\x05DigitalSignature`    → legacy Content Hash (MD5 over `ContentNormalizedData`; MS-OSHARED §4.3)
/// - `\x05DigitalSignatureEx`  → legacy Agile Content Hash (MD5 over
///   `ContentNormalizedData || FormsNormalizedData`; MS-OSHARED §4.3)
/// - `\x05DigitalSignatureExt` → MS-OVBA v3 transcript; `formula-vba` currently compares the signed
///   digest bytes to [`crate::contents_hash_v3`] (SHA-256 helper)
///
/// When the stream path is unknown or ambiguous, this uses a best-effort fallback:
/// - Compute every supported contents-hash version that could plausibly match the signed digest
///   length.
/// - If **any** candidate matches and **all** plausible candidates were computed successfully,
///   return [`VbaSignatureBinding::Bound`].
/// - If **no** candidates match and **all** plausible candidates were computed successfully,
///   return [`VbaSignatureBinding::NotBound`].
/// - Otherwise (unsupported digest, missing project data, or missing candidates), return
///   [`VbaSignatureBinding::Unknown`].
///
/// This is a best-effort helper: it returns [`VbaSignatureBinding::Unknown`] when the signature
/// does not contain a supported digest structure, uses an unsupported hash algorithm, or the
/// binding digest cannot be computed.
pub fn verify_vba_signature_binding_with_stream_path(
    vba_project_bin: &[u8],
    signature_stream_path: &str,
    signature: &[u8],
) -> VbaSignatureBinding {
    #[cfg(target_arch = "wasm32")]
    {
        let _ = (vba_project_bin, signature_stream_path, signature);
        VbaSignatureBinding::Unknown
    }

    #[cfg(not(target_arch = "wasm32"))]
    {
        let signed = match extract_vba_signature_signed_digest(signature) {
            Ok(Some(v)) => v,
            _ => return VbaSignatureBinding::Unknown,
        };

        let signed_digest = signed.digest.as_slice();

        let stream_kind = signature_path_known_variant(signature_stream_path);

        match stream_kind {
            Some(VbaSignatureStreamKind::DigitalSignature) => {
                let Ok(content_normalized) = content_normalized_data(vba_project_bin) else {
                    return VbaSignatureBinding::Unknown;
                };
                let content_hash_md5: [u8; 16] = Md5::digest(&content_normalized).into();
                if signed_digest == content_hash_md5.as_slice() {
                    VbaSignatureBinding::Bound
                } else {
                    VbaSignatureBinding::NotBound
                }
            }
            Some(VbaSignatureStreamKind::DigitalSignatureEx) => {
                let Ok(content_normalized) = content_normalized_data(vba_project_bin) else {
                    return VbaSignatureBinding::Unknown;
                };
                let Ok(forms_normalized) = forms_normalized_data(vba_project_bin) else {
                    return VbaSignatureBinding::Unknown;
                };
                let mut h = Md5::new();
                h.update(&content_normalized);
                h.update(&forms_normalized);
                let agile_hash_md5: [u8; 16] = h.finalize().into();
                if signed_digest == agile_hash_md5.as_slice() {
                    VbaSignatureBinding::Bound
                } else {
                    VbaSignatureBinding::NotBound
                }
            }
            Some(VbaSignatureStreamKind::DigitalSignatureExt) => {
                let Ok(computed) = crate::contents_hash_v3(vba_project_bin) else {
                    return VbaSignatureBinding::Unknown;
                };
                if signed_digest == computed.as_slice() {
                    VbaSignatureBinding::Bound
                } else {
                    VbaSignatureBinding::NotBound
                }
            }
            Some(VbaSignatureStreamKind::Unknown) | None => {
                let mut match_count = 0usize;
                let mut missing_candidate = false;

                // ContentsHashV1 / ContentsHashV2 are MD5 (16 bytes).
                if signed_digest.len() == 16 {
                    let content_normalized = match content_normalized_data(vba_project_bin) {
                        Ok(v) => Some(v),
                        Err(_) => {
                            // Both v1 and v2 are plausible candidates for a 16-byte digest.
                            missing_candidate = true;
                            None
                        }
                    };

                    if let Some(content_normalized) = content_normalized.as_deref() {
                        let content_hash_md5: [u8; 16] = Md5::digest(content_normalized).into();
                        if signed_digest == content_hash_md5.as_slice() {
                            match_count += 1;
                        }

                        match forms_normalized_data(vba_project_bin) {
                            Ok(forms_normalized) => {
                                let mut h = Md5::new();
                                h.update(content_normalized);
                                h.update(&forms_normalized);
                                let agile_hash_md5: [u8; 16] = h.finalize().into();
                                if signed_digest == agile_hash_md5.as_slice() {
                                    match_count += 1;
                                }
                            }
                            Err(_) => {
                                missing_candidate = true;
                            }
                        }
                    }
                }

                // In this crate, `contents_hash_v3` is a 32-byte SHA-256 digest. If the signed
                // digest is 32 bytes, attempt v3 binding comparison.
                if signed_digest.len() == 32 {
                    match crate::contents_hash_v3(vba_project_bin) {
                        Ok(computed) => {
                            if signed_digest == computed.as_slice() {
                                match_count += 1;
                            }
                        }
                        Err(_) => missing_candidate = true,
                    }
                }

                if missing_candidate {
                    return VbaSignatureBinding::Unknown;
                }
                if match_count > 0 {
                    return VbaSignatureBinding::Bound;
                }
                if signed_digest.len() == 16 || signed_digest.len() == 32 {
                    return VbaSignatureBinding::NotBound;
                }
                VbaSignatureBinding::Unknown
            }
        }
    }
}

/// Verify whether a VBA signature blob is bound to the given `vbaProject.bin` payload via the
/// MS-OVBA "Contents Hash" mechanism.
///
/// This is a best-effort helper: it returns [`VbaSignatureBinding::Unknown`] when the signature
/// does not contain a supported digest structure, uses an unsupported hash algorithm, or the VBA
/// binding digest cannot be computed.
///
/// Note: signature binding for `\x05DigitalSignatureExt` depends on knowing which signature stream
/// variant is being verified. If you know the stream path, prefer
/// [`verify_vba_signature_binding_with_stream_path`].
pub fn verify_vba_signature_binding(
    vba_project_bin: &[u8],
    signature: &[u8],
) -> VbaSignatureBinding {
    verify_vba_signature_binding_with_stream_path(vba_project_bin, "", signature)
}

/// Evaluate whether the signing certificate embedded in a PKCS#7/CMS VBA signature blob chains to
/// a trusted root.
///
/// Returns [`VbaCertificateTrust::Unknown`] when:
/// - no trusted roots are provided, or
/// - running on `wasm32` targets (OpenSSL is unavailable).
///
/// Note: callers should typically only treat this result as meaningful when the signature's
/// internal verification state is [`VbaSignatureVerification::SignedVerified`].
pub fn verify_vba_signature_certificate_trust(
    signature: &[u8],
    options: &VbaSignatureTrustOptions,
) -> VbaCertificateTrust {
    if options.trusted_root_certs_der.is_empty() {
        return VbaCertificateTrust::Unknown;
    }
    verify_pkcs7_trust(signature, &options.trusted_root_certs_der)
}
fn digest_alg_from_oid_str(oid: &str) -> Option<crate::DigestAlg> {
    // DigestInfo.algorithm values found in Authenticode `SpcIndirectDataContent`.
    //
    // - MD5:     1.2.840.113549.2.5
    // - SHA-1:   1.3.14.3.2.26
    // - SHA-256: 2.16.840.1.101.3.4.2.1
    match oid.trim() {
        // id-md5 (RFC 1321 / PKCS#1)
        "1.2.840.113549.2.5" => Some(crate::DigestAlg::Md5),
        // id-sha1 (RFC 3279)
        "1.3.14.3.2.26" => Some(crate::DigestAlg::Sha1),
        // id-sha256 (NIST)
        "2.16.840.1.101.3.4.2.1" => Some(crate::DigestAlg::Sha256),
        "1.2.840.113549.1.1.4" => Some(crate::DigestAlg::Md5),
        "1.2.840.113549.1.1.5" => Some(crate::DigestAlg::Sha1),
        "1.2.840.113549.1.1.11" => Some(crate::DigestAlg::Sha256),
        _ => None,
    }
}
fn digest_name_from_oid_str(oid: &str) -> Option<&'static str> {
    digest_alg_from_oid_str(oid).map(|alg| match alg {
        crate::DigestAlg::Md5 => "MD5",
        crate::DigestAlg::Sha1 => "SHA-1",
        crate::DigestAlg::Sha256 => "SHA-256",
    })
}
fn signature_kind_rank(kind: VbaSignatureStreamKind) -> u8 {
    match kind {
        VbaSignatureStreamKind::DigitalSignatureExt => 0,
        VbaSignatureStreamKind::DigitalSignatureEx => 1,
        VbaSignatureStreamKind::DigitalSignature => 2,
        VbaSignatureStreamKind::Unknown => 3,
    }
}

fn signature_component_stream_kind(component: &str) -> Option<VbaSignatureStreamKind> {
    // Excel/VBA signature storages/streams are control-character prefixed in OLE; normalize by
    // stripping leading C0 control chars.
    let trimmed = component.trim_start_matches(|c: char| c <= '\u{001F}');
    match trimmed {
        "DigitalSignatureExt" => Some(VbaSignatureStreamKind::DigitalSignatureExt),
        "DigitalSignatureEx" => Some(VbaSignatureStreamKind::DigitalSignatureEx),
        "DigitalSignature" => Some(VbaSignatureStreamKind::DigitalSignature),
        // Some files may use a signature-like storage name that doesn't exactly match one of the
        // known variants. Treat it as a signature candidate, but mark the kind as unknown.
        _ if trimmed.starts_with("DigitalSignature") => Some(VbaSignatureStreamKind::Unknown),
        _ => None,
    }
}

/// Determine the `DigitalSignature*` variant from an OLE stream path for purposes of selecting
/// the MS-OVBA contents-hash version.
///
/// This is intentionally stricter than [`signature_path_stream_kind`]:
/// - Only exact, known variant names map to a contents-hash version.
/// - If the path contains multiple distinct known variants, we treat it as ambiguous and return
///   `None` (callers should fall back to a conservative "try all versions" strategy).
#[cfg(not(target_arch = "wasm32"))]
fn signature_path_known_variant(path: &str) -> Option<VbaSignatureStreamKind> {
    let mut found: Option<VbaSignatureStreamKind> = None;
    for component in path.split(['/', ':']) {
        let trimmed = component.trim_start_matches(|c: char| c <= '\u{001F}');
        let kind = match trimmed {
            "DigitalSignatureExt" => VbaSignatureStreamKind::DigitalSignatureExt,
            "DigitalSignatureEx" => VbaSignatureStreamKind::DigitalSignatureEx,
            "DigitalSignature" => VbaSignatureStreamKind::DigitalSignature,
            _ => continue,
        };
        if let Some(prev) = found {
            if prev != kind {
                return None;
            }
        }
        found = Some(kind);
    }
    found
}

fn signature_path_stream_kind(path: &str) -> Option<VbaSignatureStreamKind> {
    let mut best: Option<(u8, VbaSignatureStreamKind)> = None;
    // Stream paths are normally OLE paths (components separated by `/`), but higher-level wrappers
    // (e.g. `formula-xlsx`) may prefix the OLE path with an OPC part name using `:`, like:
    // `xl/vbaProjectSignature.bin:\x05DigitalSignatureExt`.
    //
    // Be permissive and scan both `/` and `:`-delimited components for `DigitalSignature*`.
    for component in path.split(['/', ':']) {
        let Some(kind) = signature_component_stream_kind(component) else {
            continue;
        };
        let rank = signature_kind_rank(kind);
        match best {
            None => best = Some((rank, kind)),
            Some((best_rank, _)) if rank < best_rank => best = Some((rank, kind)),
            _ => {}
        }
    }
    best.map(|(_, kind)| kind)
}

fn is_signature_component(component: &str) -> bool {
    signature_component_stream_kind(component).is_some()
}

fn signature_path_rank(path: &str) -> u8 {
    // Lower = higher priority.
    //
    // Excel stores the VBA project signature in one of three `\x05DigitalSignature*` streams:
    // - "\x05DigitalSignature"     (legacy; older Office versions)
    // - "\x05DigitalSignatureEx"   (extended; newer Office versions, used for SHA-2-era signatures)
    // - "\x05DigitalSignatureExt"  (extension; newest variant)
    //
    // When multiple signature streams are present, Excel prefers the newest stream name.
    //
    // Note: MS-OVBA specifies how a VBA project's digest is computed (see "Contents Hash"
    // §2.4.2 and "Project Integrity Verification" §4.1), but it does not normatively specify the
    // on-disk stream name precedence between the `DigitalSignature*` streams. This ordering is
    // based on observed Office/Excel behavior and existing files produced by different Office
    // versions.
    // https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/
    signature_path_stream_kind(path)
        .map(signature_kind_rank)
        .unwrap_or(3)
}

fn bytes_to_lower_hex(bytes: &[u8]) -> String {
    const HEX: &[u8; 16] = b"0123456789abcdef";
    let mut out = String::new();
    let _ = out.try_reserve(bytes.len().saturating_mul(2));
    for &b in bytes {
        out.push(HEX[(b >> 4) as usize] as char);
        out.push(HEX[(b & 0x0F) as usize] as char);
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
    bytes.get(..consumed_len)
}

fn scan_for_embedded_der_certificates(bytes: &[u8]) -> impl Iterator<Item = &[u8]> {
    use x509_parser::prelude::parse_x509_certificate;

    // Heuristic: certificates are DER-encoded and begin with a SEQUENCE (0x30) tag.
    // Yield each candidate DER slice that parses as a certificate.
    (0..bytes.len()).filter_map(move |start| {
        if bytes.get(start).copied()? != 0x30 {
            return None;
        }
        let slice = bytes.get(start..)?;
        let (rem, _) = parse_x509_certificate(slice).ok()?;
        let consumed_len = slice.len().saturating_sub(rem.len());
        if consumed_len == 0 {
            return None;
        }
        slice.get(..consumed_len)
    })
}

fn verify_pkcs7_trust(
    _signature: &[u8],
    _trusted_root_certs_der: &[Vec<u8>],
) -> VbaCertificateTrust {
    #[cfg(target_arch = "wasm32")]
    {
        let _ = (_signature, _trusted_root_certs_der);
        // `openssl` isn't available on wasm targets.
        VbaCertificateTrust::Unknown
    }

    #[cfg(not(target_arch = "wasm32"))]
    {
        verify_pkcs7_trust_openssl(_signature, _trusted_root_certs_der)
    }
}

#[cfg(not(target_arch = "wasm32"))]
fn verify_pkcs7_trust_openssl(
    signature: &[u8],
    trusted_root_certs_der: &[Vec<u8>],
) -> VbaCertificateTrust {
    use openssl::pkcs7::Pkcs7Flags;
    use openssl::stack::Stack;
    use openssl::x509::store::X509StoreBuilder;
    use openssl::x509::X509;

    let Some((pkcs7, pkcs7_offset)) = parse_pkcs7_with_offset(signature) else {
        // Internal verification succeeded, but we can't re-parse for trust evaluation.
        return VbaCertificateTrust::Untrusted;
    };

    let mut builder = match X509StoreBuilder::new() {
        Ok(builder) => builder,
        Err(_) => return VbaCertificateTrust::Untrusted,
    };

    for root_der in trusted_root_certs_der {
        let cert = match X509::from_der(root_der) {
            Ok(cert) => cert,
            Err(_) => return VbaCertificateTrust::Untrusted,
        };
        if builder.add_cert(cert).is_err() {
            return VbaCertificateTrust::Untrusted;
        }
    }

    let store = builder.build();

    let empty_certs = match Stack::<X509>::new() {
        Ok(stack) => stack,
        Err(_) => return VbaCertificateTrust::Untrusted,
    };
    let certs = pkcs7
        .signed()
        .and_then(|s| s.certificates())
        .unwrap_or(&empty_certs);

    let flags = Pkcs7Flags::BINARY;

    // Verify with embedded content first.
    if pkcs7.verify(certs, &store, None, None, flags).is_ok() {
        return VbaCertificateTrust::Trusted;
    }

    // If the PKCS#7 blob was found after a prefix/header, try treating the prefix as detached
    // content.
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
            || pkcs7
                .verify(certs, &store, Some(prefix), None, flags)
                .is_ok()
        {
            return VbaCertificateTrust::Trusted;
        }
    }

    VbaCertificateTrust::Untrusted
}

#[cfg(not(target_arch = "wasm32"))]
fn parse_pkcs7_with_offset(signature: &[u8]) -> Option<(openssl::pkcs7::Pkcs7, usize)> {
    use openssl::pkcs7::Pkcs7;

    fn parse_pkcs7_exact(bytes: &[u8]) -> Option<Pkcs7> {
        let pkcs7 = Pkcs7::from_der(bytes).ok()?;
        // OpenSSL's decoder may accept BER-ish inputs; re-encode and ensure the PKCS#7 we parsed
        // corresponds to the full byte slice to avoid false positives when scanning.
        let der = pkcs7.to_der().ok()?;
        if der == bytes {
            Some(pkcs7)
        } else {
            None
        }
    }

    // Some producers include a small header before the PKCS#7 payload. Try parsing from the start
    // first, then scan for an embedded SEQUENCE that parses as PKCS#7.
    if let Some(pkcs7) = parse_pkcs7_exact(signature) {
        return Some((pkcs7, 0));
    }

    // Keep the last plausible candidate found during scanning.
    //
    // Real-world signature streams can contain multiple PKCS#7 SignedData blobs (e.g. a PKCS#7
    // certificate store followed by the actual signature). The signature payload usually appears
    // last, so we prefer the last candidate in the stream.
    let mut best: Option<(Pkcs7, usize)> = None;

    // Office apps sometimes wrap the PKCS#7 blob in an [MS-OSHARED] DigSigBlob/WordSigBlob that
    // points at the actual signature buffer via offsets. Parsing it avoids false positives when the
    // stream contains multiple SignedData blobs (e.g. cert stores or other metadata).
    if let Some(info) = crate::offcrypto::parse_wordsig_blob(signature) {
        let end = info.pkcs7_offset.saturating_add(info.pkcs7_len);
        if end <= signature.len() {
            let slice = &signature[info.pkcs7_offset..end];
            if let Ok(pkcs7) = Pkcs7::from_der(slice) {
                return Some((pkcs7, info.pkcs7_offset));
            }
        }
    }
    if let Some(info) = crate::offcrypto::parse_digsig_blob(signature) {
        let end = info.pkcs7_offset.saturating_add(info.pkcs7_len);
        if end <= signature.len() {
            let slice = &signature[info.pkcs7_offset..end];
            if let Ok(pkcs7) = Pkcs7::from_der(slice) {
                return Some((pkcs7, info.pkcs7_offset));
            }
        }
    }

    // Many `\x05DigitalSignature*` streams start with a length-prefixed DigSigInfoSerialized-like
    // header. Parsing it is deterministic and avoids the worst-case behavior of scanning/parsing
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
        let slice = &signature[start..];

        // Avoid repeatedly attempting full OpenSSL parses at every 0x30 byte offset. Instead, do
        // a lightweight BER scan for a plausible PKCS#7 SignedData ContentInfo and only then hand
        // the exact slice to OpenSSL.
        let Some(len) = crate::offcrypto::pkcs7_signed_data_len(slice) else {
            continue;
        };
        let candidate = &slice[..len];

        // Prefer the last candidate in the stream.
        if let Some(pkcs7) = parse_pkcs7_exact(candidate) {
            best = Some((pkcs7, start));
            continue;
        }
        if let Ok(pkcs7) = Pkcs7::from_der(candidate) {
            best = Some((pkcs7, start));
        }
    }

    if let Some(best) = best {
        return Some(best);
    }

    // Fallback: use our BER/DER tolerant SignedData locator (handles indefinite-length encodings).
    // We only do this if the definite-length scanner didn't find any candidates.
    if let Some((offset, len)) = crate::authenticode::locate_pkcs7_signed_data_bounds(signature) {
        let end = offset.saturating_add(len);
        if end <= signature.len() {
            let slice = &signature[offset..end];
            if let Ok(pkcs7) = Pkcs7::from_der(slice) {
                return Some((pkcs7, offset));
            }
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
        VbaSignatureVerification::SignedButUnverified
    }

    #[cfg(not(target_arch = "wasm32"))]
    use openssl::pkcs7::Pkcs7Flags;
    #[cfg(not(target_arch = "wasm32"))]
    use openssl::stack::Stack;
    #[cfg(not(target_arch = "wasm32"))]
    use openssl::x509::store::X509StoreBuilder;
    #[cfg(not(target_arch = "wasm32"))]
    use openssl::x509::X509;

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
                || pkcs7
                    .verify(certs, &store, Some(prefix), None, flags)
                    .is_ok()
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
        if bytes.get(start).copied() != Some(0x30) {
            continue;
        }
        let Some(slice) = bytes.get(start..) else {
            continue;
        };
        if let Ok((_, cert)) = parse_x509_certificate(slice) {
            return Some(cert.subject().to_string());
        }
    }

    None
}

/// Verify MS-OVBA "Contents Hash" signature binding when the signature payload is provided
/// separately (e.g. `xl/vbaProjectSignature.bin` in OOXML / XLSM).
///
/// `signature_bytes` can be:
/// - An OLE container holding `\u{0005}DigitalSignature*` streams (like `vbaProjectSignature.bin`)
/// - A raw `\u{0005}DigitalSignature*` stream payload
/// - A raw PKCS#7/CMS blob (DER), optionally wrapped in an Office DigSig structure
///
/// Returns [`VbaProjectBindingVerification::BoundUnknown`] when no signed binding digest could be
/// extracted (or when we can't compare it to the project bytes).
pub fn verify_vba_project_signature_binding(
    project_ole: &[u8],
    signature_bytes: &[u8],
) -> Result<VbaProjectBindingVerification, SignatureError> {
    let payloads = signature_payload_candidates(signature_bytes);

    let mut any_signed_digest = None::<VbaProjectDigestDebugInfo>;
    // First comparison attempt (may be ambiguous/unknown).
    let mut first_any_comparison = None::<VbaProjectDigestDebugInfo>;
    // First *definite* mismatch (we could compute the relevant digest candidate(s) and none
    // matched).
    let mut first_definite_mismatch = None::<VbaProjectDigestDebugInfo>;

    // Lazily computed MS-OVBA v1/v2 digests for the project bytes.
    //
    // Outer Option = attempted; inner Option = computed successfully.
    let mut content_hash_md5: Option<Option<[u8; 16]>> = None;
    let mut content_normalized: Option<Vec<u8>> = None;
    let mut agile_hash_md5: Option<Option<[u8; 16]>> = None;

    // Lazily computed v3 binding digest (`contents_hash_v3`, currently a 32-byte SHA-256).
    // Outer Option = attempted; inner Option = computed successfully.
    let mut contents_hash_v3: Option<Option<Vec<u8>>> = None;

    for payload in payloads {
        let signed = match extract_vba_signature_signed_digest(&payload.bytes) {
            Ok(Some(signed)) => signed,
            Ok(None) | Err(_) => continue,
        };

        let mut debug = VbaProjectDigestDebugInfo {
            hash_algorithm_oid: Some(signed.digest_algorithm_oid.clone()),
            hash_algorithm_name: digest_name_from_oid_str(&signed.digest_algorithm_oid)
                .map(str::to_owned),
            signed_digest: Some(signed.digest.clone()),
            ..Default::default()
        };

        if any_signed_digest.is_none() {
            any_signed_digest = Some(debug.clone());
        }

        let signed_digest = signed.digest.as_slice();

        // Helper to record debug info for later if we don't find a bound signature.
        let mut record_any = |debug: &VbaProjectDigestDebugInfo| {
            if first_any_comparison.is_none() {
                first_any_comparison = Some(debug.clone());
            }
        };
        let mut record_definite_mismatch = |debug: &VbaProjectDigestDebugInfo| {
            record_any(debug);
            if first_definite_mismatch.is_none() {
                first_definite_mismatch = Some(debug.clone());
            }
        };

        match payload.stream_kind {
            Some(VbaSignatureStreamKind::DigitalSignatureExt) => {
                if contents_hash_v3.is_none() {
                    contents_hash_v3 = Some(crate::contents_hash_v3(project_ole).ok());
                }
                let Some(computed) = contents_hash_v3.as_ref().and_then(|v| v.as_ref()) else {
                    record_any(&debug);
                    continue;
                };

                debug.computed_digest = Some(computed.clone());

                if signed_digest == computed.as_slice() {
                    return Ok(VbaProjectBindingVerification::BoundVerified(debug));
                }
                record_definite_mismatch(&debug);
            }

            Some(VbaSignatureStreamKind::DigitalSignature) => {
                // ContentsHashV1: MD5(ContentNormalizedData).
                if content_hash_md5.is_none() {
                    match content_normalized_data(project_ole) {
                        Ok(v) => {
                            let hash: [u8; 16] = Md5::digest(&v).into();
                            content_normalized = Some(v);
                            content_hash_md5 = Some(Some(hash));
                        }
                        Err(_) => {
                            content_hash_md5 = Some(None);
                        }
                    }
                }

                let Some(Some(content_hash_md5)) = content_hash_md5 else {
                    record_any(&debug);
                    continue;
                };

                debug.computed_digest = Some(content_hash_md5.to_vec());

                if signed_digest == content_hash_md5.as_slice() {
                    return Ok(VbaProjectBindingVerification::BoundVerified(debug));
                }
                record_definite_mismatch(&debug);
            }

            Some(VbaSignatureStreamKind::DigitalSignatureEx) => {
                // ContentsHashV2 (Agile): MD5(ContentNormalizedData || FormsNormalizedData).
                if content_hash_md5.is_none() {
                    match content_normalized_data(project_ole) {
                        Ok(v) => {
                            let hash: [u8; 16] = Md5::digest(&v).into();
                            content_normalized = Some(v);
                            content_hash_md5 = Some(Some(hash));
                        }
                        Err(_) => {
                            content_hash_md5 = Some(None);
                        }
                    }
                }

                // Surface ContentHashV1 as a best-effort debug prefix even though we need the Agile
                // hash for actual comparison.
                if let Some(Some(content_hash_md5)) = content_hash_md5 {
                    debug.computed_digest = Some(content_hash_md5.to_vec());
                }

                if agile_hash_md5.is_none() {
                    if let Some(content_normalized) = content_normalized.as_deref() {
                        agile_hash_md5 = Some(forms_normalized_data(project_ole).ok().map(|forms| {
                            let mut h = Md5::new();
                            h.update(content_normalized);
                            h.update(&forms);
                            h.finalize().into()
                        }));
                    } else {
                        agile_hash_md5 = Some(None);
                    }
                }

                match agile_hash_md5 {
                    Some(Some(agile)) => {
                        debug.computed_digest = Some(agile.to_vec());
                        if signed_digest == agile.as_slice() {
                            return Ok(VbaProjectBindingVerification::BoundVerified(debug));
                        }
                        record_definite_mismatch(&debug);
                    }
                    _ => {
                        // Can't compute FormsNormalizedData; treat as unknown.
                        record_any(&debug);
                    }
                }
            }

            Some(VbaSignatureStreamKind::Unknown) | None => {
                // Unknown stream kind: conservative fallback.
                //
                // Unknown stream kind: best-effort comparison against every plausible contents-hash
                // candidate derived from the digest length.
                //
                // If we can compute *all* plausible candidates:
                // - any match => BoundVerified (the signature is bound, even if we can't disambiguate
                //   between V1/V2 when the digests are identical),
                // - no match  => BoundMismatch (definite "not bound" under any supported scheme).
                let mut match_count = 0usize;
                let mut missing_candidate = false;
                let mut first_computed: Option<Vec<u8>> = None;
                let mut matching_digest: Option<Vec<u8>> = None;

                if signed_digest.len() == 16 {
                    if content_hash_md5.is_none() {
                        match content_normalized_data(project_ole) {
                            Ok(v) => {
                                let hash: [u8; 16] = Md5::digest(&v).into();
                                content_normalized = Some(v);
                                content_hash_md5 = Some(Some(hash));
                            }
                            Err(_) => {
                                content_hash_md5 = Some(None);
                            }
                        }
                    }

                    match content_hash_md5 {
                        Some(Some(content)) => {
                            let v = content.to_vec();
                            if first_computed.is_none() {
                                first_computed = Some(v.clone());
                            }
                            if signed_digest == content.as_slice() {
                                match_count += 1;
                                matching_digest = Some(v);
                            }
                        }
                        _ => missing_candidate = true,
                    }

                    if agile_hash_md5.is_none() {
                        if let Some(content_normalized) = content_normalized.as_deref() {
                            agile_hash_md5 = Some(forms_normalized_data(project_ole).ok().map(|forms| {
                                let mut h = Md5::new();
                                h.update(content_normalized);
                                h.update(&forms);
                                h.finalize().into()
                            }));
                        } else {
                            agile_hash_md5 = Some(None);
                        }
                    }

                    match agile_hash_md5 {
                        Some(Some(agile)) => {
                            let v = agile.to_vec();
                            if first_computed.is_none() {
                                first_computed = Some(v.clone());
                            }
                            if signed_digest == agile.as_slice() {
                                match_count += 1;
                                matching_digest = Some(v);
                            }
                        }
                        _ => missing_candidate = true,
                    }
                }

                if signed_digest.len() == 32 {
                    if contents_hash_v3.is_none() {
                        contents_hash_v3 = Some(crate::contents_hash_v3(project_ole).ok());
                    }

                    match contents_hash_v3.as_ref().and_then(|v| v.as_ref()) {
                        Some(v3) => {
                            if first_computed.is_none() {
                                first_computed = Some(v3.clone());
                            }
                            if signed_digest == v3.as_slice() {
                                match_count += 1;
                                matching_digest = Some(v3.clone());
                            }
                        }
                        None => missing_candidate = true,
                    }
                }

                debug.computed_digest = matching_digest.or(first_computed);

                if !missing_candidate {
                    if match_count > 0 {
                        return Ok(VbaProjectBindingVerification::BoundVerified(debug));
                    }
                    // If we computed every plausible candidate and none matched, treat this as a
                    // definite mismatch (the signature is not bound under any supported scheme).
                    if match_count == 0 && debug.computed_digest.is_some() {
                        record_definite_mismatch(&debug);
                        continue;
                    }
                }

                record_any(&debug);
            }
        }
    }

    if let Some(debug) = first_definite_mismatch {
        return Ok(VbaProjectBindingVerification::BoundMismatch(debug));
    }
    if let Some(debug) = first_any_comparison {
        return Ok(VbaProjectBindingVerification::BoundUnknown(debug));
    }

    Ok(VbaProjectBindingVerification::BoundUnknown(
        any_signed_digest.unwrap_or_default(),
    ))
}



#[derive(Debug, Clone)]
struct SignaturePayloadCandidate {
    stream_kind: Option<VbaSignatureStreamKind>,
    bytes: Vec<u8>,
}

fn signature_payload_candidates(signature_bytes: &[u8]) -> Vec<SignaturePayloadCandidate> {
    // 1) If it looks like an OLE container, extract `\x05DigitalSignature*` streams.
    if let Ok(mut ole) = OleFile::open(signature_bytes) {
        if let Ok(streams) = ole.list_streams() {
            let mut candidates = streams
                .into_iter()
                .filter(|path| path.split('/').any(is_signature_component))
                .collect::<Vec<_>>();
            candidates.sort_by(|a, b| {
                signature_path_rank(a)
                    .cmp(&signature_path_rank(b))
                    .then(a.cmp(b))
            });

            let mut out = Vec::new();
            for path in candidates {
                if let Ok(Some(bytes)) = ole.read_stream_opt(&path) {
                    out.push(SignaturePayloadCandidate {
                        stream_kind: signature_path_stream_kind(&path),
                        bytes,
                    });
                }
            }
            if !out.is_empty() {
                return out;
            }
        }
    }

    // 2) Otherwise treat the whole buffer as a signature blob/stream payload.
    vec![SignaturePayloadCandidate {
        stream_kind: None,
        bytes: signature_bytes.to_vec(),
    }]
}

#[cfg(test)]
mod tests {
    use crate::DigestAlg;

    use super::{digest_alg_from_oid_str, digest_name_from_oid_str};

    #[test]
    fn digest_alg_from_oid_str_maps_known_digest_oids() {
        assert_eq!(
            digest_alg_from_oid_str("1.2.840.113549.2.5"),
            Some(DigestAlg::Md5)
        );
        assert_eq!(
            digest_alg_from_oid_str("1.3.14.3.2.26"),
            Some(DigestAlg::Sha1)
        );
        assert_eq!(
            digest_alg_from_oid_str("2.16.840.1.101.3.4.2.1"),
            Some(DigestAlg::Sha256)
        );

        // Be permissive about surrounding whitespace.
        assert_eq!(
            digest_alg_from_oid_str("  1.2.840.113549.2.5  "),
            Some(DigestAlg::Md5)
        );
    }

    #[test]
    fn digest_alg_from_oid_str_maps_common_signature_oids() {
        // Some signatures incorrectly use signature algorithm OIDs where a digest algorithm OID
        // would normally appear; we accept a small set of these in best-effort mode.
        assert_eq!(
            digest_alg_from_oid_str("1.2.840.113549.1.1.4"),
            Some(DigestAlg::Md5)
        ); // md5WithRSAEncryption
        assert_eq!(
            digest_alg_from_oid_str("1.2.840.113549.1.1.5"),
            Some(DigestAlg::Sha1)
        ); // sha1WithRSAEncryption
        assert_eq!(
            digest_alg_from_oid_str("1.2.840.113549.1.1.11"),
            Some(DigestAlg::Sha256)
        ); // sha256WithRSAEncryption
    }

    #[test]
    fn digest_name_from_oid_str_is_exhaustive_for_supported_algs() {
        assert_eq!(digest_name_from_oid_str("1.2.840.113549.2.5"), Some("MD5"));
        assert_eq!(digest_name_from_oid_str("1.3.14.3.2.26"), Some("SHA-1"));
        assert_eq!(
            digest_name_from_oid_str("2.16.840.1.101.3.4.2.1"),
            Some("SHA-256")
        );

        assert_eq!(digest_name_from_oid_str("0.0"), None);
    }
}
