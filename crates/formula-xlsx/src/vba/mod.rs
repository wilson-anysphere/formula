//! Integration glue between `formula-xlsx` and `formula-vba`.

use std::collections::BTreeMap;

use crate::{XlsxDocument, XlsxPackage};

pub use formula_vba::{
    SignatureError, VbaDigitalSignature, VbaDigitalSignatureBound, VbaProjectBindingVerification,
    VbaProjectDigestDebugInfo, VbaSignatureVerification, VBAModule, VBAProject, VBAReference,
    VbaCertificateTrust, VbaDigitalSignatureTrusted, VbaSignatureStreamKind, VbaSignatureTrustOptions,
};

const VBA_PROJECT_BIN: &str = "xl/vbaProject.bin";
const VBA_PROJECT_BIN_RELS: &str = "xl/_rels/vbaProject.bin.rels";
const VBA_PROJECT_SIGNATURE_BIN: &str = "xl/vbaProjectSignature.bin";
const VBA_PROJECT_SIGNATURE_REL_TYPE: &str =
    "http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature";

impl XlsxPackage {
    /// Parse and return a structured VBA project model (for UI display).
    pub fn vba_project(&self) -> Result<Option<VBAProject>, formula_vba::ParseError> {
        let Some(bin) = get_part(self.parts_map(), VBA_PROJECT_BIN) else {
            return Ok(None);
        };
        Ok(Some(VBAProject::parse(bin)?))
    }

    /// Inspect the workbook's VBA project for a digital signature stream.
    ///
    /// Signed XLSM files may store signature streams in a separate OPC part
    /// (`xl/vbaProjectSignature.bin`, or another part referenced from
    /// `xl/_rels/vbaProject.bin.rels`). We prefer that part when present, but fall back to
    /// inspecting `xl/vbaProject.bin`.
    pub fn parse_vba_digital_signature(&self) -> Result<Option<VbaDigitalSignature>, SignatureError> {
        parse_vba_digital_signature_from_parts(self.parts_map())
    }

    /// Inspect and (best-effort) cryptographically verify the VBA project digital signature.
    ///
    /// This mirrors [`formula_vba::verify_vba_digital_signature`], but prefers the dedicated
    /// signature part when present.
    ///
    /// Some producers store `xl/vbaProjectSignature.bin` as raw PKCS#7/CMS bytes (not an OLE
    /// compound file). In that case we fall back to verifying it as a raw signature blob.
    pub fn verify_vba_digital_signature(&self) -> Result<Option<VbaDigitalSignature>, SignatureError> {
        verify_vba_digital_signature_from_parts(self.parts_map())
    }

    /// Inspect and (best-effort) cryptographically verify the VBA project digital signature,
    /// optionally evaluating publisher trust.
    ///
    /// This mirrors [`formula_vba::verify_vba_digital_signature_with_trust`], but prefers the
    /// dedicated signature part when present (`xl/vbaProjectSignature.bin` or the part referenced
    /// from `xl/_rels/vbaProject.bin.rels`).
    ///
    /// Returns `Ok(None)` when `xl/vbaProject.bin` is absent.
    pub fn verify_vba_digital_signature_with_trust(
        &self,
        options: &formula_vba::VbaSignatureTrustOptions,
    ) -> Result<Option<formula_vba::VbaDigitalSignatureTrusted>, formula_vba::SignatureError> {
        verify_vba_digital_signature_with_trust_from_parts(self.parts_map(), options)
    }

    /// Verify MS-OVBA "Contents Hash" signature binding for an embedded VBA project.
    ///
    /// Returns `Ok(None)` when there is no `xl/vbaProject.bin`.
    pub fn vba_project_signature_binding(
        &self,
    ) -> Result<Option<formula_vba::VbaProjectBindingVerification>, SignatureError> {
        let parts = self.parts_map();
        let Some(project_ole) = get_part(parts, VBA_PROJECT_BIN) else {
            return Ok(None);
        };

        let signature_bytes = resolve_vba_signature_part_name(parts)
            .and_then(|name| get_part(parts, &name))
            .unwrap_or(project_ole);

        let binding = formula_vba::verify_vba_project_signature_binding(project_ole, signature_bytes)?;
        Ok(Some(upgrade_vba_project_signature_binding(project_ole, binding)))
    }
}

impl XlsxDocument {
    /// Parse and return a structured VBA project model (for UI display).
    pub fn vba_project(&self) -> Result<Option<VBAProject>, formula_vba::ParseError> {
        let Some(bin) = get_part(self.parts(), VBA_PROJECT_BIN) else {
            return Ok(None);
        };
        Ok(Some(VBAProject::parse(bin)?))
    }

    /// Inspect the workbook's VBA project for a digital signature stream.
    pub fn parse_vba_digital_signature(&self) -> Result<Option<VbaDigitalSignature>, SignatureError> {
        parse_vba_digital_signature_from_parts(self.parts())
    }

    /// Inspect and (best-effort) cryptographically verify the VBA project digital signature.
    pub fn verify_vba_digital_signature(&self) -> Result<Option<VbaDigitalSignature>, SignatureError> {
        verify_vba_digital_signature_from_parts(self.parts())
    }

    /// Inspect and (best-effort) cryptographically verify the VBA project digital signature,
    /// optionally evaluating publisher trust.
    ///
    /// This mirrors [`formula_vba::verify_vba_digital_signature_with_trust`], but prefers the
    /// dedicated signature part when present (`xl/vbaProjectSignature.bin` or the part referenced
    /// from `xl/_rels/vbaProject.bin.rels`).
    ///
    /// Returns `Ok(None)` when `xl/vbaProject.bin` is absent.
    pub fn verify_vba_digital_signature_with_trust(
        &self,
        options: &formula_vba::VbaSignatureTrustOptions,
    ) -> Result<Option<formula_vba::VbaDigitalSignatureTrusted>, formula_vba::SignatureError> {
        verify_vba_digital_signature_with_trust_from_parts(self.parts(), options)
    }

    /// Verify MS-OVBA "Contents Hash" signature binding for an embedded VBA project.
    ///
    /// Returns `Ok(None)` when there is no `xl/vbaProject.bin`.
    pub fn vba_project_signature_binding(
        &self,
    ) -> Result<Option<formula_vba::VbaProjectBindingVerification>, SignatureError> {
        let parts = self.parts();
        let Some(project_ole) = get_part(parts, VBA_PROJECT_BIN) else {
            return Ok(None);
        };

        let signature_bytes = resolve_vba_signature_part_name(parts)
            .and_then(|name| get_part(parts, &name))
            .unwrap_or(project_ole);

        let binding = formula_vba::verify_vba_project_signature_binding(project_ole, signature_bytes)?;
        Ok(Some(upgrade_vba_project_signature_binding(project_ole, binding)))
    }
}

fn parse_vba_digital_signature_from_parts(
    parts: &BTreeMap<String, Vec<u8>>,
) -> Result<Option<VbaDigitalSignature>, SignatureError> {
    if let Some(signature_part_name) = resolve_vba_signature_part_name(parts) {
        if let Some(bytes) = get_part(parts, &signature_part_name) {
            match formula_vba::parse_vba_digital_signature(bytes) {
                Ok(Some(mut sig)) => {
                    sig.stream_path = format!("{signature_part_name}:{}", sig.stream_path);
                    return Ok(Some(sig));
                }
                Ok(None) => {}
                Err(_) => {
                    // Non-OLE signature part: treat it as a raw signature blob for parsing.
                    //
                    // We can't reliably distinguish "valid raw signature" from "garbage" without
                    // additional parsing/verification. As a best-effort heuristic, only treat it as a
                    // signature when we can find an embedded signer certificate.
                    let signer_subject =
                        formula_vba::extract_signer_certificate_info(bytes).map(|info| info.subject);
                    if signer_subject.is_some() {
                        return Ok(Some(VbaDigitalSignature {
                            stream_path: signature_part_name,
                            stream_kind: formula_vba::VbaSignatureStreamKind::Unknown,
                            signer_subject,
                            signature: bytes.to_vec(),
                            verification: VbaSignatureVerification::SignedButUnverified,
                            binding: formula_vba::VbaSignatureBinding::Unknown,
                        }));
                    }
                }
            }
        }
    }

    let Some(vba_bin) = get_part(parts, VBA_PROJECT_BIN) else {
        return Ok(None);
    };
    formula_vba::parse_vba_digital_signature(vba_bin)
}

fn verify_vba_digital_signature_from_parts(
    parts: &BTreeMap<String, Vec<u8>>,
) -> Result<Option<VbaDigitalSignature>, SignatureError> {
    let signature_part_name = resolve_vba_signature_part_name(parts);
    let mut signature_part_result: Option<VbaDigitalSignature> = None;
    let vba_project_bin = get_part(parts, VBA_PROJECT_BIN);

    if let Some(signature_part_name) = signature_part_name {
        if let Some(bytes) = get_part(parts, &signature_part_name) {
            // Attempt to treat the part as an OLE/CFB container first (the most common format).
            let verified = match vba_project_bin {
                Some(vba_project_bin) => {
                    formula_vba::verify_vba_digital_signature_with_project(vba_project_bin, bytes)
                }
                None => formula_vba::verify_vba_digital_signature(bytes),
            };

            match verified {
                Ok(Some(mut sig)) => {
                    // Preserve which part we read the signature from to avoid ambiguity.
                    sig.stream_path = format!("{signature_part_name}:{}", sig.stream_path);
                    // If we couldn't locate the corresponding `vbaProject.bin`, we can't evaluate the
                    // MS-OVBA "Contents Hash" binding and should not report a mismatch.
                    if vba_project_bin.is_none() {
                        sig.binding = formula_vba::VbaSignatureBinding::Unknown;
                    }
                    signature_part_result = Some(sig);
                }
                Ok(None) => {}
                Err(_) => {
                    // Not an OLE container: fall back to verifying the part bytes as a raw PKCS#7/CMS
                    // signature blob.
                    let (verification, signer_subject) = formula_vba::verify_vba_signature_blob(bytes);
                    signature_part_result = Some(VbaDigitalSignature {
                        stream_path: signature_part_name,
                        stream_kind: formula_vba::VbaSignatureStreamKind::Unknown,
                        signer_subject,
                        signature: bytes.to_vec(),
                        verification,
                        binding: formula_vba::VbaSignatureBinding::Unknown,
                    });
                }
            }
        }
    }

    if let Some(sig) = signature_part_result.as_mut() {
        if sig.verification == VbaSignatureVerification::SignedVerified {
            // For raw signature blobs (`vbaProjectSignature.bin` is not an OLE container), attempt
            // to verify MS-OVBA Contents Hash binding against the actual `vbaProject.bin`.
            if sig.binding == formula_vba::VbaSignatureBinding::Unknown {
                if let Some(vba_project_bin) = vba_project_bin {
                    sig.binding = match formula_vba::verify_vba_project_signature_binding(vba_project_bin, &sig.signature) {
                        Ok(binding) => match upgrade_vba_project_signature_binding(vba_project_bin, binding) {
                            VbaProjectBindingVerification::BoundVerified(_) => {
                                formula_vba::VbaSignatureBinding::Bound
                            }
                            VbaProjectBindingVerification::BoundMismatch(_) => {
                                formula_vba::VbaSignatureBinding::NotBound
                            }
                            VbaProjectBindingVerification::BoundUnknown(_) => {
                                formula_vba::VbaSignatureBinding::Unknown
                            }
                        },
                        Err(_) => formula_vba::VbaSignatureBinding::Unknown,
                    };
                }
            }
            return Ok(signature_part_result);
        }
    }

    // Fall back to inspecting `xl/vbaProject.bin` for embedded signature streams.
    let Some(vba_project_bin) = vba_project_bin else {
        return Ok(signature_part_result);
    };

    let embedded = match formula_vba::verify_vba_digital_signature(vba_project_bin) {
        Ok(sig) => sig,
        Err(err) => {
            // If we got anything useful from the signature part, return it rather than failing.
            if signature_part_result.is_some() {
                return Ok(signature_part_result);
            }
            return Err(err);
        }
    };

    if embedded
        .as_ref()
        .is_some_and(|sig| sig.verification == VbaSignatureVerification::SignedVerified)
    {
        return Ok(embedded);
    }

    Ok(signature_part_result.or(embedded))
}

fn verify_vba_digital_signature_with_trust_from_parts(
    parts: &BTreeMap<String, Vec<u8>>,
    options: &formula_vba::VbaSignatureTrustOptions,
) -> Result<Option<formula_vba::VbaDigitalSignatureTrusted>, formula_vba::SignatureError> {
    // Unlike the non-trusty helper, require that the workbook actually contains a VBA project.
    // Trust-center semantics are only meaningful in the context of `xl/vbaProject.bin`.
    let Some(vba_project_bin) = get_part(parts, VBA_PROJECT_BIN) else {
        return Ok(None);
    };

    let signature_part_name = resolve_vba_signature_part_name(parts);
    let mut signature_part_result: Option<formula_vba::VbaDigitalSignatureTrusted> = None;

    if let Some(signature_part_name) = signature_part_name {
        if let Some(bytes) = get_part(parts, &signature_part_name) {
            // Attempt to treat the part as an OLE/CFB container first (the most common format).
            match formula_vba::verify_vba_digital_signature_with_project(vba_project_bin, bytes) {
                Ok(Some(mut sig)) => {
                    // Preserve which part we read the signature from to avoid ambiguity.
                    sig.stream_path = format!("{signature_part_name}:{}", sig.stream_path);

                    let cert_trust = if sig.verification == VbaSignatureVerification::SignedVerified {
                        formula_vba::verify_vba_signature_certificate_trust(&sig.signature, options)
                    } else {
                        formula_vba::VbaCertificateTrust::Unknown
                    };

                    signature_part_result = Some(formula_vba::VbaDigitalSignatureTrusted {
                        signature: sig,
                        cert_trust,
                    });
                }
                Ok(None) => {}
                Err(_) => {
                    // Not an OLE container: fall back to verifying the part bytes as a raw PKCS#7/CMS
                    // signature blob.
                    let (verification, signer_subject) = formula_vba::verify_vba_signature_blob(bytes);

                    let binding = if verification == VbaSignatureVerification::SignedVerified {
                        match formula_vba::verify_vba_project_signature_binding(vba_project_bin, bytes) {
                            Ok(binding) => match upgrade_vba_project_signature_binding(vba_project_bin, binding) {
                                VbaProjectBindingVerification::BoundVerified(_) => {
                                    formula_vba::VbaSignatureBinding::Bound
                                }
                                VbaProjectBindingVerification::BoundMismatch(_) => {
                                    formula_vba::VbaSignatureBinding::NotBound
                                }
                                VbaProjectBindingVerification::BoundUnknown(_) => {
                                    formula_vba::VbaSignatureBinding::Unknown
                                }
                            },
                            Err(_) => formula_vba::VbaSignatureBinding::Unknown,
                        }
                    } else {
                        formula_vba::VbaSignatureBinding::Unknown
                    };

                    let cert_trust = if verification == VbaSignatureVerification::SignedVerified {
                        formula_vba::verify_vba_signature_certificate_trust(bytes, options)
                    } else {
                        formula_vba::VbaCertificateTrust::Unknown
                    };

                    signature_part_result = Some(formula_vba::VbaDigitalSignatureTrusted {
                        signature: VbaDigitalSignature {
                            stream_path: signature_part_name,
                            stream_kind: formula_vba::VbaSignatureStreamKind::Unknown,
                            signer_subject,
                            signature: bytes.to_vec(),
                            verification,
                            binding,
                        },
                        cert_trust,
                    });
                }
            }
        }
    }

    if signature_part_result
        .as_ref()
        .is_some_and(|sig| sig.signature.verification == VbaSignatureVerification::SignedVerified)
    {
        return Ok(signature_part_result);
    }

    // Fall back to inspecting `xl/vbaProject.bin` for embedded signature streams.
    let embedded = match formula_vba::verify_vba_digital_signature_with_trust(vba_project_bin, options) {
        Ok(sig) => sig,
        Err(err) => {
            // If we got anything useful from the signature part, return it rather than failing.
            if signature_part_result.is_some() {
                return Ok(signature_part_result);
            }
            return Err(err);
        }
    };

    if embedded
        .as_ref()
        .is_some_and(|sig| sig.signature.verification == VbaSignatureVerification::SignedVerified)
    {
        return Ok(embedded);
    }

    Ok(signature_part_result.or(embedded))
}

fn resolve_vba_signature_part_name(parts: &BTreeMap<String, Vec<u8>>) -> Option<String> {
    // Prefer explicit relationship resolution when available.
    if let Some(rels_bytes) = get_part(parts, VBA_PROJECT_BIN_RELS) {
        if let Ok(rels) = crate::openxml::parse_relationships(rels_bytes) {
            for rel in rels {
                if rel
                    .target_mode
                    .as_deref()
                    .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
                {
                    continue;
                }

                if rel.type_uri != VBA_PROJECT_SIGNATURE_REL_TYPE {
                    continue;
                }

                let target = strip_fragment(&rel.target);
                let resolved = crate::path::resolve_target(VBA_PROJECT_BIN, target);
                if get_part(parts, &resolved).is_some() {
                    return Some(resolved);
                }
            }
        }
    }

    // Default part name used by Excel.
    if get_part(parts, VBA_PROJECT_SIGNATURE_BIN).is_some() {
        return Some(VBA_PROJECT_SIGNATURE_BIN.to_string());
    }

    None
}

fn upgrade_vba_project_signature_binding(
    project_ole: &[u8],
    binding: formula_vba::VbaProjectBindingVerification,
) -> formula_vba::VbaProjectBindingVerification {
    let formula_vba::VbaProjectBindingVerification::BoundUnknown(mut debug) = binding else {
        return binding;
    };

    let Some(signed_digest) = debug.signed_digest.as_deref() else {
        return formula_vba::VbaProjectBindingVerification::BoundUnknown(debug);
    };

    match signed_digest.len() {
        16 => {
            let Ok(content_hash_md5) = formula_vba::content_hash_md5(project_ole) else {
                return formula_vba::VbaProjectBindingVerification::BoundUnknown(debug);
            };
            let Ok(agile_hash_md5) = formula_vba::agile_content_hash_md5(project_ole) else {
                return formula_vba::VbaProjectBindingVerification::BoundUnknown(debug);
            };
            let Some(agile_hash_md5) = agile_hash_md5 else {
                // If we can't compute FormsNormalizedData, we can't definitively rule out an
                // Agile binding mismatch/match.
                return formula_vba::VbaProjectBindingVerification::BoundUnknown(debug);
            };

            let content_bytes = content_hash_md5.as_slice();
            let agile_bytes = agile_hash_md5.as_slice();

            if signed_digest == content_bytes || signed_digest == agile_bytes {
                // Use the matching digest as the "computed" value for debug display.
                debug.computed_digest = Some(signed_digest.to_vec());
                formula_vba::VbaProjectBindingVerification::BoundVerified(debug)
            } else {
                debug.computed_digest = Some(content_hash_md5.to_vec());
                formula_vba::VbaProjectBindingVerification::BoundMismatch(debug)
            }
        }
        32 => {
            let Ok(computed) = formula_vba::contents_hash_v3(project_ole) else {
                return formula_vba::VbaProjectBindingVerification::BoundUnknown(debug);
            };
            debug.computed_digest = Some(computed.clone());
            if signed_digest == computed.as_slice() {
                formula_vba::VbaProjectBindingVerification::BoundVerified(debug)
            } else {
                formula_vba::VbaProjectBindingVerification::BoundMismatch(debug)
            }
        }
        _ => formula_vba::VbaProjectBindingVerification::BoundUnknown(debug),
    }
}

fn strip_fragment(target: &str) -> &str {
    target.split_once('#').map(|(t, _)| t).unwrap_or(target)
}

fn get_part<'a>(parts: &'a BTreeMap<String, Vec<u8>>, name: &str) -> Option<&'a [u8]> {
    parts
        .get(name)
        .map(Vec::as_slice)
        .or_else(|| {
            name.strip_prefix('/')
                .or_else(|| name.strip_prefix('\\'))
                .and_then(|name| parts.get(name).map(Vec::as_slice))
        })
        .or_else(|| {
            if name.starts_with('/') || name.starts_with('\\') {
                return None;
            }
            // Some producers incorrectly store OPC part names with a leading `/` in the ZIP.
            // Preserve exact names for round-trip, but make lookups resilient.
            let mut with_slash = String::new();
            if with_slash.try_reserve(name.len().saturating_add(1)).is_err() {
                return None;
            }
            with_slash.push('/');
            with_slash.push_str(name);
            parts.get(with_slash.as_str()).map(Vec::as_slice)
        })
        .or_else(|| {
            parts
                .iter()
                .find(|(key, _)| crate::zip_util::zip_part_names_equivalent(key.as_str(), name))
                .map(|(_, bytes)| bytes.as_slice())
        })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn get_part_tolerates_noncanonical_zip_entries() {
        let mut parts = BTreeMap::new();
        parts.insert("XL\\VBAPROJECT.BIN".to_string(), b"dummy".to_vec());

        let found = get_part(&parts, "xl/vbaProject.bin").expect("part should be found");
        assert_eq!(found, b"dummy");
    }
}
