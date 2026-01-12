//! Integration glue between `formula-xlsx` and `formula-vba`.

use std::collections::BTreeMap;

use crate::{XlsxDocument, XlsxPackage};

pub use formula_vba::{
    SignatureError, VbaDigitalSignature, VbaSignatureVerification, VBAModule, VBAProject,
    VBAReference,
};

impl XlsxPackage {
    /// Parse and return a structured VBA project model (for UI display).
    pub fn vba_project(&self) -> Result<Option<VBAProject>, formula_vba::ParseError> {
        let Some(bin) = self.vba_project_bin() else {
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
}

impl XlsxDocument {
    /// Parse and return a structured VBA project model (for UI display).
    pub fn vba_project(&self) -> Result<Option<VBAProject>, formula_vba::ParseError> {
        let Some(bin) = self.parts().get("xl/vbaProject.bin") else {
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
}

fn parse_vba_digital_signature_from_parts(
    parts: &BTreeMap<String, Vec<u8>>,
) -> Result<Option<VbaDigitalSignature>, SignatureError> {
    if let Some(signature_part_name) = resolve_vba_signature_part_name(parts) {
        if let Some(bytes) = parts.get(&signature_part_name) {
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
                    let signer_subject = formula_vba::extract_signer_certificate_info(bytes)
                        .map(|info| info.subject);
                    if signer_subject.is_some() {
                        return Ok(Some(VbaDigitalSignature {
                            stream_path: signature_part_name,
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

    let Some(vba_bin) = parts.get("xl/vbaProject.bin") else {
        return Ok(None);
    };
    formula_vba::parse_vba_digital_signature(vba_bin)
}

fn verify_vba_digital_signature_from_parts(
    parts: &BTreeMap<String, Vec<u8>>,
) -> Result<Option<VbaDigitalSignature>, SignatureError> {
    let signature_part_name = resolve_vba_signature_part_name(parts);
    let mut signature_part_result: Option<VbaDigitalSignature> = None;

    if let Some(signature_part_name) = signature_part_name {
        if let Some(bytes) = parts.get(&signature_part_name) {
            // Attempt to treat the part as an OLE/CFB container first (the most common format).
            match formula_vba::verify_vba_digital_signature(bytes) {
                Ok(Some(mut sig)) => {
                    // Preserve which part we read the signature from to avoid ambiguity.
                    sig.stream_path = format!("{signature_part_name}:{}", sig.stream_path);
                    // `vbaProjectSignature.bin` does not contain the full VBA project streams, so the
                    // MS-OVBA "binding" digest check cannot be meaningfully evaluated here.
                    sig.binding = formula_vba::VbaSignatureBinding::Unknown;
                    signature_part_result = Some(sig);
                }
                Ok(None) => {}
                Err(_) => {
                    // Not an OLE container: fall back to verifying the part bytes as a raw PKCS#7/CMS
                    // signature blob.
                    let (verification, signer_subject) = formula_vba::verify_vba_signature_blob(bytes);
                    signature_part_result = Some(VbaDigitalSignature {
                        stream_path: signature_part_name,
                        signer_subject,
                        signature: bytes.to_vec(),
                        verification,
                        binding: formula_vba::VbaSignatureBinding::Unknown,
                    });
                }
            }
        }
    }

    if signature_part_result
        .as_ref()
        .is_some_and(|sig| sig.verification == VbaSignatureVerification::SignedVerified)
    {
        return Ok(signature_part_result);
    }

    // Fall back to inspecting `xl/vbaProject.bin` for embedded signature streams.
    let Some(vba_project_bin) = parts.get("xl/vbaProject.bin") else {
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

fn resolve_vba_signature_part_name(parts: &BTreeMap<String, Vec<u8>>) -> Option<String> {
    // Prefer explicit relationship resolution when available.
    if let Some(rels_bytes) = parts.get("xl/_rels/vbaProject.bin.rels") {
        if let Ok(rels) = crate::openxml::parse_relationships(rels_bytes) {
            for rel in rels {
                if rel
                    .target_mode
                    .as_deref()
                    .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
                {
                    continue;
                }

                if !rel
                    .type_uri
                    .to_ascii_lowercase()
                    .contains("vbaprojectsignature")
                {
                    continue;
                }

                let target = strip_fragment(&rel.target);
                let resolved = crate::path::resolve_target("xl/vbaProject.bin", target);
                if parts.contains_key(&resolved) {
                    return Some(resolved);
                }
            }
        }
    }

    // Default part name used by Excel.
    if parts.contains_key("xl/vbaProjectSignature.bin") {
        return Some("xl/vbaProjectSignature.bin".to_string());
    }

    None
}

fn strip_fragment(target: &str) -> &str {
    target.split_once('#').map(|(t, _)| t).unwrap_or(target)
}
