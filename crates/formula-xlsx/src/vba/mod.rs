//! Integration glue between `formula-xlsx` and `formula-vba`.

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
    /// Signed XLSM files commonly store signature streams in a separate OPC part
    /// (`xl/vbaProjectSignature.bin`). We prefer that part when present, but fall back to
    /// `xl/vbaProject.bin` when the signature part cannot be opened as an OLE/CFB container.
    pub fn parse_vba_digital_signature(&self) -> Result<Option<VbaDigitalSignature>, SignatureError> {
        if let Some(sig_part) = self.vba_project_signature_bin() {
            match formula_vba::parse_vba_digital_signature(sig_part) {
                Ok(sig) => return Ok(sig),
                Err(err) => {
                    if let Some(vba_bin) = self.vba_project_bin() {
                        return formula_vba::parse_vba_digital_signature(vba_bin);
                    }
                    return Err(err);
                }
            }
        }

        let Some(vba_bin) = self.vba_project_bin() else {
            return Ok(None);
        };
        formula_vba::parse_vba_digital_signature(vba_bin)
    }

    /// Inspect and (best-effort) cryptographically verify the VBA project digital signature.
    ///
    /// This mirrors [`formula_vba::verify_vba_digital_signature`], but prefers the dedicated
    /// `xl/vbaProjectSignature.bin` part when present.
    pub fn verify_vba_digital_signature(&self) -> Result<Option<VbaDigitalSignature>, SignatureError> {
        if let Some(sig_part) = self.vba_project_signature_bin() {
            match formula_vba::verify_vba_digital_signature(sig_part) {
                Ok(sig) => return Ok(sig),
                Err(err) => {
                    if let Some(vba_bin) = self.vba_project_bin() {
                        return formula_vba::verify_vba_digital_signature(vba_bin);
                    }
                    return Err(err);
                }
            }
        }

        let Some(vba_bin) = self.vba_project_bin() else {
            return Ok(None);
        };
        formula_vba::verify_vba_digital_signature(vba_bin)
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
        if let Some(sig_part) = self.parts().get("xl/vbaProjectSignature.bin") {
            match formula_vba::parse_vba_digital_signature(sig_part) {
                Ok(sig) => return Ok(sig),
                Err(err) => {
                    if let Some(vba_bin) = self.parts().get("xl/vbaProject.bin") {
                        return formula_vba::parse_vba_digital_signature(vba_bin);
                    }
                    return Err(err);
                }
            }
        }

        let Some(vba_bin) = self.parts().get("xl/vbaProject.bin") else {
            return Ok(None);
        };
        formula_vba::parse_vba_digital_signature(vba_bin)
    }

    /// Inspect and (best-effort) cryptographically verify the VBA project digital signature.
    pub fn verify_vba_digital_signature(&self) -> Result<Option<VbaDigitalSignature>, SignatureError> {
        if let Some(sig_part) = self.parts().get("xl/vbaProjectSignature.bin") {
            match formula_vba::verify_vba_digital_signature(sig_part) {
                Ok(sig) => return Ok(sig),
                Err(err) => {
                    if let Some(vba_bin) = self.parts().get("xl/vbaProject.bin") {
                        return formula_vba::verify_vba_digital_signature(vba_bin);
                    }
                    return Err(err);
                }
            }
        }

        let Some(vba_bin) = self.parts().get("xl/vbaProject.bin") else {
            return Ok(None);
        };
        formula_vba::verify_vba_digital_signature(vba_bin)
    }
}
