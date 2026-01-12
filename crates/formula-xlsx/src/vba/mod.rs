//! Integration glue between `formula-xlsx` and `formula-vba`.

use crate::{XlsxDocument, XlsxPackage};

pub use formula_vba::{VBAProject, VBAModule, VBAReference};

impl XlsxPackage {
    /// Parse and return a structured VBA project model (for UI display).
    pub fn vba_project(&self) -> Result<Option<VBAProject>, formula_vba::ParseError> {
        let Some(bin) = self.vba_project_bin() else {
            return Ok(None);
        };
        Ok(Some(VBAProject::parse(bin)?))
    }

    /// Inspect the embedded VBA project for a digital signature stream.
    ///
    /// Returns `Ok(None)` when the workbook has no `xl/vbaProject.bin` part.
    pub fn parse_vba_digital_signature(
        &self,
    ) -> Result<Option<formula_vba::VbaDigitalSignature>, formula_vba::SignatureError> {
        let Some(bin) = self.vba_project_bin() else {
            return Ok(None);
        };
        formula_vba::parse_vba_digital_signature(bin)
    }

    /// Inspect and (best-effort) cryptographically verify the VBA project digital signature.
    ///
    /// Returns `Ok(None)` when the workbook has no `xl/vbaProject.bin` part.
    pub fn verify_vba_digital_signature(
        &self,
    ) -> Result<Option<formula_vba::VbaDigitalSignature>, formula_vba::SignatureError> {
        let Some(bin) = self.vba_project_bin() else {
            return Ok(None);
        };
        formula_vba::verify_vba_digital_signature(bin)
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

    /// Inspect the embedded VBA project for a digital signature stream.
    ///
    /// Returns `Ok(None)` when the workbook has no `xl/vbaProject.bin` part.
    pub fn parse_vba_digital_signature(
        &self,
    ) -> Result<Option<formula_vba::VbaDigitalSignature>, formula_vba::SignatureError> {
        let Some(bin) = self.parts().get("xl/vbaProject.bin") else {
            return Ok(None);
        };
        formula_vba::parse_vba_digital_signature(bin)
    }

    /// Inspect and (best-effort) cryptographically verify the VBA project digital signature.
    ///
    /// Returns `Ok(None)` when the workbook has no `xl/vbaProject.bin` part.
    pub fn verify_vba_digital_signature(
        &self,
    ) -> Result<Option<formula_vba::VbaDigitalSignature>, formula_vba::SignatureError> {
        let Some(bin) = self.parts().get("xl/vbaProject.bin") else {
            return Ok(None);
        };
        formula_vba::verify_vba_digital_signature(bin)
    }
}
