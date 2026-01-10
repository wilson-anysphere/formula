//! Integration glue between `formula-xlsx` and `formula-vba`.

use crate::XlsxPackage;

pub use formula_vba::{VBAProject, VBAModule, VBAReference};

impl XlsxPackage {
    /// Parse and return a structured VBA project model (for UI display).
    pub fn vba_project(&self) -> Result<Option<VBAProject>, formula_vba::ParseError> {
        let Some(bin) = self.vba_project_bin() else {
            return Ok(None);
        };
        Ok(Some(VBAProject::parse(bin)?))
    }
}

