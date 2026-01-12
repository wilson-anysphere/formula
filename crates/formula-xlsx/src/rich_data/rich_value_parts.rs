//! Convenience wrapper for loading the four `richValue*` parts from an XLSX package.
//!
//! Excel stores RichData payloads for rich values across multiple parts under `xl/richData/`.
//! This module focuses specifically on the *core* rich-value tables:
//! - `xl/richData/richValue.xml`
//! - `xl/richData/richValueRel.xml`
//! - `xl/richData/richValueTypes.xml`
//! - `xl/richData/richValueStructure.xml`
//!
//! The higher-level cell binding lives elsewhere (`xl/metadata.xml` VM mapping).

use crate::{XlsxError, XlsxPackage};

use super::rich_value::{RichValues, RICH_VALUE_XML};
use super::rich_value_rel::{RichValueRels, RICH_VALUE_REL_XML};
use super::rich_value_structure::{parse_rich_value_structure_from_package, RichValueStructures};
use super::rich_value_types::{parse_rich_value_types_from_package, RichValueTypes};

/// Parsed `richValue*` tables from an [`XlsxPackage`].
#[derive(Debug, Clone, Default, PartialEq)]
pub struct RichValueParts {
    pub rich_value: Option<RichValues>,
    pub rich_value_rel: Option<RichValueRels>,
    pub rich_value_types: Option<RichValueTypes>,
    pub rich_value_structure: Option<RichValueStructures>,
}

impl RichValueParts {
    /// Parse all present `richValue*` parts from the package.
    ///
    /// Missing parts yield `None`; malformed XML yields an error.
    pub fn from_package(pkg: &XlsxPackage) -> Result<Self, XlsxError> {
        Ok(Self {
            rich_value: RichValues::from_package(pkg)?,
            rich_value_rel: RichValueRels::from_package(pkg)?,
            rich_value_types: parse_rich_value_types_from_package(pkg)?,
            rich_value_structure: parse_rich_value_structure_from_package(pkg)?,
        })
    }

    /// Convenience: return whether the package contains any of the canonical `richValue*` parts.
    pub fn any_present(&self) -> bool {
        self.rich_value.is_some()
            || self.rich_value_rel.is_some()
            || self.rich_value_types.is_some()
            || self.rich_value_structure.is_some()
    }

    /// Convenience: return the canonical rich value part name.
    pub fn rich_value_part_name() -> &'static str {
        RICH_VALUE_XML
    }

    /// Convenience: return the canonical rich value rel part name.
    pub fn rich_value_rel_part_name() -> &'static str {
        RICH_VALUE_REL_XML
    }
}
