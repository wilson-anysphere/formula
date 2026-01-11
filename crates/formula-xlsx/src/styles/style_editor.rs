use std::collections::HashMap;

use formula_model::StyleTable;

use super::{StylesPart, StylesPartError};

/// A high-level helper for editing `xl/styles.xml` while working with `formula_model` style IDs.
///
/// Excel stores cell styles as integer indices (`xf` records in `cellXfs`) that are referenced
/// from worksheet cells via `c/@s`. `formula_model`, on the other hand, stores styles in a
/// deduplicated [`StyleTable`] and cells reference a `style_id`.
///
/// `XlsxStylesEditor` bridges those representations by parsing `styles.xml` into a [`StylesPart`]
/// and exposing a small API tailored for patch/save flows.
#[derive(Debug, Clone)]
pub struct XlsxStylesEditor {
    part: StylesPart,
}

impl XlsxStylesEditor {
    /// Parse `xl/styles.xml` bytes (or a default payload when missing), interning existing `cellXfs`
    /// entries into `style_table`.
    pub fn parse_or_default(
        bytes: Option<&[u8]>,
        style_table: &mut StyleTable,
    ) -> Result<Self, StylesPartError> {
        Ok(Self {
            part: StylesPart::parse_or_default(bytes, style_table)?,
        })
    }

    /// Map an XLSX `xf` index (worksheet `c/@s`) to a `formula_model` `style_id`.
    pub fn style_id_for_xf(&self, xf: u32) -> u32 {
        self.part.style_id_for_xf(xf)
    }

    /// Map a `formula_model` `style_id` to an XLSX `xf` index.
    ///
    /// If the style is not present in `styles.xml` it will be appended deterministically so that
    /// existing `xf` indices remain stable.
    pub fn xf_for_style_id(
        &mut self,
        style_id: u32,
        style_table: &StyleTable,
    ) -> Result<u32, StylesPartError> {
        self.part.xf_index_for_style(style_id, style_table)
    }

    /// Ensure every `style_id` in `style_ids` has a corresponding `xf` index.
    ///
    /// The returned map can be used to populate worksheet `c/@s` attributes.
    pub fn ensure_styles_for_style_ids(
        &mut self,
        style_ids: impl IntoIterator<Item = u32>,
        style_table: &StyleTable,
    ) -> Result<HashMap<u32, u32>, StylesPartError> {
        self.part.xf_indices_for_style_ids(style_ids, style_table)
    }

    /// Serialize the updated `xl/styles.xml` payload.
    pub fn to_styles_xml_bytes(&self) -> Vec<u8> {
        self.part.to_xml_bytes()
    }

    /// Access the underlying `StylesPart` (primarily for advanced callers).
    pub fn styles_part(&self) -> &StylesPart {
        &self.part
    }

    /// Mutably access the underlying `StylesPart` (primarily for advanced callers).
    pub fn styles_part_mut(&mut self) -> &mut StylesPart {
        &mut self.part
    }
}

