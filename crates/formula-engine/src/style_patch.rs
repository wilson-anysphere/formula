use std::collections::HashMap;

use formula_model::HorizontalAlignment;
use serde::{Deserialize, Serialize};

/// A style patch entry mirroring DocumentController / JS `styleTable` semantics.
///
/// Unlike `formula_model::Style`, patch objects rely on **property presence** semantics:
/// - missing keys mean "no override"
/// - explicit `null` means "clear" (override to empty / default)
/// - booleans must distinguish absent vs `false` (explicit clear)
///
/// This type only models the subset of formatting/protection metadata needed by worksheet
/// information functions (`CELL`/`INFO`) for now.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct StylePatch {
    /// DocumentController: `numberFormat`
    ///
    /// Tri-state semantics:
    /// - `None`: key absent (no override)
    /// - `Some(None)`: key present with `null` (explicit clear)
    /// - `Some(Some(s))`: set number format string
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub number_format: Option<Option<String>>,

    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub alignment: Option<AlignmentPatch>,

    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub protection: Option<ProtectionPatch>,

    // Optional for future correctness (explicit false overrides).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font: Option<FontPatch>,
}

#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct AlignmentPatch {
    /// DocumentController: `alignment.horizontal`
    ///
    /// Tri-state semantics:
    /// - `None`: key absent (no override)
    /// - `Some(None)`: key present with `null` (explicit clear)
    /// - `Some(Some(v))`: set alignment value
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub horizontal: Option<Option<HorizontalAlignment>>,
}

#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ProtectionPatch {
    /// DocumentController: `protection.locked`
    ///
    /// Tri-state semantics:
    /// - `None`: key absent (no override)
    /// - `Some(None)`: key present with `null` (explicit clear)
    /// - `Some(Some(v))`: set locked value
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub locked: Option<Option<bool>>,
}

#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct FontPatch {
    /// DocumentController: `font.bold`
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bold: Option<bool>,
    /// DocumentController: `font.italic`
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub italic: Option<bool>,
    /// DocumentController: `font.underline`
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub underline: Option<bool>,
    /// DocumentController: `font.strike`
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub strike: Option<bool>,
}

/// A run-length encoded formatting layer entry (DocumentController `formatRunsByCol`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct FormatRun {
    pub start_row: u32,
    pub end_row_exclusive: u32,
    pub style_id: u32,
}

/// Engine-side style table holding patch entries, keyed by `style_id`.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct StylePatchTable {
    patches: HashMap<u32, StylePatch>,
}

impl StylePatchTable {
    pub fn new() -> Self {
        Self {
            patches: HashMap::new(),
        }
    }

    pub fn insert(&mut self, style_id: u32, patch: StylePatch) {
        if style_id == 0 {
            // Style id 0 is always the implicit default/empty patch.
            return;
        }
        self.patches.insert(style_id, patch);
    }

    pub fn remove(&mut self, style_id: u32) {
        if style_id == 0 {
            return;
        }
        self.patches.remove(&style_id);
    }

    pub fn get(&self, style_id: u32) -> Option<&StylePatch> {
        if style_id == 0 {
            return None;
        }
        self.patches.get(&style_id)
    }
}

/// Tuple of contributing style IDs across formatting layers.
///
/// Precedence order matches DocumentController: `sheet < col < row < range-run < cell`.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct CellStyleLayers {
    pub sheet: u32,
    pub col: u32,
    pub row: u32,
    pub range_run: u32,
    pub cell: u32,
}

impl CellStyleLayers {
    pub fn in_precedence_order(self) -> [u32; 5] {
        [self.sheet, self.col, self.row, self.range_run, self.cell]
    }
}

/// Effective style values needed by worksheet information functions.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EffectiveStyle {
    pub number_format: Option<String>,
    pub alignment_horizontal: Option<HorizontalAlignment>,
    /// Effective `protection.locked`.
    ///
    /// Note: Excel defaults this to `true` when unspecified.
    pub locked: bool,
}

impl Default for EffectiveStyle {
    fn default() -> Self {
        Self {
            number_format: None,
            alignment_horizontal: None,
            locked: true,
        }
    }
}

/// Resolve effective style fields for a cell given the layered `style_id` tuple.
///
/// Later layers override earlier layers when they *specify* a property (including specifying
/// `null` to clear).
pub fn resolve_effective_style(table: &StylePatchTable, layers: CellStyleLayers) -> EffectiveStyle {
    let mut number_format: Option<String> = None;
    let mut alignment_horizontal: Option<HorizontalAlignment> = None;
    let mut locked: Option<bool> = None;

    for style_id in layers.in_precedence_order() {
        let Some(patch) = table.get(style_id) else {
            continue;
        };

        if let Some(value) = &patch.number_format {
            number_format = value.clone();
        }

        if let Some(alignment) = &patch.alignment {
            if let Some(value) = alignment.horizontal {
                alignment_horizontal = value;
            }
        }

        if let Some(protection) = &patch.protection {
            if let Some(value) = protection.locked {
                locked = value;
            }
        }
    }

    EffectiveStyle {
        number_format,
        alignment_horizontal,
        locked: locked.unwrap_or(true),
    }
}
