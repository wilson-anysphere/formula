use serde::{Deserialize, Serialize};

use crate::{CellRef, Range};

/// Hyperlink target destination.
///
/// This is the logical "where does this link go?" independent of where the
/// hyperlink is anchored (which is represented by [`Hyperlink::range`]).
#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(tag = "type", rename_all = "snake_case")]
pub enum HyperlinkTarget {
    /// A URL opened by the host OS (typically `http:` or `https:`).
    ExternalUrl { uri: String },

    /// An internal location within the same workbook (sheet + cell).
    Internal { sheet: String, cell: CellRef },

    /// An email link (typically `mailto:`).
    Email { uri: String },
}

impl HyperlinkTarget {
    pub(crate) fn rewrite_sheet_references(&mut self, old_name: &str, new_name: &str) {
        let HyperlinkTarget::Internal { sheet, .. } = self else {
            return;
        };
        if crate::sheet_name::sheet_name_eq_case_insensitive(sheet, old_name) {
            *sheet = new_name.to_string();
        }
    }
}

/// A hyperlink anchored to a cell or range.
#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub struct Hyperlink {
    /// Cell(s) this hyperlink is attached to.
    pub range: Range,

    /// Destination.
    pub target: HyperlinkTarget,

    /// Optional override for the displayed text (Excel's `display` attribute).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub display: Option<String>,

    /// Optional tooltip shown on hover (Excel's `tooltip` attribute).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tooltip: Option<String>,

    /// Optional XLSX relationship id (`r:id`) used for external hyperlinks.
    ///
    /// This is preserved on load/save to avoid churn and to maintain internal
    /// references within the OpenXML package.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rel_id: Option<String>,
}

impl Hyperlink {
    /// Convenience: create a hyperlink anchored to a single cell.
    pub fn for_cell(cell: CellRef, target: HyperlinkTarget) -> Self {
        Self {
            range: Range::new(cell, cell),
            target,
            display: None,
            tooltip: None,
            rel_id: None,
        }
    }
}
