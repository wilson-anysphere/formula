use std::fmt;

use std::borrow::Cow;

/// High-level category for a workbook part (file inside the ZIP/OPC package).
///
/// This is intended for reporting/metrics: consumers can group diffs by workbook
/// area without re-implementing ad-hoc path parsing.
///
/// Part names are normalized internally (leading `/` stripped, `\` converted to
/// `/`, and `.`/`..` segments resolved) before classification.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
#[non_exhaustive]
pub enum PartKind {
    ContentTypes,
    Rels,
    DocProps,
    Workbook,
    Worksheet,
    Styles,
    SharedStrings,
    Theme,
    CalcChain,
    Media,
    Drawings,
    Charts,
    Tables,
    Pivot,
    Vba,
    Other,
}

impl PartKind {
    pub fn as_str(self) -> &'static str {
        match self {
            PartKind::ContentTypes => "content_types",
            PartKind::Rels => "rels",
            PartKind::DocProps => "doc_props",
            PartKind::Workbook => "workbook",
            PartKind::Worksheet => "worksheet",
            PartKind::Styles => "styles",
            PartKind::SharedStrings => "shared_strings",
            PartKind::Theme => "theme",
            PartKind::CalcChain => "calc_chain",
            PartKind::Media => "media",
            PartKind::Drawings => "drawings",
            PartKind::Charts => "charts",
            PartKind::Tables => "tables",
            PartKind::Pivot => "pivot",
            PartKind::Vba => "vba",
            PartKind::Other => "other",
        }
    }
}

impl fmt::Display for PartKind {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(self.as_str())
    }
}

fn maybe_normalize_part_name(part_name: &str) -> Cow<'_, str> {
    // Fast-path common cases to avoid allocation.
    let needs_normalize = part_name.starts_with('/')
        || part_name.contains('\\')
        || part_name.contains("/./")
        || part_name.contains("/../")
        || part_name.contains("//")
        || part_name.starts_with("./")
        || part_name.starts_with("../")
        || part_name == "."
        || part_name == "..";

    if needs_normalize {
        Cow::Owned(crate::normalize_opc_part_name(part_name))
    } else {
        Cow::Borrowed(part_name)
    }
}

/// Classify an OPC part name into a high-level [`PartKind`].
///
/// This helper is intentionally conservative and "shape based": it uses the
/// standardized part directory conventions used by Excel (e.g. `xl/worksheets/`
/// for worksheets) and treats all `.rels` parts as [`PartKind::Rels`].
pub fn classify_part(part_name: &str) -> PartKind {
    let part = maybe_normalize_part_name(part_name);
    let part = part.as_ref();

    if part.eq_ignore_ascii_case("[Content_Types].xml") {
        return PartKind::ContentTypes;
    }

    let lower = part.to_ascii_lowercase();
    let part = lower.as_str();

    // Relationship parts can exist in many directories; treat them as package plumbing.
    if part.ends_with(".rels") {
        return PartKind::Rels;
    }

    if part.starts_with("docprops/") {
        return PartKind::DocProps;
    }

    if part == "xl/workbook.xml" || part == "xl/workbook.bin" {
        return PartKind::Workbook;
    }

    // "Sheets" can also include dialog sheets and macro sheets, but for reporting
    // purposes they are typically grouped with worksheets.
    if part.starts_with("xl/worksheets/")
        || part.starts_with("xl/dialogsheets/")
        || part.starts_with("xl/macrosheets/")
    {
        return PartKind::Worksheet;
    }

    if part == "xl/styles.xml" || part == "xl/styles.bin" || part == "xl/tablestyles.xml" {
        return PartKind::Styles;
    }

    if part == "xl/sharedstrings.xml" || part == "xl/sharedstrings.bin" {
        return PartKind::SharedStrings;
    }

    if part.starts_with("xl/theme/") {
        return PartKind::Theme;
    }

    if part == "xl/calcchain.xml" || part == "xl/calcchain.bin" {
        return PartKind::CalcChain;
    }

    if part.starts_with("xl/media/") {
        return PartKind::Media;
    }

    if part.starts_with("xl/drawings/") {
        return PartKind::Drawings;
    }

    if part.starts_with("xl/charts/") || part.starts_with("xl/chartsheets/") {
        return PartKind::Charts;
    }

    if part.starts_with("xl/tables/") {
        return PartKind::Tables;
    }

    if part.starts_with("xl/pivot") {
        return PartKind::Pivot;
    }

    if part == "xl/vbaproject.bin"
        || part == "xl/vbaprojectsignature.bin"
        || part.starts_with("xl/activex/")
        || part.starts_with("xl/ctrlprops/")
        || part.starts_with("xl/vba/")
    {
        return PartKind::Vba;
    }

    PartKind::Other
}
