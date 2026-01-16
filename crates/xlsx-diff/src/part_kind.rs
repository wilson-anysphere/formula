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

    // Relationship parts can exist in many directories; treat them as package plumbing.
    if crate::ascii::ends_with_ignore_case(part, ".rels") {
        return PartKind::Rels;
    }

    if crate::ascii::starts_with_ignore_case(part, "docprops/") {
        return PartKind::DocProps;
    }

    if part.eq_ignore_ascii_case("xl/workbook.xml") || part.eq_ignore_ascii_case("xl/workbook.bin") {
        return PartKind::Workbook;
    }

    // "Sheets" can also include dialog sheets and macro sheets, but for reporting
    // purposes they are typically grouped with worksheets.
    if crate::ascii::starts_with_ignore_case(part, "xl/worksheets/")
        || crate::ascii::starts_with_ignore_case(part, "xl/dialogsheets/")
        || crate::ascii::starts_with_ignore_case(part, "xl/macrosheets/")
    {
        return PartKind::Worksheet;
    }

    if part.eq_ignore_ascii_case("xl/styles.xml")
        || part.eq_ignore_ascii_case("xl/styles.bin")
        || part.eq_ignore_ascii_case("xl/tablestyles.xml")
    {
        return PartKind::Styles;
    }

    if part.eq_ignore_ascii_case("xl/sharedstrings.xml") || part.eq_ignore_ascii_case("xl/sharedstrings.bin") {
        return PartKind::SharedStrings;
    }

    if crate::ascii::starts_with_ignore_case(part, "xl/theme/") {
        return PartKind::Theme;
    }

    if part.eq_ignore_ascii_case("xl/calcchain.xml") || part.eq_ignore_ascii_case("xl/calcchain.bin") {
        return PartKind::CalcChain;
    }

    if crate::ascii::starts_with_ignore_case(part, "xl/media/") {
        return PartKind::Media;
    }

    if crate::ascii::starts_with_ignore_case(part, "xl/drawings/") {
        return PartKind::Drawings;
    }

    if crate::ascii::starts_with_ignore_case(part, "xl/charts/")
        || crate::ascii::starts_with_ignore_case(part, "xl/chartsheets/")
    {
        return PartKind::Charts;
    }

    if crate::ascii::starts_with_ignore_case(part, "xl/tables/") {
        return PartKind::Tables;
    }

    if crate::ascii::starts_with_ignore_case(part, "xl/pivot") {
        return PartKind::Pivot;
    }

    if part.eq_ignore_ascii_case("xl/vbaProject.bin")
        || part.eq_ignore_ascii_case("xl/vbaProjectSignature.bin")
        || crate::ascii::starts_with_ignore_case(part, "xl/activeX/")
        || crate::ascii::starts_with_ignore_case(part, "xl/ctrlProps/")
        || crate::ascii::starts_with_ignore_case(part, "xl/vba/")
    {
        return PartKind::Vba;
    }

    PartKind::Other
}
