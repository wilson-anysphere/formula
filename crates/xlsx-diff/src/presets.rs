use crate::IgnorePathRule;

/// Built-in ignore presets for suppressing known noisy diffs.
///
/// Presets are intentionally **opt-in** so default diffing remains strict.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, clap::ValueEnum)]
#[value(rename_all = "kebab-case")]
pub enum IgnorePreset {
    /// Ignore attributes that tend to churn across Excel re-saves (volatile IDs, rendering hints).
    ExcelVolatileIds,
}

impl IgnorePreset {
    pub fn as_str(self) -> &'static str {
        match self {
            IgnorePreset::ExcelVolatileIds => "excel-volatile-ids",
        }
    }

    pub(crate) fn rules(self) -> &'static [(&'static str, &'static str)] {
        match self {
            IgnorePreset::ExcelVolatileIds => EXCEL_VOLATILE_IDS,
        }
    }

    pub(crate) fn owned_rules(self) -> impl Iterator<Item = IgnorePathRule> {
        self.rules()
            .iter()
            .map(|(part_glob, path_substring)| IgnorePathRule {
                part: Some((*part_glob).to_string()),
                path_substring: (*path_substring).to_string(),
                kind: None,
            })
    }
}

impl std::fmt::Display for IgnorePreset {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str(self.as_str())
    }
}

// NOTE: XML diff paths include resolved namespace URIs (not prefixes). These strings must match
// the format emitted by `xlsx_diff::diff_xml` (e.g. `@{uri}local`).
const EXCEL_VOLATILE_IDS: &[(&str, &str)] = &[
    // Excel assigns volatile revision IDs to many elements (worksheets, drawings, etc).
    // These are emitted as `xr:uid="..."` (and friends) and can change on every save.
    (
        "xl/**/*.xml",
        "@{http://schemas.microsoft.com/office/spreadsheetml/2014/revision}uid",
    ),
    (
        "xl/**/*.xml",
        "@{http://schemas.microsoft.com/office/spreadsheetml/2015/revision2}uid",
    ),
    (
        "xl/**/*.xml",
        "@{http://schemas.microsoft.com/office/spreadsheetml/2016/revision3}uid",
    ),
    // DrawingML image state hint that is frequently rewritten by Excel.
    ("xl/**/*.xml", "@cstate"),
    // Font metric hint that is frequently rewritten by Excel.
    (
        "xl/**/*.xml",
        "@{http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac}dyDescent",
    ),
];
