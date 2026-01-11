use serde::{Deserialize, Serialize};

/// Excel workbook date system used to interpret serial date values.
///
/// Excel supports two base date systems:
/// - `Excel1900` (default on Windows; includes the Lotus 1-2-3 leap year bug)
/// - `Excel1904` (default on older Mac versions)
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub enum DateSystem {
    #[serde(rename = "excel1900")]
    Excel1900,
    #[serde(rename = "excel1904")]
    Excel1904,
}

impl Default for DateSystem {
    fn default() -> Self {
        Self::Excel1900
    }
}

impl From<DateSystem> for formula_format::DateSystem {
    fn from(value: DateSystem) -> Self {
        match value {
            DateSystem::Excel1900 => formula_format::DateSystem::Excel1900,
            DateSystem::Excel1904 => formula_format::DateSystem::Excel1904,
        }
    }
}

impl From<formula_format::DateSystem> for DateSystem {
    fn from(value: formula_format::DateSystem) -> Self {
        match value {
            formula_format::DateSystem::Excel1900 => DateSystem::Excel1900,
            formula_format::DateSystem::Excel1904 => DateSystem::Excel1904,
        }
    }
}
