use serde::{Deserialize, Serialize};

/// Workbook-wide calculation settings.
///
/// These settings are modeled after Excel's workbook calculation options and map to
/// the `calcPr` element in `xl/workbook.xml` for `.xlsx` files.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct CalcSettings {
    /// Workbook calculation mode (automatic vs manual).
    pub calculation_mode: CalculationMode,
    /// Whether the workbook should be recalculated prior to saving.
    ///
    /// XLSX: `calcOnSave`.
    pub calculate_before_save: bool,
    /// Iterative calculation settings (for circular references).
    pub iterative: IterativeCalculationSettings,
    /// When `true`, calculations use full double precision.
    ///
    /// When `false`, the workbook is in "precision as displayed" mode
    /// (Excel: "Set precision as displayed").
    ///
    /// XLSX: `fullPrecision` (1 = full precision, 0 = precision as displayed).
    pub full_precision: bool,
}

impl Default for CalcSettings {
    fn default() -> Self {
        Self {
            calculation_mode: CalculationMode::Automatic,
            // Excel defaults to calculating on save.
            calculate_before_save: true,
            iterative: IterativeCalculationSettings::default(),
            // Excel defaults to full precision.
            full_precision: true,
        }
    }
}

impl CalcSettings {
    #[must_use]
    pub fn is_manual(&self) -> bool {
        self.calculation_mode == CalculationMode::Manual
    }

    #[must_use]
    pub fn is_automatic(&self) -> bool {
        !self.is_manual()
    }
}

/// Excel workbook calculation mode (`calcMode`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum CalculationMode {
    /// Excel: `auto`.
    Automatic,
    /// Excel: `autoNoTable`.
    ///
    /// Treated like [`Automatic`] by the current engine implementation but preserved
    /// when round-tripping XLSX.
    AutomaticNoTable,
    /// Excel: `manual`.
    Manual,
}

impl Default for CalculationMode {
    fn default() -> Self {
        Self::Automatic
    }
}

impl CalculationMode {
    #[must_use]
    pub fn as_calc_mode_attr(self) -> &'static str {
        match self {
            Self::Automatic => "auto",
            Self::AutomaticNoTable => "autoNoTable",
            Self::Manual => "manual",
        }
    }

    #[must_use]
    pub fn from_calc_mode_attr(value: &str) -> Option<Self> {
        match value {
            "auto" => Some(Self::Automatic),
            "autoNoTable" => Some(Self::AutomaticNoTable),
            "manual" => Some(Self::Manual),
            _ => None,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Serialize, Deserialize)]
pub struct IterativeCalculationSettings {
    /// Enable iterative calculation to resolve circular references.
    ///
    /// XLSX: `iterative`.
    pub enabled: bool,
    /// Maximum number of iterations.
    ///
    /// XLSX: `iterateCount`.
    pub max_iterations: u32,
    /// Maximum change / convergence tolerance.
    ///
    /// XLSX: `iterateDelta`.
    pub max_change: f64,
}

impl Default for IterativeCalculationSettings {
    fn default() -> Self {
        Self {
            enabled: false,
            // Excel defaults to 100 iterations and 0.001 maximum change.
            max_iterations: 100,
            max_change: 0.001,
        }
    }
}
