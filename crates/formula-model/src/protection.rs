use serde::{Deserialize, Serialize};

fn is_false(v: &bool) -> bool {
    !*v
}

fn is_true(v: &bool) -> bool {
    *v
}

/// Excel-compatible worksheet protection state.
///
/// This models the legacy `sheetProtection` element in OOXML as a set of booleans
/// indicating which operations are allowed when protection is enabled.
#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub struct SheetProtection {
    /// Whether the sheet protection is enabled.
    #[serde(default, skip_serializing_if = "is_false")]
    pub enabled: bool,

    /// Allow selecting locked cells while the sheet is protected.
    ///
    /// Excel defaults this to true when protecting a sheet.
    #[serde(
        default = "crate::serde_defaults::default_true",
        skip_serializing_if = "is_true"
    )]
    pub select_locked_cells: bool,

    /// Allow selecting unlocked cells while the sheet is protected.
    ///
    /// Excel defaults this to true when protecting a sheet.
    #[serde(
        default = "crate::serde_defaults::default_true",
        skip_serializing_if = "is_true"
    )]
    pub select_unlocked_cells: bool,

    /// Allow formatting cells.
    #[serde(default, skip_serializing_if = "is_false")]
    pub format_cells: bool,

    /// Allow formatting columns.
    #[serde(default, skip_serializing_if = "is_false")]
    pub format_columns: bool,

    /// Allow formatting rows.
    #[serde(default, skip_serializing_if = "is_false")]
    pub format_rows: bool,

    /// Allow inserting columns.
    #[serde(default, skip_serializing_if = "is_false")]
    pub insert_columns: bool,

    /// Allow inserting rows.
    #[serde(default, skip_serializing_if = "is_false")]
    pub insert_rows: bool,

    /// Allow inserting hyperlinks.
    #[serde(default, skip_serializing_if = "is_false")]
    pub insert_hyperlinks: bool,

    /// Allow deleting columns.
    #[serde(default, skip_serializing_if = "is_false")]
    pub delete_columns: bool,

    /// Allow deleting rows.
    #[serde(default, skip_serializing_if = "is_false")]
    pub delete_rows: bool,

    /// Allow sorting.
    #[serde(default, skip_serializing_if = "is_false")]
    pub sort: bool,

    /// Allow using AutoFilter.
    #[serde(default, skip_serializing_if = "is_false")]
    pub auto_filter: bool,

    /// Allow using PivotTables.
    #[serde(default, skip_serializing_if = "is_false")]
    pub pivot_tables: bool,

    /// Allow editing drawing objects.
    #[serde(default, skip_serializing_if = "is_false")]
    pub edit_objects: bool,

    /// Allow editing scenarios.
    #[serde(default, skip_serializing_if = "is_false")]
    pub edit_scenarios: bool,

    /// Optional legacy password hash (OOXML `sheetProtection password="..."`).
    ///
    /// Excel stores this as a 16-bit hash rendered as 4 hex digits.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub password_hash: Option<u16>,
}

impl Default for SheetProtection {
    fn default() -> Self {
        Self {
            enabled: false,
            select_locked_cells: true,
            select_unlocked_cells: true,
            format_cells: false,
            format_columns: false,
            format_rows: false,
            insert_columns: false,
            insert_rows: false,
            insert_hyperlinks: false,
            delete_columns: false,
            delete_rows: false,
            sort: false,
            auto_filter: false,
            pivot_tables: false,
            edit_objects: false,
            edit_scenarios: false,
            password_hash: None,
        }
    }
}

impl SheetProtection {
    pub fn is_default(v: &Self) -> bool {
        v == &Self::default()
    }
}

/// Excel-compatible workbook protection state.
#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize, Default)]
pub struct WorkbookProtection {
    /// Lock the workbook structure (sheets cannot be added/moved/renamed/deleted).
    #[serde(default, skip_serializing_if = "is_false")]
    pub lock_structure: bool,

    /// Lock workbook windows (legacy feature; rarely used).
    #[serde(default, skip_serializing_if = "is_false")]
    pub lock_windows: bool,

    /// Optional legacy password hash (OOXML `workbookProtection workbookPassword="..."`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub password_hash: Option<u16>,
}

impl WorkbookProtection {
    pub fn is_default(v: &Self) -> bool {
        v == &Self::default()
    }
}

/// Actions gated by worksheet protection.
#[derive(Copy, Clone, Debug, PartialEq, Eq, Hash, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum SheetProtectionAction {
    SelectLockedCells,
    SelectUnlockedCells,
    FormatCells,
    FormatColumns,
    FormatRows,
    InsertColumns,
    InsertRows,
    InsertHyperlinks,
    DeleteColumns,
    DeleteRows,
    Sort,
    AutoFilter,
    PivotTables,
    EditObjects,
    EditScenarios,
}

/// Hash a password using Excel's legacy worksheet/workbook protection algorithm.
///
/// This produces the 16-bit value stored in OOXML attributes such as:
/// - `sheetProtection password="...."`
/// - `workbookProtection workbookPassword="...."`
///
/// The algorithm is a simple XOR scheme and is **not** cryptographically secure.
#[must_use]
pub fn hash_legacy_password(password: &str) -> u16 {
    let mut hash: u16 = 0;
    let mut len: u16 = 0;

    // Excel truncates legacy passwords to 15 characters.
    for (i, ch) in password.encode_utf16().take(15).enumerate() {
        len = len.saturating_add(1);
        let shift = (i + 1) as u32;
        // Rotate within 15 bits.
        let rotated =
            (((ch as u32) << shift) & 0x7FFF) | ((ch as u32) >> (15u32.saturating_sub(shift)));
        hash ^= rotated as u16;
    }

    hash ^= len;
    hash ^= 0xCE4B;
    hash
}

#[must_use]
pub fn verify_legacy_password(password: &str, hash: u16) -> bool {
    hash_legacy_password(password) == hash
}
