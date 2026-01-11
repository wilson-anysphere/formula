use core::fmt;

use serde::de::Error as _;
use serde::{Deserialize, Serialize};

use crate::drawings::ImageStore;
use crate::{
    rewrite_sheet_names_in_formula, CalcSettings, SheetVisibility, Style, StyleTable, TabColor,
    Table, Worksheet, WorksheetId,
};

/// Identifier for a workbook.
pub type WorkbookId = u32;

fn default_schema_version() -> u32 {
    crate::SCHEMA_VERSION
}

/// A workbook containing worksheets and shared style resources.
#[derive(Clone, Debug, Serialize)]
pub struct Workbook {
    /// Serialization schema version.
    #[serde(default = "default_schema_version")]
    pub schema_version: u32,

    /// Workbook identifier (optional; higher layers may assign meaning).
    #[serde(default)]
    pub id: WorkbookId,

    /// Worksheets contained in the workbook.
    #[serde(default)]
    pub sheets: Vec<Worksheet>,

    /// Workbook style table (deduplicated).
    #[serde(default)]
    pub styles: StyleTable,

    /// Workbook image store (shared across all sheets).
    #[serde(default)]
    pub images: ImageStore,

    /// Workbook calculation options.
    #[serde(default)]
    pub calc_settings: CalcSettings,

    /// Next worksheet id to allocate (runtime-only).
    #[serde(skip)]
    next_sheet_id: WorksheetId,
}

/// Errors raised when renaming a worksheet.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum RenameSheetError {
    SheetNotFound,
    EmptyName,
    DuplicateName,
}

impl fmt::Display for RenameSheetError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            RenameSheetError::SheetNotFound => f.write_str("sheet not found"),
            RenameSheetError::EmptyName => f.write_str("sheet name cannot be empty"),
            RenameSheetError::DuplicateName => f.write_str("sheet name already exists"),
        }
    }
}

impl std::error::Error for RenameSheetError {}

impl Default for Workbook {
    fn default() -> Self {
        Self::new()
    }
}

impl Workbook {
    /// Create a new empty workbook.
    pub fn new() -> Self {
        Self {
            schema_version: crate::SCHEMA_VERSION,
            id: 0,
            sheets: Vec::new(),
            styles: StyleTable::new(),
            images: ImageStore::default(),
            calc_settings: CalcSettings::default(),
            next_sheet_id: 1,
        }
    }

    /// Add a worksheet, returning its id.
    pub fn add_sheet(&mut self, name: impl Into<String>) -> WorksheetId {
        let id = self.next_sheet_id;
        self.next_sheet_id = self.next_sheet_id.wrapping_add(1);
        self.sheets.push(Worksheet::new(id, name));
        id
    }

    /// Rename a worksheet and rewrite formulas that reference it.
    pub fn rename_sheet(
        &mut self,
        id: WorksheetId,
        new_name: &str,
    ) -> Result<(), RenameSheetError> {
        let new_name = new_name.trim();
        if new_name.is_empty() {
            return Err(RenameSheetError::EmptyName);
        }

        let sheet_index = self
            .sheets
            .iter()
            .position(|s| s.id == id)
            .ok_or(RenameSheetError::SheetNotFound)?;

        for sheet in &self.sheets {
            if sheet.id != id && sheet.name.eq_ignore_ascii_case(new_name) {
                return Err(RenameSheetError::DuplicateName);
            }
        }

        let old_name = self.sheets[sheet_index].name.clone();

        for sheet in &mut self.sheets {
            for (_, cell) in sheet.iter_cells_mut() {
                let Some(formula) = cell.formula.clone() else {
                    continue;
                };
                cell.formula = Some(rewrite_sheet_names_in_formula(
                    &formula, &old_name, new_name,
                ));
            }
        }

        self.sheets[sheet_index].name = new_name.to_string();
        Ok(())
    }

    /// Reorder a worksheet within the workbook's sheet list.
    pub fn reorder_sheet(&mut self, id: WorksheetId, new_index: usize) -> bool {
        let Some(current) = self.sheets.iter().position(|s| s.id == id) else {
            return false;
        };
        if new_index >= self.sheets.len() {
            return false;
        }
        if current == new_index {
            return true;
        }
        let sheet = self.sheets.remove(current);
        self.sheets.insert(new_index, sheet);
        true
    }

    /// Set sheet visibility.
    pub fn set_sheet_visibility(&mut self, id: WorksheetId, visibility: SheetVisibility) -> bool {
        let Some(sheet) = self.sheet_mut(id) else {
            return false;
        };
        sheet.visibility = visibility;
        true
    }

    /// Set sheet tab color.
    pub fn set_sheet_tab_color(&mut self, id: WorksheetId, tab_color: Option<TabColor>) -> bool {
        let Some(sheet) = self.sheet_mut(id) else {
            return false;
        };
        sheet.tab_color = tab_color;
        true
    }

    /// Get a sheet by id.
    pub fn sheet(&self, id: WorksheetId) -> Option<&Worksheet> {
        self.sheets.iter().find(|s| s.id == id)
    }

    /// Get a mutable sheet by id.
    pub fn sheet_mut(&mut self, id: WorksheetId) -> Option<&mut Worksheet> {
        self.sheets.iter_mut().find(|s| s.id == id)
    }

    /// Find a sheet by name (case-insensitive, like Excel).
    pub fn sheet_by_name(&self, name: &str) -> Option<&Worksheet> {
        self.sheets.iter().find(|s| s.name.eq_ignore_ascii_case(name))
    }

    /// Find a table by its workbook-scoped name.
    pub fn find_table(&self, table_name: &str) -> Option<(&Worksheet, &Table)> {
        for sheet in &self.sheets {
            if let Some(table) = sheet
                .tables
                .iter()
                .find(|t| t.name.eq_ignore_ascii_case(table_name) || t.display_name.eq_ignore_ascii_case(table_name))
            {
                return Some((sheet, table));
            }
        }
        None
    }

    /// Intern (deduplicate) a style into the workbook's style table.
    pub fn intern_style(&mut self, style: Style) -> u32 {
        self.styles.intern(style)
    }
}

impl<'de> Deserialize<'de> for Workbook {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: serde::Deserializer<'de>,
    {
        #[derive(Deserialize)]
        struct Helper {
            #[serde(default = "default_schema_version")]
            schema_version: u32,
            #[serde(default)]
            id: WorkbookId,
            #[serde(default)]
            sheets: Vec<Worksheet>,
            #[serde(default)]
            styles: StyleTable,
            #[serde(default)]
            images: ImageStore,
            #[serde(default)]
            calc_settings: CalcSettings,
        }

        let helper = Helper::deserialize(deserializer)?;

        if helper.schema_version > crate::SCHEMA_VERSION {
            return Err(D::Error::custom(format!(
                "unsupported schema_version {} (max supported: {})",
                helper.schema_version,
                crate::SCHEMA_VERSION
            )));
        }

        let next_sheet_id = helper
            .sheets
            .iter()
            .map(|s| s.id)
            .max()
            .unwrap_or(0)
            .wrapping_add(1);

        Ok(Workbook {
            schema_version: helper.schema_version,
            id: helper.id,
            sheets: helper.sheets,
            styles: helper.styles,
            images: helper.images,
            calc_settings: helper.calc_settings,
            next_sheet_id,
        })
    }
}
