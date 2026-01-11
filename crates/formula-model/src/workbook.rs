use core::fmt;

use serde::de::Error as _;
use serde::{Deserialize, Serialize};

use crate::drawings::ImageStore;
use crate::names::{
    validate_defined_name, DefinedName, DefinedNameError, DefinedNameId, DefinedNameScope,
};
use crate::{
    rewrite_sheet_names_in_formula, CalcSettings, DateSystem, SheetVisibility, Style, StyleTable,
    TabColor, Table, ThemePalette, WorkbookProtection, Worksheet, WorksheetId,
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

    /// Excel workbook date system (1900 vs 1904) used to interpret serial dates.
    #[serde(default)]
    pub date_system: DateSystem,

    /// Workbook theme palette used to resolve `Color::Theme` references.
    #[serde(default, skip_serializing_if = "ThemePalette::is_default")]
    pub theme: ThemePalette,
    /// Workbook protection state (Excel-compatible).
    #[serde(default, skip_serializing_if = "WorkbookProtection::is_default")]
    pub workbook_protection: WorkbookProtection,

    /// Defined names (named ranges / constants / formulas).
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub defined_names: Vec<DefinedName>,

    /// Next worksheet id to allocate (runtime-only).
    #[serde(skip)]
    next_sheet_id: WorksheetId,

    /// Next defined name id to allocate (runtime-only).
    #[serde(skip)]
    next_defined_name_id: DefinedNameId,
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
            date_system: DateSystem::default(),
            theme: ThemePalette::default(),
            workbook_protection: WorkbookProtection::default(),
            defined_names: Vec::new(),
            next_sheet_id: 1,
            next_defined_name_id: 1,
        }
    }

    /// Convenience helper for formatting cell values according to this workbook's
    /// date system.
    pub fn format_options(&self, locale: formula_format::Locale) -> formula_format::FormatOptions {
        formula_format::FormatOptions {
            locale,
            date_system: self.date_system.into(),
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
            if sheet.id != id
                && crate::formula_rewrite::sheet_name_eq_case_insensitive(&sheet.name, new_name)
            {
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

        for name in &mut self.defined_names {
            name.refers_to = rewrite_sheet_names_in_formula(&name.refers_to, &old_name, new_name);
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
        self.sheets
            .iter()
            .find(|s| crate::formula_rewrite::sheet_name_eq_case_insensitive(&s.name, name))
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

    /// Create a new defined name (named range / constant / formula).
    ///
    /// `refers_to` is stored without a leading `=`, matching how other formula fields
    /// (e.g. table column formulas) are stored throughout the model.
    pub fn create_defined_name(
        &mut self,
        scope: DefinedNameScope,
        name: impl Into<String>,
        refers_to: impl Into<String>,
        comment: Option<String>,
        hidden: bool,
        xlsx_local_sheet_id: Option<u32>,
    ) -> Result<DefinedNameId, DefinedNameError> {
        let name = name.into();
        let name = name.trim().to_string();
        validate_defined_name(&name).map_err(DefinedNameError::InvalidName)?;

        if let DefinedNameScope::Sheet(sheet_id) = scope {
            if self.sheet(sheet_id).is_none() {
                return Err(DefinedNameError::SheetNotFound(sheet_id));
            }
        }

        if self
            .defined_names
            .iter()
            .any(|n| n.scope == scope && n.name.eq_ignore_ascii_case(&name))
        {
            return Err(DefinedNameError::DuplicateName);
        }

        let refers_to = normalize_refers_to(refers_to.into());

        let id = self.next_defined_name_id;
        self.next_defined_name_id = self.next_defined_name_id.wrapping_add(1);
        self.defined_names.push(DefinedName {
            id,
            name,
            scope,
            refers_to,
            comment,
            hidden,
            xlsx_local_sheet_id,
        });
        Ok(id)
    }

    /// Delete a defined name by id.
    pub fn delete_defined_name(&mut self, id: DefinedNameId) -> bool {
        let Some(idx) = self.defined_names.iter().position(|n| n.id == id) else {
            return false;
        };
        self.defined_names.remove(idx);
        true
    }

    /// Rename a defined name by id.
    pub fn rename_defined_name(
        &mut self,
        id: DefinedNameId,
        new_name: &str,
    ) -> Result<(), DefinedNameError> {
        let new_name = new_name.trim().to_string();
        validate_defined_name(&new_name).map_err(DefinedNameError::InvalidName)?;

        let Some(idx) = self.defined_names.iter().position(|n| n.id == id) else {
            return Err(DefinedNameError::DefinedNameNotFound(id));
        };

        let scope = self.defined_names[idx].scope;
        if self
            .defined_names
            .iter()
            .any(|n| n.id != id && n.scope == scope && n.name.eq_ignore_ascii_case(&new_name))
        {
            return Err(DefinedNameError::DuplicateName);
        }

        self.defined_names[idx].name = new_name;
        Ok(())
    }

    /// Find a defined name by scope and name (case-insensitive, like Excel).
    pub fn get_defined_name(&self, scope: DefinedNameScope, name: &str) -> Option<&DefinedName> {
        self.defined_names
            .iter()
            .find(|n| n.scope == scope && n.name.eq_ignore_ascii_case(name))
    }

    /// List defined names, optionally filtered by scope.
    pub fn list_defined_names(&self, scope: Option<DefinedNameScope>) -> Vec<&DefinedName> {
        self.defined_names
            .iter()
            .filter(|n| scope.map_or(true, |s| n.scope == s))
            .collect()
    }
}

fn normalize_refers_to(refers_to: String) -> String {
    let trimmed = refers_to.trim();
    trimmed
        .strip_prefix('=')
        .unwrap_or(trimmed)
        .to_string()
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
            #[serde(default)]
            date_system: DateSystem,
            #[serde(default)]
            theme: ThemePalette,
            #[serde(default)]
            workbook_protection: WorkbookProtection,
            #[serde(default)]
            defined_names: Vec<DefinedName>,
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

        let next_defined_name_id = helper
            .defined_names
            .iter()
            .map(|n| n.id)
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
            date_system: helper.date_system,
            theme: helper.theme,
            workbook_protection: helper.workbook_protection,
            defined_names: helper.defined_names,
            next_sheet_id,
            next_defined_name_id,
        })
    }
}
