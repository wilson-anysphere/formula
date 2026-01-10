use serde::{Deserialize, Serialize};

use crate::{Style, StyleTable, Worksheet, WorksheetId};

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

    /// Next worksheet id to allocate (runtime-only).
    #[serde(skip)]
    next_sheet_id: WorksheetId,
}

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

    /// Get a sheet by id.
    pub fn sheet(&self, id: WorksheetId) -> Option<&Worksheet> {
        self.sheets.iter().find(|s| s.id == id)
    }

    /// Get a mutable sheet by id.
    pub fn sheet_mut(&mut self, id: WorksheetId) -> Option<&mut Worksheet> {
        self.sheets.iter_mut().find(|s| s.id == id)
    }

    /// Find a sheet by name (case sensitive, like Excel).
    pub fn sheet_by_name(&self, name: &str) -> Option<&Worksheet> {
        self.sheets.iter().find(|s| s.name == name)
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
        }

        let helper = Helper::deserialize(deserializer)?;
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
            next_sheet_id,
        })
    }
}
