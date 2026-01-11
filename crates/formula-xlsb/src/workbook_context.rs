use std::collections::HashMap;

/// Workbook metadata needed to encode/decode sheet-qualified references and defined names.
///
/// In XLSB formulas, 3D references (e.g. `Sheet2!A1` or `Sheet1:Sheet3!A1`) are encoded via an
/// `ixti` index into the workbook's ExternSheet table, and defined names are encoded via their
/// name index. This context provides the forward/backward mappings needed by the rgce codec.
#[derive(Debug, Clone, Default)]
pub struct WorkbookContext {
    /// Maps a sheet or sheet range `(first, last)` to the corresponding ExternSheet index (`ixti`).
    ///
    /// Keys are normalized (ASCII-lowercased) for case-insensitive lookup.
    extern_sheets: HashMap<(String, String), u16>,
    /// Reverse lookup for `ixti` -> `(first, last)` display names.
    extern_sheets_rev: HashMap<u16, (String, String)>,

    /// Maps `(scope, name)` to the defined name index.
    ///
    /// Keys are normalized (ASCII-lowercased) for case-insensitive lookup.
    names: HashMap<NameKey, u32>,
    /// Reverse lookup for name index -> definition (display name + scope).
    names_rev: HashMap<u32, NameDefinition>,
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
enum NameKey {
    Workbook(String),
    Sheet { sheet: String, name: String },
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum NameScope {
    Workbook,
    Sheet(String),
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NameDefinition {
    pub index: u32,
    pub name: String,
    pub scope: NameScope,
}

impl WorkbookContext {
    /// Registers an ExternSheet table entry so formulas can encode/decode 3D references.
    pub fn add_extern_sheet(
        &mut self,
        first_sheet: impl Into<String>,
        last_sheet: impl Into<String>,
        ixti: u16,
    ) {
        let first_sheet = first_sheet.into();
        let last_sheet = last_sheet.into();
        let key = (normalize_key(&first_sheet), normalize_key(&last_sheet));
        self.extern_sheets.insert(key, ixti);
        self.extern_sheets_rev
            .insert(ixti, (first_sheet, last_sheet));
    }

    /// Returns the ExternSheet index (`ixti`) for a sheet.
    pub fn extern_sheet_index(&self, sheet: &str) -> Option<u16> {
        self.extern_sheet_range_index(sheet, sheet)
    }

    /// Returns the ExternSheet index (`ixti`) for a sheet range.
    pub fn extern_sheet_range_index(&self, first_sheet: &str, last_sheet: &str) -> Option<u16> {
        self.extern_sheets
            .get(&(normalize_key(first_sheet), normalize_key(last_sheet)))
            .copied()
    }

    /// Returns the `(first_sheet, last_sheet)` names for an ExternSheet index (`ixti`).
    pub fn extern_sheet_names(&self, ixti: u16) -> Option<(&str, &str)> {
        self.extern_sheets_rev
            .get(&ixti)
            .map(|(a, b)| (a.as_str(), b.as_str()))
    }

    /// Registers a workbook-scoped defined name.
    pub fn add_workbook_name(&mut self, name: impl Into<String>, index: u32) {
        let name = name.into();
        let key = NameKey::Workbook(normalize_key(&name));
        self.names.insert(key, index);
        self.names_rev.insert(
            index,
            NameDefinition {
                index,
                name,
                scope: NameScope::Workbook,
            },
        );
    }

    /// Registers a sheet-scoped defined name.
    pub fn add_sheet_name(
        &mut self,
        sheet: impl Into<String>,
        name: impl Into<String>,
        index: u32,
    ) {
        let sheet = sheet.into();
        let name = name.into();
        let key = NameKey::Sheet {
            sheet: normalize_key(&sheet),
            name: normalize_key(&name),
        };
        self.names.insert(key, index);
        self.names_rev.insert(
            index,
            NameDefinition {
                index,
                name,
                scope: NameScope::Sheet(sheet),
            },
        );
    }

    /// Resolves a defined name to its name index.
    ///
    /// If `sheet` is `Some`, only that sheet's scope is considered. If `sheet` is `None`,
    /// workbook-scope names take precedence; otherwise, if there is exactly one matching
    /// sheet-scoped name, it is returned.
    pub fn name_index(&self, name: &str, sheet: Option<&str>) -> Option<u32> {
        let normalized_name = normalize_key(name);
        if let Some(sheet) = sheet {
            return self
                .names
                .get(&NameKey::Sheet {
                    sheet: normalize_key(sheet),
                    name: normalized_name,
                })
                .copied();
        }

        if let Some(idx) = self
            .names
            .get(&NameKey::Workbook(normalized_name.clone()))
            .copied()
        {
            return Some(idx);
        }

        let mut matches = self.names.iter().filter_map(|(k, v)| match k {
            NameKey::Sheet { name, .. } if name == &normalized_name => Some(*v),
            _ => None,
        });
        let first = matches.next()?;
        if matches.next().is_some() {
            // Ambiguous across multiple sheet scopes.
            return None;
        }
        Some(first)
    }

    /// Returns the display information for a defined name index.
    pub fn name_definition(&self, index: u32) -> Option<&NameDefinition> {
        self.names_rev.get(&index)
    }
}

fn normalize_key(s: &str) -> String {
    s.to_ascii_lowercase()
}
