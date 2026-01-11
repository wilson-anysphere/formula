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

    /// SupBook table backing `PtgNameX` references (external names / add-ins).
    namex_supbooks: Vec<SupBook>,
    /// ExternName table keyed by (supbook index, extern name index).
    namex_extern_names: HashMap<(u16, u16), ExternName>,
    /// Map `ixti` (ExternSheet index) -> supbook index for `PtgNameX`.
    namex_ixti_supbooks: HashMap<u16, u16>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum SupBookKind {
    /// References the current workbook.
    Internal,
    /// References an external workbook (another file).
    ExternalWorkbook,
    /// References an add-in / XLL.
    AddIn,
    /// Unknown / unclassified SupBook.
    Unknown,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SupBook {
    /// Raw SupBook identifier (often a file name/path or special marker).
    pub raw_name: String,
    pub kind: SupBookKind,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExternSheet {
    /// Index into the workbook SupBook table.
    pub supbook_index: u16,
    pub sheet_first: u32,
    pub sheet_last: u32,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExternName {
    pub name: String,
    pub is_function: bool,
    /// Optional sheet scope within the referenced SupBook.
    pub scope_sheet: Option<u32>,
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

    pub(crate) fn set_namex_tables(
        &mut self,
        supbooks: Vec<SupBook>,
        extern_names: HashMap<(u16, u16), ExternName>,
        ixti_supbooks: HashMap<u16, u16>,
    ) {
        self.namex_supbooks = supbooks;
        self.namex_extern_names = extern_names;
        self.namex_ixti_supbooks = ixti_supbooks;
    }

    pub(crate) fn format_namex(&self, ixti: u16, name_index: u16) -> Option<String> {
        // In BIFF, PtgNameX stores `ixti` (index into ExternSheet). Some writers appear to store a
        // SupBook index directly when the ExternSheet table is missing. Handle both.
        let supbook_index = self
            .namex_ixti_supbooks
            .get(&ixti)
            .copied()
            .unwrap_or(ixti);

        let extern_name = self
            .namex_extern_names
            .get(&(supbook_index, name_index))?;

        if extern_name.is_function {
            return Some(extern_name.name.clone());
        }

        match self
            .namex_supbooks
            .get(supbook_index as usize)
            .map(|s| &s.kind)
        {
            Some(SupBookKind::ExternalWorkbook) => {
                let raw = &self.namex_supbooks.get(supbook_index as usize)?.raw_name;
                Some(format!("[{}]{}", display_supbook_name(raw), extern_name.name))
            }
            _ => Some(extern_name.name.clone()),
        }
    }

    pub(crate) fn namex_function_ref(&self, name: &str) -> Option<(u16, u16)> {
        let normalized = normalize_key(name);

        // HashMap iteration order is nondeterministic; pick the lowest `(supbook, name_index)`
        // match so encoding is stable.
        let mut best: Option<(u16, u16)> = None;
        for (&(supbook_index, name_index), extern_name) in &self.namex_extern_names {
            if !extern_name.is_function {
                continue;
            }
            if normalize_key(&extern_name.name) != normalized {
                continue;
            }
            match best {
                None => best = Some((supbook_index, name_index)),
                Some((best_supbook, best_name)) => {
                    if (supbook_index, name_index) < (best_supbook, best_name) {
                        best = Some((supbook_index, name_index));
                    }
                }
            }
        }

        let (supbook_index, name_index) = best?;

        // Find an ExternSheet (`ixti`) that points at this SupBook. If the workbook doesn't have an
        // ExternSheet table we fall back to encoding the SupBook index directly (mirrors
        // `format_namex`).
        let ixti = self
            .namex_ixti_supbooks
            .iter()
            .filter_map(|(&ixti, &sb)| (sb == supbook_index).then_some(ixti))
            .min()
            .unwrap_or(supbook_index);

        Some((ixti, name_index))
    }
}

fn normalize_key(s: &str) -> String {
    s.to_ascii_lowercase()
}

fn display_supbook_name(raw: &str) -> String {
    raw.rsplit(['/', '\\']).next().unwrap_or(raw).to_string()
}
