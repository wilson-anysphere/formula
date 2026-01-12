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
    /// Reverse lookup for `ixti` -> external target (optional workbook + sheet range).
    ///
    /// This is used to decode 3D references that point at external workbooks, which Excel
    /// represents via the SupBook + ExternSheet tables.
    extern_sheet_targets_rev: HashMap<u16, ExternSheetTarget>,

    /// Maps `(scope, name)` to the defined name index.
    ///
    /// Keys are normalized (ASCII-lowercased) for case-insensitive lookup.
    names: HashMap<NameKey, u32>,
    /// Reverse lookup for name index -> definition (display name + scope).
    names_rev: HashMap<u32, NameDefinition>,

    /// SupBook table backing `PtgNameX` references (external names / add-ins).
    namex_supbooks: Vec<SupBook>,
    /// Sheet name tables for each SupBook, used for sheet-scoped `PtgNameX` external names.
    ///
    /// Indexed by `supbook_index` (parallel to [`Self::namex_supbooks`]), then by sheet index.
    namex_supbook_sheets: Vec<Vec<String>>,
    /// ExternName table keyed by (supbook index, extern name index).
    namex_extern_names: HashMap<(u16, u16), ExternName>,
    /// Map `ixti` (ExternSheet index) -> supbook index for `PtgNameX`.
    namex_ixti_supbooks: HashMap<u16, u16>,

    /// Excel table (ListObject) metadata, keyed by table id.
    ///
    /// Structured references in XLSB formulas (`Table1[Col]`, `[@Col]`, etc.) are encoded using
    /// numeric table + column identifiers. The rgce decoder needs a mapping back to display names
    /// to reconstruct Excel-canonical formula text.
    tables: HashMap<u32, TableInfo>,
}

#[derive(Debug, Clone, Default)]
struct TableInfo {
    name: String,
    columns: HashMap<u32, String>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct ExternSheetTarget {
    workbook: Option<String>,
    first_sheet: String,
    last_sheet: String,
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
        self.extern_sheet_targets_rev.insert(
            ixti,
            ExternSheetTarget {
                workbook: None,
                first_sheet: self
                    .extern_sheets_rev
                    .get(&ixti)
                    .map(|(a, _)| a.clone())
                    .unwrap_or_default(),
                last_sheet: self
                    .extern_sheets_rev
                    .get(&ixti)
                    .map(|(_, b)| b.clone())
                    .unwrap_or_default(),
            },
        );
    }

    /// Registers an ExternSheet table entry targeting an external workbook so formulas can decode
    /// 3D references like `'[Book2.xlsx]Sheet1'!A1`.
    pub fn add_extern_sheet_external_workbook(
        &mut self,
        workbook: impl Into<String>,
        first_sheet: impl Into<String>,
        last_sheet: impl Into<String>,
        ixti: u16,
    ) {
        self.extern_sheet_targets_rev.insert(
            ixti,
            ExternSheetTarget {
                workbook: Some(workbook.into()),
                first_sheet: first_sheet.into(),
                last_sheet: last_sheet.into(),
            },
        );
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

    /// Returns the target of an ExternSheet index (`ixti`).
    ///
    /// The workbook name is `None` for internal references and `Some` for external workbook
    /// references.
    pub fn extern_sheet_target(&self, ixti: u16) -> Option<(Option<&str>, &str, &str)> {
        if let Some(target) = self.extern_sheet_targets_rev.get(&ixti) {
            return Some((
                target.workbook.as_deref(),
                target.first_sheet.as_str(),
                target.last_sheet.as_str(),
            ));
        }
        self.extern_sheets_rev
            .get(&ixti)
            .map(|(a, b)| (None, a.as_str(), b.as_str()))
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
        supbook_sheets: Vec<Vec<String>>,
        extern_names: HashMap<(u16, u16), ExternName>,
        ixti_supbooks: HashMap<u16, u16>,
    ) {
        self.namex_supbooks = supbooks;
        self.namex_supbook_sheets = supbook_sheets;
        self.namex_extern_names = extern_names;
        self.namex_ixti_supbooks = ixti_supbooks;
    }

    pub(crate) fn format_namex(&self, ixti: u16, name_index: u16) -> Option<String> {
        // In BIFF, PtgNameX stores `ixti` (index into ExternSheet). Some writers appear to store a
        // SupBook index directly when the ExternSheet table is missing. Handle both.
        let supbook_index = self.namex_ixti_supbooks.get(&ixti).copied().unwrap_or(ixti);

        let extern_name = self.namex_extern_names.get(&(supbook_index, name_index))?;

        if extern_name.is_function {
            return Some(extern_name.name.clone());
        }

        let supbook = self.namex_supbooks.get(supbook_index as usize)?;
        match supbook.kind {
            SupBookKind::ExternalWorkbook => {
                let book = display_supbook_name(&supbook.raw_name);

                if let Some(scope_sheet) = extern_name.scope_sheet {
                    let sheet_name = self
                        .namex_supbook_sheets
                        .get(supbook_index as usize)
                        .and_then(|sheets| {
                            // `scope_sheet` is commonly 0-based, but some producers may store
                            // it as 1-based. Prefer the direct index and fall back to `-1`.
                            sheets.get(scope_sheet as usize).or_else(|| {
                                scope_sheet
                                    .checked_sub(1)
                                    .and_then(|i| sheets.get(i as usize))
                            })
                        });

                    if let Some(sheet_name) = sheet_name {
                        let sheet_token = format!("[{book}]{sheet_name}");
                        return Some(format!(
                            "{}!{}",
                            quote_excel_quoted_ident(&sheet_token),
                            extern_name.name
                        ));
                    }
                }

                // Workbook-scoped external names use the Excel form `[Book]Name`, but the
                // formula-engine parser currently can't disambiguate `[Book]Name` from a
                // structured reference. Quote the entire token so it becomes a `QuotedIdent`.
                let token = format!("[{book}]{}", extern_name.name);
                Some(quote_excel_quoted_ident(&token))
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

    // --- Tables / structured references -----------------------------------------------

    /// Registers an Excel table (ListObject) by id.
    pub fn add_table(&mut self, table_id: u32, name: impl Into<String>) {
        let name = name.into();
        self.tables
            .entry(table_id)
            .and_modify(|t| t.name = name.clone())
            .or_insert_with(|| TableInfo {
                name,
                columns: HashMap::new(),
            });
    }

    /// Registers a table column name for structured reference decoding.
    pub fn add_table_column(&mut self, table_id: u32, column_id: u32, name: impl Into<String>) {
        let name = name.into();
        let entry = self.tables.entry(table_id).or_insert_with(|| TableInfo {
            name: format!("Table{table_id}"),
            columns: HashMap::new(),
        });
        entry.columns.insert(column_id, name);
    }

    /// Returns the display name for a table id.
    pub fn table_name(&self, table_id: u32) -> Option<&str> {
        self.tables.get(&table_id).map(|t| t.name.as_str())
    }

    /// Returns the display name for a table column id.
    pub fn table_column_name(&self, table_id: u32, column_id: u32) -> Option<&str> {
        self.tables
            .get(&table_id)
            .and_then(|t| t.columns.get(&column_id))
            .map(|s| s.as_str())
    }
}

fn normalize_key(s: &str) -> String {
    s.to_ascii_lowercase()
}

pub(crate) fn display_supbook_name(raw: &str) -> String {
    raw.rsplit(['/', '\\']).next().unwrap_or(raw).to_string()
}

fn quote_excel_quoted_ident(raw: &str) -> String {
    // Excel escapes embedded `'` by doubling them within a quoted identifier.
    if !raw.contains('\'') {
        return format!("'{raw}'");
    }

    let quote_count = raw.chars().filter(|&ch| ch == '\'').count();
    let mut out = String::with_capacity(raw.len() + quote_count + 2);
    out.push('\'');
    for ch in raw.chars() {
        if ch == '\'' {
            out.push('\'');
            out.push('\'');
        } else {
            out.push(ch);
        }
    }
    out.push('\'');
    out
}

#[cfg(test)]
mod tests {
    use super::*;
    use formula_engine::parse_formula;

    #[test]
    fn format_namex_adds_external_workbook_prefix_for_names() {
        let mut ctx = WorkbookContext::default();

        let supbooks = vec![SupBook {
            raw_name: r"C:\tmp\Book2.xlsb".to_string(),
            kind: SupBookKind::ExternalWorkbook,
        }];
        let extern_names = HashMap::from([(
            (0u16, 1u16),
            ExternName {
                name: "MyName".to_string(),
                is_function: false,
                scope_sheet: None,
            },
        )]);
        let ixti_supbooks = HashMap::from([(0u16, 0u16)]);

        ctx.set_namex_tables(supbooks, vec![Vec::new()], extern_names, ixti_supbooks);

        let txt = ctx.format_namex(0, 1).expect("format");
        assert_eq!(txt, "'[Book2.xlsb]MyName'");
        parse_formula(&format!("={txt}"), Default::default()).expect("should parse");
    }

    #[test]
    fn format_namex_prefers_sheet_scoped_external_names_when_available() {
        let mut ctx = WorkbookContext::default();

        let supbooks = vec![SupBook {
            raw_name: r"C:\tmp\Book2.xlsb".to_string(),
            kind: SupBookKind::ExternalWorkbook,
        }];
        let extern_names = HashMap::from([(
            (0u16, 1u16),
            ExternName {
                name: "MyName".to_string(),
                is_function: false,
                scope_sheet: Some(0),
            },
        )]);
        let ixti_supbooks = HashMap::from([(0u16, 0u16)]);

        ctx.set_namex_tables(
            supbooks,
            vec![vec!["Sheet1".to_string()]],
            extern_names,
            ixti_supbooks,
        );

        let txt = ctx.format_namex(0, 1).expect("format");
        assert_eq!(txt, "'[Book2.xlsb]Sheet1'!MyName");
        parse_formula(&format!("={txt}"), Default::default()).expect("should parse");
    }

    #[test]
    fn namex_function_ref_resolves_ixti_and_index_deterministically() {
        let mut ctx = WorkbookContext::default();

        let supbooks = vec![SupBook {
            raw_name: "\u{0001}".to_string(),
            kind: SupBookKind::AddIn,
        }];
        let extern_names = HashMap::from([
            (
                (0u16, 2u16),
                ExternName {
                    name: "MYFUNC".to_string(),
                    is_function: true,
                    scope_sheet: None,
                },
            ),
            (
                // Lower name index should win.
                (0u16, 1u16),
                ExternName {
                    name: "MyFunc".to_string(),
                    is_function: true,
                    scope_sheet: None,
                },
            ),
        ]);
        let ixti_supbooks = HashMap::from([(5u16, 0u16), (2u16, 0u16)]);

        ctx.set_namex_tables(supbooks, vec![Vec::new()], extern_names, ixti_supbooks);

        // Chooses smallest (supbook, name_index) match and smallest ixti pointing at that supbook.
        assert_eq!(ctx.namex_function_ref("myfunc"), Some((2u16, 1u16)));
    }
}
