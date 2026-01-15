use std::collections::HashMap;

use formula_model::external_refs::{
    escape_external_workbook_name_for_prefix, format_external_key, format_external_span_key,
    format_external_workbook_key,
};
use formula_model::sheet_name_casefold;
#[cfg(feature = "write")]
use formula_model::sheet_name_eq_case_insensitive;

/// Workbook metadata needed to encode/decode sheet-qualified references and defined names.
///
/// In XLSB formulas, 3D references (e.g. `Sheet2!A1` or `Sheet1:Sheet3!A1`) are encoded via an
/// `ixti` index into the workbook's ExternSheet table, and defined names are encoded via their
/// name index. This context provides the forward/backward mappings needed by the rgce codec.
#[derive(Debug, Clone, Default)]
pub struct WorkbookContext {
    /// Maps a sheet or sheet range `(first, last)` to the corresponding ExternSheet index (`ixti`).
    ///
    /// Keys are normalized (Unicode NFKC + uppercasing) for Excel-like case-insensitive lookup.
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
    /// Keys are normalized (Unicode NFKC + uppercasing) for Excel-like case-insensitive lookup.
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
    range: Option<TableRange>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct TableRange {
    /// Normalized (Unicode NFKC + uppercasing) sheet name.
    sheet_key: String,
    /// Bounding box (0-indexed, inclusive) for the table's `ref` range.
    min_row: u32,
    max_row: u32,
    min_col: u32,
    max_col: u32,
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
    /// and encode 3D references like `'[Book2.xlsx]Sheet1'!A1`.
    pub fn add_extern_sheet_external_workbook(
        &mut self,
        workbook: impl Into<String>,
        first_sheet: impl Into<String>,
        last_sheet: impl Into<String>,
        ixti: u16,
    ) {
        let workbook = workbook.into();
        let workbook = escape_external_workbook_name_for_prefix(&workbook).into_owned();
        let first_sheet = first_sheet.into();
        let last_sheet = last_sheet.into();

        // Also populate the forward map used by formula encoders. `formula-engine` represents
        // external workbook refs as `[Book]Sheet`, so store the ExternSheet mapping using the same
        // prefix on both ends of the span.
        let first_key = format_external_key(&workbook, &first_sheet);
        let last_key = format_external_key(&workbook, &last_sheet);
        self.extern_sheets
            .insert((normalize_key(&first_key), normalize_key(&last_key)), ixti);

        self.extern_sheet_targets_rev.insert(
            ixti,
            ExternSheetTarget {
                workbook: Some(workbook),
                first_sheet,
                last_sheet,
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
                let display = display_supbook_name(&supbook.raw_name);
                let book = escape_external_workbook_name_for_prefix(&display);

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
                        let sheet_token = format_external_key(book.as_ref(), sheet_name);
                        return Some(format!(
                            "{}!{}",
                            quote_excel_quoted_ident(&sheet_token),
                            extern_name.name
                        ));
                    }
                }

                // Sheet-range scoped external defined names are encoded with a sheet span `ixti`
                // (index into ExternSheet), but do not carry a single `scope_sheet`.
                //
                // When we can resolve the ixti to a span, render it as
                // `'[Book]SheetA:SheetB'!Name` so the prefix is parseable as a single quoted token
                // (mirrors `rgce`'s canonical 3D-ref formatting).
                if extern_name.scope_sheet.is_none() {
                    if let Some((Some(workbook), first_sheet, last_sheet)) =
                        self.extern_sheet_target(ixti)
                    {
                        if normalize_key(first_sheet) != normalize_key(last_sheet) {
                            let token = format_external_span_key(workbook, first_sheet, last_sheet);
                            return Some(format!(
                                "{}!{}",
                                quote_excel_quoted_ident(&token),
                                extern_name.name
                            ));
                        }
                    }
                }

                // Workbook-scoped external names use the Excel form `[Book]Name`, but the
                // formula-engine parser currently can't disambiguate `[Book]Name` from a
                // structured reference. Quote the entire token so it becomes a `QuotedIdent`.
                let token = format!(
                    "{}{}",
                    format_external_workbook_key(book.as_ref()),
                    extern_name.name
                );
                Some(quote_excel_quoted_ident(&token))
            }
            SupBookKind::Internal => {
                // Sheet-range scoped internal defined names are encoded with a 3D `ixti` (sheet
                // span). Excel's canonical formula text uses `Sheet1:Sheet3!Name`, but to keep
                // the prefix parseable for `formula-engine` we emit the combined span as a single
                // quoted identifier: `'Sheet1:Sheet3'!Name`.
                if extern_name.scope_sheet.is_none() {
                    if let Some((None, first_sheet, last_sheet)) = self.extern_sheet_target(ixti) {
                        if normalize_key(first_sheet) != normalize_key(last_sheet) {
                            let token = format!("{first_sheet}:{last_sheet}");
                            return Some(format!(
                                "{}!{}",
                                quote_excel_quoted_ident(&token),
                                extern_name.name
                            ));
                        }
                    }
                }
                Some(extern_name.name.clone())
            }
            SupBookKind::AddIn => {
                // Excel uses a special SupBook marker (`\u{0001}`) for add-ins and encodes both
                // add-in functions and other extern names via `PtgNameX`.
                //
                // Function extern names should render without qualification (handled above). For
                // *non*-function extern names, include a workbook-like qualifier so the decoded
                // formula text remains unambiguous and round-trippable.
                let display = display_addin_supbook_name(&supbook.raw_name);
                let book = escape_external_workbook_name_for_prefix(&display);

                if let Some(scope_sheet) = extern_name.scope_sheet {
                    let sheet_name = self
                        .namex_supbook_sheets
                        .get(supbook_index as usize)
                        .and_then(|sheets| {
                            sheets.get(scope_sheet as usize).or_else(|| {
                                scope_sheet
                                    .checked_sub(1)
                                    .and_then(|i| sheets.get(i as usize))
                            })
                        });

                    if let Some(sheet_name) = sheet_name {
                        let sheet_token = format_external_key(book.as_ref(), sheet_name);
                        return Some(format!(
                            "{}!{}",
                            quote_excel_quoted_ident(&sheet_token),
                            extern_name.name
                        ));
                    }
                }

                let token = format!(
                    "{}{}",
                    format_external_workbook_key(book.as_ref()),
                    extern_name.name
                );
                Some(quote_excel_quoted_ident(&token))
            }
            _ => Some(extern_name.name.clone()),
        }
    }

    pub(crate) fn namex_defined_name_index_for_ixti(&self, ixti: u16, name: &str) -> Option<u16> {
        let supbook_index = self.namex_ixti_supbooks.get(&ixti).copied().unwrap_or(ixti);
        let normalized_name = normalize_key(name);

        // Prefer an ExternName without a single-sheet scope when encoding 3D sheet spans like
        // `Sheet1:Sheet3!MyName`. Some files also include sheet-scoped ExternName entries for the
        // same display name; keep a best-effort fallback for those cases.
        let mut best_no_scope: Option<u16> = None;
        let mut best_any: Option<u16> = None;

        for (&(sb, name_index), extern_name) in &self.namex_extern_names {
            if sb != supbook_index {
                continue;
            }
            if extern_name.is_function {
                continue;
            }
            if normalize_key(&extern_name.name) != normalized_name {
                continue;
            }

            if extern_name.scope_sheet.is_none() {
                best_no_scope = Some(best_no_scope.map_or(name_index, |best| best.min(name_index)));
            }
            best_any = Some(best_any.map_or(name_index, |best| best.min(name_index)));
        }

        best_no_scope.or(best_any)
    }

    pub(crate) fn namex_function_ref(&self, name: &str) -> Option<(u16, u16)> {
        // Excel uses `_xlfn.` as a forward-compat namespace for newer functions.
        //
        // - In XLSX formula text, callers may specify either `XLOOKUP(...)` or `_xlfn.XLOOKUP(...)`.
        // - In BIFF token streams, these functions are encoded via a NameX extern-function entry
        //   paired with the UDF sentinel (`iftab=255`).
        //
        // Some producers store the extern-function name with the `_xlfn.` prefix, while others
        // store the base name only. Normalize both sides by stripping `_xlfn.` so lookups are
        // robust across writers and across `formula-engine` (which strips `_xlfn.` in its AST).
        let normalized = normalize_function_key(name);

        // HashMap iteration order is nondeterministic; pick the lowest `(supbook, name_index)`
        // match so encoding is stable.
        let mut best: Option<(u16, u16)> = None;
        for (&(supbook_index, name_index), extern_name) in &self.namex_extern_names {
            if !extern_name.is_function {
                continue;
            }
            if normalize_function_key(&extern_name.name) != normalized {
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

    #[cfg(feature = "write")]
    pub(crate) fn namex_ref(
        &self,
        workbook: Option<&str>,
        sheet: Option<&str>,
        name: &str,
    ) -> Option<(u16, u16)> {
        let normalized_name = normalize_key(name);
        let normalized_book = workbook.map(normalize_key);

        // HashMap iteration order is nondeterministic; pick the lowest `(supbook, name_index)`
        // match so encoding is stable.
        let mut best: Option<(u16, u16)> = None;

        for (&(supbook_index, name_index), extern_name) in &self.namex_extern_names {
            if normalize_key(&extern_name.name) != normalized_name {
                continue;
            }

            let Some(supbook) = self.namex_supbooks.get(supbook_index as usize) else {
                continue;
            };
            match (normalized_book.as_deref(), &supbook.kind) {
                // Disallow implicit binding to external workbook names when the formula text does
                // not specify a workbook prefix.
                (None, SupBookKind::ExternalWorkbook) => continue,
                (Some(book), SupBookKind::ExternalWorkbook | SupBookKind::Unknown) => {
                    let display = display_supbook_name(&supbook.raw_name);
                    let display = escape_external_workbook_name_for_prefix(&display);
                    if normalize_key(display.as_ref()) != book {
                        continue;
                    }
                }
                (Some(book), SupBookKind::AddIn) => {
                    let display = display_addin_supbook_name(&supbook.raw_name);
                    let display = escape_external_workbook_name_for_prefix(&display);
                    if normalize_key(display.as_ref()) != book {
                        continue;
                    }
                }
                (Some(_), _) => continue,
                (None, _) => {}
            }

            if let Some(wanted_sheet) = sheet {
                let Some(scope_sheet) = extern_name.scope_sheet else {
                    continue;
                };
                let Some(sheets) = self.namex_supbook_sheets.get(supbook_index as usize) else {
                    continue;
                };
                let sheet_name = sheets.get(scope_sheet as usize).or_else(|| {
                    scope_sheet
                        .checked_sub(1)
                        .and_then(|idx| sheets.get(idx as usize))
                });
                let Some(sheet_name) = sheet_name else {
                    continue;
                };
                if !sheet_name_eq_case_insensitive(sheet_name, wanted_sheet) {
                    continue;
                }
            } else if workbook.is_some() && extern_name.scope_sheet.is_some() {
                // Workbook-scoped references should not match sheet-scoped external names.
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
                range: None,
            });
    }

    /// Registers a table column name for structured reference decoding.
    pub fn add_table_column(&mut self, table_id: u32, column_id: u32, name: impl Into<String>) {
        let name = name.into();
        let entry = self.tables.entry(table_id).or_insert_with(|| TableInfo {
            name: format!("Table{table_id}"),
            columns: HashMap::new(),
            range: None,
        });
        entry.columns.insert(column_id, name);
    }

    /// Registers the bounding box (`ref`) for a table on a specific sheet.
    ///
    /// The row/col indices are 0-based and inclusive.
    pub fn add_table_range(
        &mut self,
        table_id: u32,
        sheet: String,
        r1: u32,
        c1: u32,
        r2: u32,
        c2: u32,
    ) {
        let entry = self.tables.entry(table_id).or_insert_with(|| TableInfo {
            name: format!("Table{table_id}"),
            columns: HashMap::new(),
            range: None,
        });

        let (min_row, max_row) = if r1 <= r2 { (r1, r2) } else { (r2, r1) };
        let (min_col, max_col) = if c1 <= c2 { (c1, c2) } else { (c2, c1) };

        entry.range = Some(TableRange {
            sheet_key: normalize_key(&sheet),
            min_row,
            max_row,
            min_col,
            max_col,
        });
    }

    /// Returns the table id for a given cell coordinate, but only when exactly one table range on
    /// the provided sheet contains the cell.
    ///
    /// This is used to infer table-less structured references such as `[@Qty]`.
    pub fn table_id_for_cell(&self, sheet: &str, row: u32, col: u32) -> Option<u32> {
        let wanted_sheet = normalize_key(sheet);

        let mut found: Option<u32> = None;
        for (&table_id, info) in &self.tables {
            let Some(range) = &info.range else {
                continue;
            };
            if range.sheet_key != wanted_sheet {
                continue;
            }
            if row < range.min_row
                || row > range.max_row
                || col < range.min_col
                || col > range.max_col
            {
                continue;
            }
            if found.is_some() {
                // Ambiguous: multiple tables contain this cell.
                return None;
            }
            found = Some(table_id);
        }
        found
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

    /// Returns the table id for a display name.
    ///
    /// Table names are treated case-insensitively to match Excel.
    pub fn table_id_by_name(&self, name: &str) -> Option<u32> {
        let wanted = normalize_key(name);
        // HashMap iteration order is nondeterministic; pick the lowest id so encoding is stable.
        self.tables
            .iter()
            .filter_map(|(&id, info)| (normalize_key(&info.name) == wanted).then_some(id))
            .min()
    }

    /// Returns a column id for a column display name within a table.
    ///
    /// Column names are treated case-insensitively to match Excel.
    pub fn table_column_id_by_name(&self, table_id: u32, name: &str) -> Option<u32> {
        let wanted = normalize_key(name);
        let table = self.tables.get(&table_id)?;
        // HashMap iteration order is nondeterministic; pick the lowest id so encoding is stable.
        table
            .columns
            .iter()
            .filter_map(|(&id, col_name)| (normalize_key(col_name) == wanted).then_some(id))
            .min()
    }

    /// Returns the table id if the workbook context contains exactly one table.
    ///
    /// This supports encoding "table-less" structured references like `[@Col]`, which Excel
    /// interprets as referencing the current row of the containing table.
    pub fn single_table_id(&self) -> Option<u32> {
        if self.tables.len() == 1 {
            self.tables.keys().next().copied()
        } else {
            None
        }
    }

    /// Returns `Some(true)` when the workbook context knows the sheet containing `table_id` and it
    /// matches `sheet` (case-insensitive, Excel-like normalization).
    ///
    /// Returns `None` when the table is unknown or the workbook context does not have a sheet
    /// association for the table (i.e. the table range was not registered via
    /// [`Self::add_table_range`]).
    pub fn table_is_on_sheet(&self, table_id: u32, sheet: &str) -> Option<bool> {
        let info = self.tables.get(&table_id)?;
        let range = info.range.as_ref()?;
        Some(range.sheet_key == normalize_key(sheet))
    }

    /// Returns the index of the first AddIn `SupBook` (when present).
    #[cfg(feature = "write")]
    pub(crate) fn addin_supbook_index(&self) -> Option<u16> {
        self.namex_supbooks
            .iter()
            .enumerate()
            .find_map(|(idx, sb)| (sb.kind == SupBookKind::AddIn).then_some(idx as u16))
    }

    /// Append a new `SupBook` + sheet table entry to the NameX context, returning its index.
    #[cfg(feature = "write")]
    pub(crate) fn push_namex_supbook(&mut self, supbook: SupBook, sheets: Vec<String>) -> u16 {
        let idx = u16::try_from(self.namex_supbooks.len()).unwrap_or(u16::MAX);
        self.namex_supbooks.push(supbook);
        self.namex_supbook_sheets.push(sheets);
        idx
    }

    /// Inserts/overwrites an extern-name entry in the NameX table.
    #[cfg(feature = "write")]
    pub(crate) fn insert_namex_extern_name(
        &mut self,
        supbook_index: u16,
        name_index: u16,
        extern_name: ExternName,
    ) {
        self.namex_extern_names
            .insert((supbook_index, name_index), extern_name);
    }
}

fn normalize_key(s: &str) -> String {
    // Must match Excel's case-insensitive name matching.
    //
    // Reuse the shared Unicode-aware casefold implementation used across the engine/model.
    sheet_name_casefold(s)
}

fn normalize_function_key(s: &str) -> String {
    let key = normalize_key(s);
    key.strip_prefix("_XLFN.").unwrap_or(&key).to_string()
}

pub(crate) fn display_supbook_name(raw: &str) -> String {
    // SUPBOOK values are often file paths. For formula rendering we want a best-effort workbook
    // *basename* without path separators.
    //
    // Some producers also store the workbook in brackets (e.g. `C:\tmp\[Book.xlsx]`) or wrap the
    // entire path in brackets (`[C:\tmp\Book.xlsx]`). Normalize these cases so we can safely
    // produce `[Book]Sheet` prefixes for external workbook references.
    let without_nuls = raw.replace('\0', "");
    let trimmed_full = without_nuls.trim();
    let has_full_wrapper = trimmed_full.starts_with('[') && trimmed_full.ends_with(']');

    let basename = trimmed_full
        .rsplit(['/', '\\'])
        .next()
        .unwrap_or(trimmed_full);
    let trimmed = basename.trim();
    let has_basename_wrapper = trimmed.starts_with('[') && trimmed.ends_with(']');

    let mut inner = trimmed;
    if has_full_wrapper || has_basename_wrapper {
        inner = inner.strip_prefix('[').unwrap_or(inner);
        inner = inner.strip_suffix(']').unwrap_or(inner);
    }
    inner.to_string()
}

fn display_addin_supbook_name(raw: &str) -> String {
    // BIFF uses a special SupBook name (`\u{0001}`) to represent loaded add-ins (XLL / XLAM).
    // The actual add-in file name is not available in this marker form. When rendering formula
    // text for non-function extern names, we still want a stable qualifier so the output is
    // unambiguous and can be parsed by `formula-engine`.
    if raw.is_empty() || raw == "\u{0001}" {
        return "AddIn".to_string();
    }
    display_supbook_name(raw)
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
    fn extern_sheet_index_is_unicode_case_insensitive() {
        let mut ctx = WorkbookContext::default();
        ctx.add_extern_sheet("Ünicode", "Ünicode", 7u16);
        assert_eq!(ctx.extern_sheet_index("ünicode"), Some(7u16));
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

    #[test]
    fn display_supbook_name_strips_paths_brackets_and_nuls() {
        assert_eq!(
            display_supbook_name("C:\\tmp\\[Book2.xlsb]\u{0000}"),
            "Book2.xlsb"
        );
        assert_eq!(
            display_supbook_name("[C:\\tmp\\Book2.xlsb]\u{0000}"),
            "Book2.xlsb"
        );
        assert_eq!(display_supbook_name("[Book2.xlsb]"), "Book2.xlsb");
    }

    #[test]
    fn display_supbook_name_preserves_literal_brackets_in_workbook_names() {
        // Workbook names may contain literal `[` / `]` characters. Preserve these when the input
        // is not wrapper-bracketed.
        assert_eq!(
            display_supbook_name("[LeadingBracket.xlsb"),
            "[LeadingBracket.xlsb"
        );
        assert_eq!(display_supbook_name("Book2.xlsb]"), "Book2.xlsb]");
    }
}
