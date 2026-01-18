use core::fmt;

use std::collections::{HashMap, HashSet};

use serde::de::Error as _;
use serde::{Deserialize, Serialize};

use crate::drawings::ImageStore;
use crate::names::{
    validate_defined_name, DefinedName, DefinedNameError, DefinedNameId, DefinedNameScope,
};
use crate::pivots::{
    PivotCacheId, PivotCacheModel, PivotChartModel, PivotDestination, PivotSource, PivotTableModel,
    SlicerModel, TimelineModel,
};
use crate::sheet_name::{validate_sheet_name, SheetNameError};
use crate::table::{validate_table_name, TableError, TableIdentifier};
use crate::value::text_eq_case_insensitive;
use crate::{
    rewrite_deleted_sheet_references_in_formula, rewrite_sheet_names_in_formula,
    rewrite_table_names_in_formula, CalcSettings, DateSystem, ManualPageBreaks, PageSetup,
    PrintTitles, Range, SheetPrintSettings, SheetVisibility, Style, StyleTable, TabColor, Table,
    ThemePalette, WorkbookPrintSettings, WorkbookProtection, WorkbookView, Worksheet, WorksheetId,
};

/// Identifier for a workbook.
pub type WorkbookId = u32;

fn default_schema_version() -> u32 {
    crate::SCHEMA_VERSION
}

fn default_codepage() -> u16 {
    // Excel/Windows "ANSI" codepage.
    1252
}

fn is_default_codepage(codepage: &u16) -> bool {
    *codepage == default_codepage()
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

    /// Workbook "ANSI" text codepage (BIFF `CODEPAGE` record).
    ///
    /// This is used when interpreting legacy 8-bit text (including Excel DBCS `*B` semantics).
    #[serde(
        default = "default_codepage",
        skip_serializing_if = "is_default_codepage",
        alias = "text_codepage",
        alias = "textCodepage"
    )]
    pub codepage: u16,

    /// Workbook theme palette used to resolve `Color::Theme` references.
    #[serde(default, skip_serializing_if = "ThemePalette::is_default")]
    pub theme: ThemePalette,
    /// Workbook protection state (Excel-compatible).
    #[serde(default, skip_serializing_if = "WorkbookProtection::is_default")]
    pub workbook_protection: WorkbookProtection,

    /// Defined names (named ranges / constants / formulas).
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub defined_names: Vec<DefinedName>,

    /// Pivot table definitions.
    #[serde(
        default,
        skip_serializing_if = "Vec::is_empty",
        rename = "pivotTables",
        alias = "pivot_tables"
    )]
    pub pivot_tables: Vec<PivotTableModel>,

    /// Pivot caches (shared across pivot tables and slicers).
    #[serde(
        default,
        skip_serializing_if = "Vec::is_empty",
        rename = "pivotCaches",
        alias = "pivot_caches"
    )]
    pub pivot_caches: Vec<PivotCacheModel>,

    /// Pivot chart definitions bound to pivot tables.
    #[serde(
        default,
        skip_serializing_if = "Vec::is_empty",
        rename = "pivotCharts",
        alias = "pivot_charts"
    )]
    pub pivot_charts: Vec<PivotChartModel>,

    /// Workbook slicers connected to pivots and placed on worksheets.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub slicers: Vec<SlicerModel>,

    /// Workbook timelines connected to pivots and placed on worksheets.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub timelines: Vec<TimelineModel>,

    /// Workbook print settings (print area/titles, page setup, margins, scaling, manual breaks).
    #[serde(default, skip_serializing_if = "WorkbookPrintSettings::is_empty")]
    pub print_settings: WorkbookPrintSettings,

    /// Workbook view state (active sheet tab, window state, etc).
    #[serde(default, skip_serializing_if = "WorkbookView::is_default")]
    pub view: WorkbookView,

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
    InvalidName(SheetNameError),
}

impl fmt::Display for RenameSheetError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            RenameSheetError::SheetNotFound => f.write_str("sheet not found"),
            RenameSheetError::InvalidName(err) => err.fmt(f),
        }
    }
}

impl std::error::Error for RenameSheetError {}

impl From<SheetNameError> for RenameSheetError {
    fn from(err: SheetNameError) -> Self {
        RenameSheetError::InvalidName(err)
    }
}

/// Errors raised when deleting a worksheet.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum DeleteSheetError {
    SheetNotFound,
    CannotDeleteLastSheet,
    AllocationFailure(&'static str),
}

impl fmt::Display for DeleteSheetError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            DeleteSheetError::SheetNotFound => f.write_str("sheet not found"),
            DeleteSheetError::CannotDeleteLastSheet => f.write_str("cannot delete last sheet"),
            DeleteSheetError::AllocationFailure(ctx) => write!(f, "allocation failed ({ctx})"),
        }
    }
}

impl std::error::Error for DeleteSheetError {}

/// Errors raised when duplicating a worksheet.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum DuplicateSheetError {
    SheetNotFound,
    InvalidName(SheetNameError),
    AllocationFailure(&'static str),
}

impl fmt::Display for DuplicateSheetError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            DuplicateSheetError::SheetNotFound => f.write_str("sheet not found"),
            DuplicateSheetError::InvalidName(err) => err.fmt(f),
            DuplicateSheetError::AllocationFailure(ctx) => {
                write!(f, "allocation failed ({ctx})")
            }
        }
    }
}

impl std::error::Error for DuplicateSheetError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            DuplicateSheetError::InvalidName(err) => Some(err),
            _ => None,
        }
    }
}

impl From<SheetNameError> for DuplicateSheetError {
    fn from(err: SheetNameError) -> Self {
        DuplicateSheetError::InvalidName(err)
    }
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
            images: ImageStore::default(),
            calc_settings: CalcSettings::default(),
            date_system: DateSystem::default(),
            codepage: default_codepage(),
            theme: ThemePalette::default(),
            workbook_protection: WorkbookProtection::default(),
            defined_names: Vec::new(),
            pivot_tables: Vec::new(),
            pivot_caches: Vec::new(),
            pivot_charts: Vec::new(),
            slicers: Vec::new(),
            timelines: Vec::new(),
            print_settings: WorkbookPrintSettings::default(),
            view: WorkbookView::default(),
            next_sheet_id: 1,
            next_defined_name_id: 1,
        }
    }

    /// Recompute runtime-only counters and normalize view/print settings.
    ///
    /// This is useful for callers that construct a workbook by directly assigning
    /// `Workbook::sheets` / `Workbook::defined_names` instead of going through the
    /// higher-level mutation APIs (e.g. when loading from an external persistence layer).
    pub fn recompute_runtime_state(&mut self) {
        self.next_sheet_id = self
            .sheets
            .iter()
            .map(|s| s.id)
            .max()
            .unwrap_or(0)
            .wrapping_add(1);

        self.next_defined_name_id = self
            .defined_names
            .iter()
            .map(|n| n.id)
            .max()
            .unwrap_or(0)
            .wrapping_add(1);

        for sheet_settings in &mut self.print_settings.sheets {
            if let Some(sheet) = self.sheets.iter().find(|s| {
                crate::sheet_name::sheet_name_eq_case_insensitive(
                    &s.name,
                    &sheet_settings.sheet_name,
                )
            }) {
                sheet_settings.sheet_name = sheet.name.clone();
            }
        }

        if let Some(active) = self.view.active_sheet_id {
            if self.sheets.iter().all(|s| s.id != active) {
                self.view.active_sheet_id = None;
            }
        }

        // Ensure deterministic ordering for serialization and UX.
        self.sort_print_settings_by_sheet_order();
    }

    /// Convenience helper for formatting cell values according to this workbook's
    /// date system.
    pub fn format_options(&self, locale: formula_format::Locale) -> formula_format::FormatOptions {
        formula_format::FormatOptions {
            locale,
            date_system: self.date_system.into(),
        }
    }

    fn validate_unique_sheet_name(
        &self,
        name: &str,
        exclude_sheet: Option<WorksheetId>,
    ) -> Result<(), SheetNameError> {
        if self.sheets.iter().any(|sheet| {
            exclude_sheet.map_or(true, |exclude| sheet.id != exclude)
                && crate::sheet_name::sheet_name_eq_case_insensitive(&sheet.name, name)
        }) {
            return Err(SheetNameError::DuplicateName);
        }
        Ok(())
    }

    /// Add a worksheet, returning its id.
    pub fn add_sheet(&mut self, name: impl Into<String>) -> Result<WorksheetId, SheetNameError> {
        let name = name.into();
        validate_sheet_name(&name)?;
        self.validate_unique_sheet_name(&name, None)?;

        let id = self.next_sheet_id;
        self.next_sheet_id = self.next_sheet_id.wrapping_add(1);
        self.sheets.push(Worksheet::new(id, name));
        if self.view.active_sheet_id.is_none() {
            self.view.active_sheet_id = Some(id);
        }
        Ok(id)
    }

    fn rewrite_sheet_references(&mut self, old_name: &str, new_name: &str) {
        for sheet in &mut self.sheets {
            for (_, cell) in sheet.iter_cells_mut() {
                if let Some(formula) = cell.formula.as_mut() {
                    *formula = rewrite_sheet_names_in_formula(formula, old_name, new_name);
                }
            }

            for table in &mut sheet.tables {
                table.rewrite_sheet_references(old_name, new_name);
            }

            for rule in &mut sheet.conditional_formatting_rules {
                rule.rewrite_sheet_references(old_name, new_name);
            }

            for link in &mut sheet.hyperlinks {
                link.target.rewrite_sheet_references(old_name, new_name);
            }

            for assignment in &mut sheet.data_validations {
                assignment
                    .validation
                    .rewrite_sheet_references(old_name, new_name);
            }
        }
    }

    /// Returns the active sheet id (Excel `activeTab`), if any.
    pub fn active_sheet_id(&self) -> Option<WorksheetId> {
        let active = self.view.active_sheet_id;
        if let Some(id) = active {
            if self.sheet(id).is_some() {
                return Some(id);
            }
        }
        self.sheets.first().map(|s| s.id)
    }

    /// Returns the active sheet, if any.
    pub fn active_sheet(&self) -> Option<&Worksheet> {
        let id = self.active_sheet_id()?;
        self.sheet(id)
    }

    /// Set the active sheet (Excel `activeTab`).
    pub fn set_active_sheet(&mut self, id: WorksheetId) -> bool {
        if self.sheet(id).is_none() {
            return false;
        }
        self.view.active_sheet_id = Some(id);
        true
    }

    /// Duplicate a worksheet using Excel-like semantics.
    ///
    /// - If `new_name` is `None`, the new sheet is named like Excel would: `Sheet1` → `Sheet1 (2)`.
    /// - Tables are deep-copied and renamed to maintain workbook-wide table name uniqueness.
    /// - Pivot tables whose destination is on the duplicated sheet are duplicated and re-targeted
    ///   to the new sheet. If a duplicated pivot's source points at data that was also duplicated
    ///   (range/table on the source sheet), the pivot cache is duplicated and marked to be
    ///   rebuilt on the next refresh.
    /// - Formulas within the duplicated sheet are rewritten so any explicit reference to the
    ///   source sheet name points at the new sheet, and structured references to duplicated
    ///   tables refer to the renamed copies.
    pub fn duplicate_sheet(
        &mut self,
        source_id: WorksheetId,
        new_name: Option<&str>,
    ) -> Result<WorksheetId, DuplicateSheetError> {
        let source_index = self
            .sheets
            .iter()
            .position(|s| s.id == source_id)
            .ok_or(DuplicateSheetError::SheetNotFound)?;

        let source_name = self.sheets[source_index].name.clone();

        let target_name = match new_name {
            Some(name) => {
                validate_sheet_name(name)?;
                self.validate_unique_sheet_name(name, None)?;
                name.to_string()
            }
            None => generate_duplicate_sheet_name(&source_name, &self.sheets),
        };

        let new_sheet_id = self.next_sheet_id;
        self.next_sheet_id = self.next_sheet_id.wrapping_add(1);

        let mut new_sheet = self.sheets[source_index].clone();
        new_sheet.id = new_sheet_id;
        new_sheet.name = target_name.clone();

        // A duplicated sheet is a new runtime object, so it should not inherit
        // XLSX relationship identifiers.
        new_sheet.xlsx_sheet_id = None;
        new_sheet.xlsx_rel_id = None;

        let mut used_table_names = collect_table_names(&self.sheets);
        let mut next_table_id = next_table_id(&self.sheets);
        let mut table_renames: Vec<(String, String)> = Vec::new();
        let mut table_id_renames: Vec<(u32, u32)> = Vec::new();

        for table in &mut new_sheet.tables {
            let old_id = table.id;
            let old_name = table.name.clone();
            let old_display_name = table.display_name.clone();
            let new_table_name =
                generate_duplicate_table_name(&old_display_name, &mut used_table_names);
            let add_display_name_mapping = !old_display_name.eq_ignore_ascii_case(&old_name);
            table.id = next_table_id;
            next_table_id = next_table_id.wrapping_add(1);
            table.name = new_table_name.clone();
            table.display_name = new_table_name.clone();
            table.relationship_id = None;
            table.part_path = None;

            table_renames.push((old_name, new_table_name.clone()));
            if add_display_name_mapping {
                table_renames.push((old_display_name, new_table_name));
            }
            table_id_renames.push((old_id, table.id));
        }

        // Rewrite formulas within the duplicated sheet only.
        for (_, cell) in new_sheet.iter_cells_mut() {
            let Some(formula) = cell.formula.as_mut() else {
                continue;
            };
            let rewritten =
                crate::formula_rewrite::rewrite_sheet_names_in_formula_internal_refs_only(
                    formula,
                    &source_name,
                    &target_name,
                );
            *formula = rewrite_table_names_in_formula(&rewritten, &table_renames);
        }

        for table in &mut new_sheet.tables {
            table.rewrite_sheet_references_internal_refs_only(&source_name, &target_name);
            table.rewrite_table_references(&table_renames);
        }

        for rule in &mut new_sheet.conditional_formatting_rules {
            rule.rewrite_sheet_references_internal_refs_only(&source_name, &target_name);
            rule.rewrite_table_references(&table_renames);
        }

        for link in &mut new_sheet.hyperlinks {
            link.target
                .rewrite_sheet_references(&source_name, &target_name);
        }

        for assignment in &mut new_sheet.data_validations {
            assignment
                .validation
                .rewrite_sheet_references_internal_refs_only(&source_name, &target_name);
            assignment
                .validation
                .rewrite_table_references(&table_renames);
        }

        let mut defined_name_id_renames: Vec<(DefinedNameId, DefinedNameId)> = Vec::new();
        let scoped_name_count = self
            .defined_names
            .iter()
            .filter(|n| n.scope == DefinedNameScope::Sheet(source_id))
            .count();
        let mut scoped_names: Vec<DefinedName> = Vec::new();
        if scoped_names.try_reserve_exact(scoped_name_count).is_err() {
            debug_assert!(
                false,
                "allocation failed (duplicate sheet scoped names, count={scoped_name_count})"
            );
            return Err(DuplicateSheetError::AllocationFailure(
                "duplicate sheet scoped names",
            ));
        }
        for name in self
            .defined_names
            .iter()
            .filter(|n| n.scope == DefinedNameScope::Sheet(source_id))
        {
            scoped_names.push(name.clone());
        }
        if !scoped_names.is_empty() {
            let mut next_id = self.next_defined_name_id;
            let mut duplicated: Vec<DefinedName> = Vec::new();
            if duplicated.try_reserve_exact(scoped_names.len()).is_err() {
                debug_assert!(
                    false,
                    "allocation failed (duplicate sheet scoped name copies, count={})",
                    scoped_names.len()
                );
                return Err(DuplicateSheetError::AllocationFailure(
                    "duplicate sheet scoped name copies",
                ));
            }
            for mut name in scoped_names {
                let old_id = name.id;
                name.id = next_id;
                next_id = next_id.wrapping_add(1);
                defined_name_id_renames.push((old_id, name.id));
                name.scope = DefinedNameScope::Sheet(new_sheet_id);
                name.xlsx_local_sheet_id = None;
                name.refers_to =
                    crate::formula_rewrite::rewrite_sheet_names_in_formula_internal_refs_only(
                        &name.refers_to,
                        &source_name,
                        &target_name,
                    );
                name.refers_to = rewrite_table_names_in_formula(&name.refers_to, &table_renames);
                duplicated.push(name);
            }
            self.next_defined_name_id = next_id;
            self.defined_names.extend(duplicated);
        }

        // The worksheet struct contains runtime-only caches (e.g. conditional formatting
        // evaluation results). Since we mutate formulas/rules above, drop any copied caches.
        new_sheet.clear_conditional_formatting_cache();

        // Excel inserts the copy immediately after the source sheet.
        self.sheets.insert(source_index + 1, new_sheet);

        let duplicated_pivots = self.duplicate_pivots_for_sheet(
            source_id,
            &source_name,
            new_sheet_id,
            &target_name,
            &table_renames,
            &table_id_renames,
            &defined_name_id_renames,
        )?;
        self.duplicate_pivot_charts_for_sheet(source_id, new_sheet_id, &duplicated_pivots)?;
        self.duplicate_slicers_for_sheet(source_id, new_sheet_id, &duplicated_pivots)?;
        self.duplicate_timelines_for_sheet(source_id, new_sheet_id, &duplicated_pivots)?;

        // Copy print settings (print area/titles, page setup, manual breaks) if present.
        if let Some(settings) = self
            .print_settings
            .sheets
            .iter()
            .find(|s| {
                crate::sheet_name::sheet_name_eq_case_insensitive(&s.sheet_name, &source_name)
            })
            .cloned()
        {
            let mut copied_settings = settings;
            copied_settings.sheet_name = target_name.clone();
            if !copied_settings.is_default() {
                self.print_settings.sheets.push(copied_settings);
            }
        }

        self.sort_print_settings_by_sheet_order();

        // Excel activates the newly inserted sheet.
        self.view.active_sheet_id = Some(new_sheet_id);

        Ok(new_sheet_id)
    }

    fn duplicate_pivots_for_sheet(
        &mut self,
        source_sheet_id: WorksheetId,
        source_sheet_name: &str,
        new_sheet_id: WorksheetId,
        new_sheet_name: &str,
        table_renames: &[(String, String)],
        table_id_renames: &[(u32, u32)],
        defined_name_id_renames: &[(DefinedNameId, DefinedNameId)],
    ) -> Result<HashMap<crate::pivots::PivotTableId, crate::pivots::PivotTableId>, DuplicateSheetError>
    {
        let pivot_count = self
            .pivot_tables
            .iter()
            .filter(|pivot| {
                pivot_destination_is_on_sheet(
                    &pivot.destination,
                    source_sheet_id,
                    source_sheet_name,
                )
            })
            .count();
        if pivot_count == 0 {
            return Ok(HashMap::new());
        }

        let mut pivots_to_duplicate: Vec<PivotTableModel> = Vec::new();
        if pivots_to_duplicate.try_reserve_exact(pivot_count).is_err() {
            debug_assert!(
                false,
                "allocation failed (duplicate sheet pivots, count={pivot_count})"
            );
            return Err(DuplicateSheetError::AllocationFailure(
                "duplicate sheet pivots",
            ));
        }
        for pivot in self.pivot_tables.iter().filter(|pivot| {
            pivot_destination_is_on_sheet(&pivot.destination, source_sheet_id, source_sheet_name)
        }) {
            pivots_to_duplicate.push(pivot.clone());
        }

        let mut id_map: HashMap<crate::pivots::PivotTableId, crate::pivots::PivotTableId> =
            HashMap::new();
        if id_map.try_reserve(pivots_to_duplicate.len()).is_err() {
            debug_assert!(
                false,
                "allocation failed (duplicate sheet pivot id map, count={})",
                pivots_to_duplicate.len()
            );
            return Err(DuplicateSheetError::AllocationFailure(
                "duplicate sheet pivot id map",
            ));
        }

        let mut used_names = collect_pivot_table_names(&self.pivot_tables)?;

        let mut table_id_map: HashMap<u32, u32> = HashMap::new();
        if table_id_map.try_reserve(table_id_renames.len()).is_err() {
            debug_assert!(
                false,
                "allocation failed (duplicate sheet pivot table id map, count={})",
                table_id_renames.len()
            );
            return Err(DuplicateSheetError::AllocationFailure(
                "duplicate sheet pivot table id map",
            ));
        }
        for (old, new) in table_id_renames {
            table_id_map.insert(*old, *new);
        }

        let mut defined_name_id_map: HashMap<DefinedNameId, DefinedNameId> = HashMap::new();
        if defined_name_id_map
            .try_reserve(defined_name_id_renames.len())
            .is_err()
        {
            debug_assert!(
                false,
                "allocation failed (duplicate sheet pivot defined name id map, count={})",
                defined_name_id_renames.len()
            );
            return Err(DuplicateSheetError::AllocationFailure(
                "duplicate sheet pivot defined name id map",
            ));
        }
        for (old, new) in defined_name_id_renames {
            defined_name_id_map.insert(*old, *new);
        }

        for pivot in pivots_to_duplicate {
            let mut duplicated = pivot.clone();
            duplicated.id = crate::new_uuid();
            duplicated.name =
                generate_duplicate_pivot_table_name(&pivot.name, &mut used_names)?;
            id_map.insert(pivot.id, duplicated.id);
            let pivot_had_cache_id = pivot.cache_id.is_some();

            // Retarget the destination to the new sheet.
            match &mut duplicated.destination {
                PivotDestination::Cell { sheet_id, .. }
                | PivotDestination::Range { sheet_id, .. } => {
                    if *sheet_id == source_sheet_id {
                        *sheet_id = new_sheet_id;
                    }
                }
                PivotDestination::CellName { sheet_name, .. }
                | PivotDestination::RangeName { sheet_name, .. } => {
                    if crate::sheet_name::sheet_name_eq_case_insensitive(
                        sheet_name,
                        source_sheet_name,
                    ) {
                        *sheet_name = new_sheet_name.to_string();
                    }
                }
            }

            // Rewrite the source when it points at duplicated data, and split the cache when the
            // original pivot already has one. (When a pivot has no cache id, we keep it that way
            // to mirror Excel's lazy cache materialization behavior.)
            let mut source_changed = false;
            match &mut duplicated.source {
                PivotSource::Range { sheet_id, .. } => {
                    if *sheet_id == source_sheet_id {
                        *sheet_id = new_sheet_id;
                        source_changed = true;
                    }
                }
                PivotSource::RangeName { sheet_name, .. } => {
                    if crate::sheet_name::sheet_name_eq_case_insensitive(
                        sheet_name,
                        source_sheet_name,
                    ) {
                        *sheet_name = new_sheet_name.to_string();
                        source_changed = true;
                    }
                }
                PivotSource::Table { table } => match table {
                    TableIdentifier::Name(name) => {
                        for (old, new) in table_renames {
                            if name.eq_ignore_ascii_case(old) {
                                *name = new.clone();
                                source_changed = true;
                                break;
                            }
                        }
                    }
                    TableIdentifier::Id(id) => {
                        if let Some(new_id) = table_id_map.get(id).copied() {
                            *id = new_id;
                            source_changed = true;
                        }
                    }
                },
                PivotSource::NamedRange { name } => match name {
                    crate::pivots::DefinedNameIdentifier::Id(id) => {
                        if let Some(new_id) = defined_name_id_map.get(id).copied() {
                            *id = new_id;
                            source_changed = true;
                        } else {
                            // Keep workbook-scoped (and unresolved) defined-name sources
                            // referencing the original name, mirroring Excel behavior.
                        }
                    }
                    crate::pivots::DefinedNameIdentifier::Name(_) => {
                        // Preserve unresolved/name-based references; we can't reliably infer scope
                        // when only given a string.
                    }
                },
                PivotSource::DataModel { .. } => {}
            }

            if source_changed && pivot_had_cache_id {
                // If the duplicated pivot's source was rewritten to point at duplicated data
                // (range/table on the new sheet), allocate a distinct cache so the original and
                // duplicated pivots can be refreshed independently (Excel-like).
                //
                // Only duplicate caches when the original pivot already had a cache id; pivots
                // without a cache should keep `cache_id = None` and allow the host/engine to
                // materialize one later.
                let cache_id: PivotCacheId = crate::new_uuid();
                duplicated.cache_id = Some(cache_id);
                self.pivot_caches.push(PivotCacheModel {
                    id: cache_id,
                    source: duplicated.source.clone(),
                    needs_refresh: true,
                });
            }

            self.pivot_tables.push(duplicated);
        }

        Ok(id_map)
    }

    fn duplicate_pivot_charts_for_sheet(
        &mut self,
        source_sheet_id: WorksheetId,
        new_sheet_id: WorksheetId,
        duplicated_pivots: &HashMap<crate::pivots::PivotTableId, crate::pivots::PivotTableId>,
    ) -> Result<(), DuplicateSheetError> {
        let chart_count = self
            .pivot_charts
            .iter()
            .filter(|chart| chart.sheet_id == Some(source_sheet_id))
            .count();
        if chart_count == 0 {
            return Ok(());
        }

        let mut charts_to_duplicate: Vec<PivotChartModel> = Vec::new();
        if charts_to_duplicate.try_reserve_exact(chart_count).is_err() {
            debug_assert!(
                false,
                "allocation failed (duplicate sheet pivot charts, count={chart_count})"
            );
            return Err(DuplicateSheetError::AllocationFailure(
                "duplicate sheet pivot charts",
            ));
        }
        for chart in self
            .pivot_charts
            .iter()
            .filter(|chart| chart.sheet_id == Some(source_sheet_id))
        {
            charts_to_duplicate.push(chart.clone());
        }

        if self.pivot_charts.try_reserve(charts_to_duplicate.len()).is_err() {
            debug_assert!(
                false,
                "allocation failed (duplicate sheet pivot charts append, count={})",
                charts_to_duplicate.len()
            );
            return Err(DuplicateSheetError::AllocationFailure(
                "duplicate sheet pivot charts append",
            ));
        }

        for mut chart in charts_to_duplicate {
            chart.id = crate::new_uuid();
            chart.sheet_id = Some(new_sheet_id);
            if let Some(new_pivot_id) = duplicated_pivots.get(&chart.pivot_table_id) {
                chart.pivot_table_id = *new_pivot_id;
            }
            self.pivot_charts.push(chart);
        }
        Ok(())
    }

    fn duplicate_slicers_for_sheet(
        &mut self,
        source_sheet_id: WorksheetId,
        new_sheet_id: WorksheetId,
        duplicated_pivots: &HashMap<crate::pivots::PivotTableId, crate::pivots::PivotTableId>,
    ) -> Result<(), DuplicateSheetError> {
        let slicer_count = self
            .slicers
            .iter()
            .filter(|slicer| slicer.sheet_id == source_sheet_id)
            .count();
        if slicer_count == 0 {
            return Ok(());
        }

        let mut slicers_to_duplicate: Vec<SlicerModel> = Vec::new();
        if slicers_to_duplicate.try_reserve_exact(slicer_count).is_err() {
            debug_assert!(
                false,
                "allocation failed (duplicate sheet slicers, count={slicer_count})"
            );
            return Err(DuplicateSheetError::AllocationFailure(
                "duplicate sheet slicers",
            ));
        }
        for slicer in self
            .slicers
            .iter()
            .filter(|slicer| slicer.sheet_id == source_sheet_id)
        {
            slicers_to_duplicate.push(slicer.clone());
        }

        if self.slicers.try_reserve(slicers_to_duplicate.len()).is_err() {
            debug_assert!(
                false,
                "allocation failed (duplicate sheet slicers append, count={})",
                slicers_to_duplicate.len()
            );
            return Err(DuplicateSheetError::AllocationFailure(
                "duplicate sheet slicers append",
            ));
        }

        for mut slicer in slicers_to_duplicate {
            slicer.id = crate::new_uuid();
            slicer.sheet_id = new_sheet_id;
            for pivot_id in &mut slicer.connected_pivots {
                if let Some(new_pivot_id) = duplicated_pivots.get(pivot_id) {
                    *pivot_id = *new_pivot_id;
                }
            }
            self.slicers.push(slicer);
        }
        Ok(())
    }

    fn duplicate_timelines_for_sheet(
        &mut self,
        source_sheet_id: WorksheetId,
        new_sheet_id: WorksheetId,
        duplicated_pivots: &HashMap<crate::pivots::PivotTableId, crate::pivots::PivotTableId>,
    ) -> Result<(), DuplicateSheetError> {
        let timeline_count = self
            .timelines
            .iter()
            .filter(|timeline| timeline.sheet_id == source_sheet_id)
            .count();
        if timeline_count == 0 {
            return Ok(());
        }

        let mut timelines_to_duplicate: Vec<TimelineModel> = Vec::new();
        if timelines_to_duplicate.try_reserve_exact(timeline_count).is_err() {
            debug_assert!(
                false,
                "allocation failed (duplicate sheet timelines, count={timeline_count})"
            );
            return Err(DuplicateSheetError::AllocationFailure(
                "duplicate sheet timelines",
            ));
        }
        for timeline in self
            .timelines
            .iter()
            .filter(|timeline| timeline.sheet_id == source_sheet_id)
        {
            timelines_to_duplicate.push(timeline.clone());
        }

        if self.timelines.try_reserve(timelines_to_duplicate.len()).is_err() {
            debug_assert!(
                false,
                "allocation failed (duplicate sheet timelines append, count={})",
                timelines_to_duplicate.len()
            );
            return Err(DuplicateSheetError::AllocationFailure(
                "duplicate sheet timelines append",
            ));
        }

        for mut timeline in timelines_to_duplicate {
            timeline.id = crate::new_uuid();
            timeline.sheet_id = new_sheet_id;
            for pivot_id in &mut timeline.connected_pivots {
                if let Some(new_pivot_id) = duplicated_pivots.get(pivot_id) {
                    *pivot_id = *new_pivot_id;
                }
            }
            self.timelines.push(timeline);
        }
        Ok(())
    }

    fn delete_pivots_for_sheet(&mut self, sheet_id: WorksheetId, sheet_name: &str) {
        self.pivot_tables.retain(|pivot| {
            !pivot_destination_is_on_sheet(&pivot.destination, sheet_id, sheet_name)
        });

        // Remove pivot charts explicitly placed on the deleted sheet.
        self.pivot_charts
            .retain(|chart| chart.sheet_id != Some(sheet_id));

        // Remove pivot charts bound to missing pivot tables.
        self.pivot_charts.retain(|chart| {
            self.pivot_tables
                .iter()
                .any(|pivot| pivot.id == chart.pivot_table_id)
        });

        // Remove slicers/timelines placed on the deleted sheet or left with no remaining pivot
        // connections after pivot deletion.
        self.slicers.retain_mut(|slicer| {
            if slicer.sheet_id == sheet_id {
                return false;
            }
            slicer.connected_pivots.retain(|pivot_id| {
                self.pivot_tables.iter().any(|pivot| pivot.id == *pivot_id)
            });
            !slicer.connected_pivots.is_empty()
        });

        self.timelines.retain_mut(|timeline| {
            if timeline.sheet_id == sheet_id {
                return false;
            }
            timeline.connected_pivots.retain(|pivot_id| {
                self.pivot_tables.iter().any(|pivot| pivot.id == *pivot_id)
            });
            !timeline.connected_pivots.is_empty()
        });

        self.garbage_collect_pivot_caches();
    }

    fn garbage_collect_pivot_caches(&mut self) {
        self.pivot_caches.retain(|cache| {
            self.pivot_tables
                .iter()
                .any(|pivot| pivot.cache_id == Some(cache.id))
        });
    }

    /// Rename a worksheet and rewrite formulas that reference it.
    pub fn rename_sheet(
        &mut self,
        id: WorksheetId,
        new_name: &str,
    ) -> Result<(), RenameSheetError> {
        let sheet_index = self
            .sheets
            .iter()
            .position(|s| s.id == id)
            .ok_or(RenameSheetError::SheetNotFound)?;

        validate_sheet_name(new_name)?;
        self.validate_unique_sheet_name(new_name, Some(id))?;

        let old_name = self.sheets[sheet_index].name.clone();
        if old_name == new_name {
            return Ok(());
        }

        self.rewrite_sheet_references(&old_name, new_name);

        for name in &mut self.defined_names {
            name.refers_to = rewrite_sheet_names_in_formula(&name.refers_to, &old_name, new_name);
        }

        for pivot in &mut self.pivot_tables {
            pivot.source.rewrite_sheet_name(&old_name, new_name);
            pivot.destination.rewrite_sheet_name(&old_name, new_name);
        }
        for cache in &mut self.pivot_caches {
            cache.source.rewrite_sheet_name(&old_name, new_name);
        }

        self.sheets[sheet_index].name = new_name.to_string();

        // Keep print settings aligned with the sheet name (XLSX print settings are keyed by name).
        for settings in &mut self.print_settings.sheets {
            if crate::sheet_name::sheet_name_eq_case_insensitive(
                &settings.sheet_name,
                &old_name,
            ) {
                settings.sheet_name = new_name.to_string();
            }
        }

        Ok(())
    }

    /// Delete a worksheet, rewriting formulas that reference it (Excel-like).
    ///
    /// This preserves all remaining sheets' `xlsx_sheet_id` / `xlsx_rel_id` values; any
    /// renumbering required for serialization is handled by the XLSX writer.
    pub fn delete_sheet(&mut self, id: WorksheetId) -> Result<(), DeleteSheetError> {
        let sheet_index = self
            .sheets
            .iter()
            .position(|s| s.id == id)
            .ok_or(DeleteSheetError::SheetNotFound)?;

        if self.sheets.len() <= 1 {
            return Err(DeleteSheetError::CannotDeleteLastSheet);
        }

        // Capture the pre-delete sheet order for 3D reference adjustment.
        let mut sheet_order: Vec<String> = Vec::new();
        if sheet_order.try_reserve_exact(self.sheets.len()).is_err() {
            debug_assert!(
                false,
                "allocation failed (delete sheet order, count={})",
                self.sheets.len()
            );
            return Err(DeleteSheetError::AllocationFailure("delete sheet order"));
        }
        for s in &self.sheets {
            sheet_order.push(s.name.clone());
        }
        let deleted_name = self.sheets[sheet_index].name.clone();

        // If the deleted sheet was active, Excel selects the nearest neighbor tab.
        let new_active_sheet_id = if self.view.active_sheet_id == Some(id) {
            if sheet_index + 1 < self.sheets.len() {
                Some(self.sheets[sheet_index + 1].id)
            } else if sheet_index > 0 {
                Some(self.sheets[sheet_index - 1].id)
            } else {
                None
            }
        } else {
            None
        };

        self.sheets.remove(sheet_index);

        if new_active_sheet_id.is_some() {
            self.view.active_sheet_id = new_active_sheet_id;
        }

        // Drop any names scoped to the deleted worksheet (Excel removes them).
        self.defined_names
            .retain(|name| name.scope != DefinedNameScope::Sheet(id));

        // Drop print settings for the deleted worksheet.
        self.print_settings.sheets.retain(|s| {
            !crate::sheet_name::sheet_name_eq_case_insensitive(&s.sheet_name, &deleted_name)
        });
        self.sort_print_settings_by_sheet_order();

        // Remove pivot-related objects owned by the deleted sheet and garbage-collect any caches
        // that become unused.
        self.delete_pivots_for_sheet(id, &deleted_name);

        for sheet in &mut self.sheets {
            for (_, cell) in sheet.iter_cells_mut() {
                if let Some(formula) = cell.formula.as_mut() {
                    *formula = rewrite_deleted_sheet_references_in_formula(
                        formula,
                        &deleted_name,
                        &sheet_order,
                    );
                }
            }

            for table in &mut sheet.tables {
                table.invalidate_deleted_sheet_references(&deleted_name, &sheet_order);
            }

            for rule in &mut sheet.conditional_formatting_rules {
                rule.invalidate_deleted_sheet_references(&deleted_name, &sheet_order);
            }

            for assignment in &mut sheet.data_validations {
                assignment
                    .validation
                    .invalidate_deleted_sheet_references(&deleted_name, &sheet_order);
            }
        }

        for name in &mut self.defined_names {
            name.refers_to = rewrite_deleted_sheet_references_in_formula(
                &name.refers_to,
                &deleted_name,
                &sheet_order,
            );
        }

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
        self.sort_print_settings_by_sheet_order();
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
            .find(|s| crate::sheet_name::sheet_name_eq_case_insensitive(&s.name, name))
    }

    /// Find a table by its workbook-scoped name.
    pub fn find_table(&self, table_name: &str) -> Option<(&Worksheet, &Table)> {
        for sheet in &self.sheets {
            if let Some(table) = sheet.tables.iter().find(|t| {
                t.name.eq_ignore_ascii_case(table_name)
                    || t.display_name.eq_ignore_ascii_case(table_name)
            }) {
                return Some((sheet, table));
            }
        }
        None
    }

    /// Find a table by its workbook-scoped name (case-insensitive, like Excel).
    pub fn find_table_case_insensitive(&self, table_name: &str) -> Option<(&Worksheet, &Table)> {
        self.find_table(table_name)
    }

    /// Add a table to `sheet_id`, enforcing Excel table name rules and workbook-wide uniqueness.
    pub fn add_table(&mut self, sheet_id: WorksheetId, mut table: Table) -> Result<(), TableError> {
        table.name = table.name.trim().to_string();
        table.display_name = table.display_name.trim().to_string();
        validate_table_name(&table.name)?;
        validate_table_name(&table.display_name)?;

        if self.find_table_case_insensitive(&table.name).is_some()
            || self
                .find_table_case_insensitive(&table.display_name)
                .is_some()
        {
            return Err(TableError::DuplicateName);
        }

        let sheet = self.sheet_mut(sheet_id).ok_or(TableError::SheetNotFound)?;
        sheet.tables.push(table);
        Ok(())
    }

    /// Remove a table by name from `sheet_id`.
    pub fn remove_table_by_name(
        &mut self,
        sheet_id: WorksheetId,
        table_name: &str,
    ) -> Result<Table, TableError> {
        let sheet = self.sheet_mut(sheet_id).ok_or(TableError::SheetNotFound)?;
        sheet
            .remove_table_by_name(table_name)
            .ok_or(TableError::TableNotFound)
    }

    /// Remove a table by id from `sheet_id`.
    pub fn remove_table_by_id(
        &mut self,
        sheet_id: WorksheetId,
        table_id: u32,
    ) -> Result<Table, TableError> {
        let sheet = self.sheet_mut(sheet_id).ok_or(TableError::SheetNotFound)?;
        sheet
            .remove_table_by_id(table_id)
            .ok_or(TableError::TableNotFound)
    }

    /// Remove a table from `sheet_id` by name or id.
    pub fn remove_table(
        &mut self,
        sheet_id: WorksheetId,
        table: impl Into<TableIdentifier>,
    ) -> Result<Table, TableError> {
        match table.into() {
            TableIdentifier::Name(name) => self.remove_table_by_name(sheet_id, &name),
            TableIdentifier::Id(id) => self.remove_table_by_id(sheet_id, id),
        }
    }

    /// Rename a table (workbook-wide) and rewrite structured references in all formulas.
    pub fn rename_table(&mut self, old_name: &str, new_name: &str) -> Result<(), TableError> {
        let new_name = new_name.trim();
        validate_table_name(new_name)?;

        let (sheet_idx, table_idx) = self
            .sheets
            .iter()
            .enumerate()
            .find_map(|(si, sheet)| {
                sheet
                    .tables
                    .iter()
                    .position(|t| {
                        t.name.eq_ignore_ascii_case(old_name)
                            || t.display_name.eq_ignore_ascii_case(old_name)
                    })
                    .map(|ti| (si, ti))
            })
            .ok_or(TableError::TableNotFound)?;

        let actual_old_name = self.sheets[sheet_idx].tables[table_idx].name.clone();
        let actual_old_display_name = self.sheets[sheet_idx].tables[table_idx]
            .display_name
            .clone();

        for (si, sheet) in self.sheets.iter().enumerate() {
            for (ti, table) in sheet.tables.iter().enumerate() {
                if si == sheet_idx && ti == table_idx {
                    continue;
                }
                if table.name.eq_ignore_ascii_case(new_name)
                    || table.display_name.eq_ignore_ascii_case(new_name)
                {
                    return Err(TableError::DuplicateName);
                }
            }
        }

        let mut renames = vec![(actual_old_name.clone(), new_name.to_string())];
        if !actual_old_display_name.eq_ignore_ascii_case(&actual_old_name) {
            renames.push((actual_old_display_name, new_name.to_string()));
        }

        for sheet in &mut self.sheets {
            for (_, cell) in sheet.iter_cells_mut() {
                if let Some(formula) = cell.formula.as_mut() {
                    *formula = rewrite_table_names_in_formula(formula, &renames);
                }
            }

            for table in &mut sheet.tables {
                table.rewrite_table_references(&renames);
            }

            for rule in &mut sheet.conditional_formatting_rules {
                rule.rewrite_table_references(&renames);
            }

            for assignment in &mut sheet.data_validations {
                assignment.validation.rewrite_table_references(&renames);
            }
        }

        for name in &mut self.defined_names {
            name.refers_to = rewrite_table_names_in_formula(&name.refers_to, &renames);
        }

        for pivot in &mut self.pivot_tables {
            for (old, new) in &renames {
                pivot.source.rewrite_table_name(old, new);
            }
        }
        for cache in &mut self.pivot_caches {
            for (old, new) in &renames {
                cache.source.rewrite_table_name(old, new);
            }
        }

        let renamed = &mut self.sheets[sheet_idx].tables[table_idx];
        renamed.name = new_name.to_string();
        renamed.display_name = new_name.to_string();
        Ok(())
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
            .any(|n| n.scope == scope && text_eq_case_insensitive(&n.name, &name))
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

        let old_name = self.defined_names[idx].name.clone();
        let scope = self.defined_names[idx].scope;
        if self
            .defined_names
            .iter()
            .any(|n| n.id != id && n.scope == scope && text_eq_case_insensitive(&n.name, &new_name))
        {
            return Err(DefinedNameError::DuplicateName);
        }

        self.defined_names[idx].name = new_name;

        for pivot in &mut self.pivot_tables {
            pivot
                .source
                .rewrite_defined_name(&old_name, &self.defined_names[idx].name);
        }
        for cache in &mut self.pivot_caches {
            cache
                .source
                .rewrite_defined_name(&old_name, &self.defined_names[idx].name);
        }

        Ok(())
    }

    /// Find a defined name by scope and name (case-insensitive, like Excel).
    pub fn get_defined_name(&self, scope: DefinedNameScope, name: &str) -> Option<&DefinedName> {
        self.defined_names
            .iter()
            .find(|n| n.scope == scope && text_eq_case_insensitive(&n.name, name))
    }

    /// List defined names, optionally filtered by scope.
    pub fn list_defined_names(&self, scope: Option<DefinedNameScope>) -> Vec<&DefinedName> {
        let count = self
            .defined_names
            .iter()
            .filter(|n| scope.map_or(true, |s| n.scope == s))
            .count();
        let mut out: Vec<&DefinedName> = Vec::new();
        if out.try_reserve_exact(count).is_err() {
            debug_assert!(
                false,
                "allocation failed (list defined names, count={count})"
            );
            return Vec::new();
        }
        for name in self
            .defined_names
            .iter()
            .filter(|n| scope.map_or(true, |s| n.scope == s))
        {
            out.push(name);
        }
        out
    }

    /// Get print settings for a sheet, defaulting when no settings are stored.
    ///
    /// If `id` does not match an existing sheet, this returns default settings with an empty
    /// `sheet_name`.
    pub fn sheet_print_settings(&self, id: WorksheetId) -> SheetPrintSettings {
        let sheet_name = self.sheet(id).map(|s| s.name.as_str()).unwrap_or_default();
        self.sheet_print_settings_by_name(sheet_name)
    }

    /// Get print settings for a sheet by name, defaulting when no settings are stored.
    pub fn sheet_print_settings_by_name(&self, sheet_name: &str) -> SheetPrintSettings {
        let sheet_name = self
            .sheet_by_name(sheet_name)
            .map(|s| s.name.as_str())
            .unwrap_or(sheet_name);

        self.print_settings
            .sheets
            .iter()
            .find(|s| {
                crate::sheet_name::sheet_name_eq_case_insensitive(&s.sheet_name, sheet_name)
            })
            .cloned()
            .map(|mut settings| {
                settings.sheet_name = sheet_name.to_string();
                settings
            })
            .unwrap_or_else(|| SheetPrintSettings::new(sheet_name))
    }

    /// Set (or clear) the print area for a sheet.
    pub fn set_sheet_print_area(
        &mut self,
        id: WorksheetId,
        print_area: Option<Vec<Range>>,
    ) -> bool {
        let Some(sheet_name) = self.sheet(id).map(|s| s.name.clone()) else {
            return false;
        };
        self.set_sheet_print_area_by_name(&sheet_name, print_area)
    }

    /// Set (or clear) the print area for a sheet by name.
    pub fn set_sheet_print_area_by_name(
        &mut self,
        sheet_name: &str,
        print_area: Option<Vec<Range>>,
    ) -> bool {
        let Some(sheet_name) = self.sheet_by_name(sheet_name).map(|s| s.name.clone()) else {
            return false;
        };

        let print_area = print_area
            .and_then(|ranges| (!ranges.is_empty()).then_some(ranges))
            .map(|ranges| {
                let mut out: Vec<Range> = Vec::new();
                if out.try_reserve_exact(ranges.len()).is_err() {
                    debug_assert!(
                        false,
                        "allocation failed (print area normalize, ranges={})",
                        ranges.len()
                    );
                    return Vec::new();
                }
                for r in ranges {
                    out.push(Range::new(r.start, r.end));
                }
                out
            });

        self.update_sheet_print_settings(&sheet_name, |settings| settings.print_area = print_area);
        true
    }

    /// Set (or clear) the print titles for a sheet.
    pub fn set_sheet_print_titles(
        &mut self,
        id: WorksheetId,
        print_titles: Option<PrintTitles>,
    ) -> bool {
        let Some(sheet_name) = self.sheet(id).map(|s| s.name.clone()) else {
            return false;
        };
        self.set_sheet_print_titles_by_name(&sheet_name, print_titles)
    }

    /// Set (or clear) the print titles for a sheet by name.
    pub fn set_sheet_print_titles_by_name(
        &mut self,
        sheet_name: &str,
        print_titles: Option<PrintTitles>,
    ) -> bool {
        let Some(sheet_name) = self.sheet_by_name(sheet_name).map(|s| s.name.clone()) else {
            return false;
        };

        let print_titles = print_titles.map(|t| PrintTitles {
            repeat_rows: t.repeat_rows.map(|r| r.normalized()),
            repeat_cols: t.repeat_cols.map(|c| c.normalized()),
        });

        self.update_sheet_print_settings(&sheet_name, |settings| {
            settings.print_titles = print_titles
        });
        true
    }

    /// Set the page setup for a sheet.
    pub fn set_sheet_page_setup(&mut self, id: WorksheetId, page_setup: PageSetup) -> bool {
        let Some(sheet_name) = self.sheet(id).map(|s| s.name.clone()) else {
            return false;
        };
        self.set_sheet_page_setup_by_name(&sheet_name, page_setup)
    }

    /// Set the page setup for a sheet by name.
    pub fn set_sheet_page_setup_by_name(
        &mut self,
        sheet_name: &str,
        page_setup: PageSetup,
    ) -> bool {
        let Some(sheet_name) = self.sheet_by_name(sheet_name).map(|s| s.name.clone()) else {
            return false;
        };

        self.update_sheet_print_settings(&sheet_name, |settings| settings.page_setup = page_setup);
        true
    }

    /// Set manual page breaks for a sheet.
    pub fn set_manual_page_breaks(
        &mut self,
        id: WorksheetId,
        manual_page_breaks: ManualPageBreaks,
    ) -> bool {
        let Some(sheet_name) = self.sheet(id).map(|s| s.name.clone()) else {
            return false;
        };
        self.set_manual_page_breaks_by_name(&sheet_name, manual_page_breaks)
    }

    /// Set manual page breaks for a sheet by name.
    pub fn set_manual_page_breaks_by_name(
        &mut self,
        sheet_name: &str,
        manual_page_breaks: ManualPageBreaks,
    ) -> bool {
        let Some(sheet_name) = self.sheet_by_name(sheet_name).map(|s| s.name.clone()) else {
            return false;
        };

        self.update_sheet_print_settings(&sheet_name, |settings| {
            settings.manual_page_breaks = manual_page_breaks
        });
        true
    }

    fn update_sheet_print_settings<F: FnOnce(&mut SheetPrintSettings)>(
        &mut self,
        sheet_name: &str,
        update: F,
    ) {
        let idx = self.print_settings.sheets.iter().position(|s| {
            crate::sheet_name::sheet_name_eq_case_insensitive(&s.sheet_name, sheet_name)
        });

        match idx {
            Some(i) => {
                // Canonicalize the stored name to match the workbook sheet name.
                self.print_settings.sheets[i].sheet_name = sheet_name.to_string();
                update(&mut self.print_settings.sheets[i]);
                if self.print_settings.sheets[i].is_default() {
                    self.print_settings.sheets.remove(i);
                }
            }
            None => {
                let mut settings = SheetPrintSettings::new(sheet_name);
                update(&mut settings);
                if !settings.is_default() {
                    self.print_settings.sheets.push(settings);
                }
            }
        }

        self.sort_print_settings_by_sheet_order();
    }

    fn sort_print_settings_by_sheet_order(&mut self) {
        use std::collections::HashMap;

        let mut order: HashMap<&str, usize> = HashMap::new();
        if order.try_reserve(self.sheets.len()).is_err() {
            debug_assert!(
                false,
                "allocation failed (print settings sheet order map, sheets={})",
                self.sheets.len()
            );
            return;
        }
        for (idx, s) in self.sheets.iter().enumerate() {
            order.insert(s.name.as_str(), idx);
        }

        self.print_settings.sheets.sort_by_key(|s| {
            order
                .get(s.sheet_name.as_str())
                .copied()
                .unwrap_or(usize::MAX)
        });
    }
}

fn normalize_refers_to(refers_to: String) -> String {
    let trimmed = refers_to.trim();
    trimmed.strip_prefix('=').unwrap_or(trimmed).to_string()
}

fn generate_duplicate_sheet_name(base: &str, sheets: &[Worksheet]) -> String {
    let mut i: u64 = 2;
    loop {
        let suffix = format!(" ({i})");
        let suffix_len = suffix.encode_utf16().count();
        let max_base_len = crate::sheet_name::EXCEL_MAX_SHEET_NAME_LEN.saturating_sub(suffix_len);
        let mut used_len = 0usize;
        let mut truncated = String::new();
        for ch in base.chars() {
            let ch_len = ch.len_utf16();
            if used_len + ch_len > max_base_len {
                break;
            }
            used_len += ch_len;
            truncated.push(ch);
        }
        let candidate = format!("{truncated}{suffix}");
        if sheets
            .iter()
            .all(|s| !crate::sheet_name::sheet_name_eq_case_insensitive(&s.name, &candidate))
        {
            return candidate;
        }
        i = i.wrapping_add(1);
    }
}

fn collect_table_names(sheets: &[Worksheet]) -> HashSet<String> {
    let mut out = HashSet::new();
    for sheet in sheets {
        for table in &sheet.tables {
            let name_lc = table.name.to_ascii_lowercase();
            out.insert(name_lc);
            // `Table` validates both names as ASCII-only identifiers; avoid allocating a second
            // lowercased copy when the display name matches.
            if !table.display_name.eq_ignore_ascii_case(&table.name) {
                out.insert(table.display_name.to_ascii_lowercase());
            }
        }
    }
    out
}

fn collect_pivot_table_names(pivots: &[PivotTableModel]) -> Result<Vec<String>, DuplicateSheetError> {
    let mut out: Vec<String> = Vec::new();
    if out.try_reserve_exact(pivots.len()).is_err() {
        debug_assert!(
            false,
            "allocation failed (collect pivot table names, count={})",
            pivots.len()
        );
        return Err(DuplicateSheetError::AllocationFailure(
            "collect pivot table names",
        ));
    }
    for pivot in pivots {
        out.push(pivot.name.clone());
    }
    Ok(out)
}

fn next_table_id(sheets: &[Worksheet]) -> u32 {
    sheets
        .iter()
        .flat_map(|s| s.tables.iter())
        .map(|t| t.id)
        .max()
        .unwrap_or(0)
        .wrapping_add(1)
}

fn generate_duplicate_table_name(base: &str, used_names: &mut HashSet<String>) -> String {
    // Excel renames duplicated tables by appending `_1`, `_2`, … to the existing name.
    //
    // Table names are validated as ASCII-only identifiers, so case-insensitive matching is
    // equivalent to ASCII-lowercasing.
    use std::fmt::Write as _;
    fn push_ascii_lowercase(out: &mut String, s: &str) {
        for &b in s.as_bytes() {
            out.push(b.to_ascii_lowercase() as char);
        }
    }
    let mut i: u64 = 1;
    loop {
        let mut key = String::new();
        if key.try_reserve_exact(base.len().saturating_add(1).saturating_add(20)).is_err() {
            // Best-effort: proceed with incremental growth.
            debug_assert!(
                false,
                "allocation failed (duplicate table name key buffer, base_len={})",
                base.len()
            );
        }
        push_ascii_lowercase(&mut key, base);
        key.push('_');
        let _ = write!(&mut key, "{i}");
        if used_names.insert(key) {
            return format!("{base}_{i}");
        }
        i = i.wrapping_add(1);
    }
}

fn generate_duplicate_pivot_table_name(
    base: &str,
    used_names: &mut Vec<String>,
) -> Result<String, DuplicateSheetError> {
    // Match Excel-style name collision behavior: `PivotTable1` -> `PivotTable1 (2)`.
    let mut i: u64 = 2;
    loop {
        let candidate = format!("{base} ({i})");
        if used_names
            .iter()
            .all(|name| !text_eq_case_insensitive(name, &candidate))
        {
            if used_names.try_reserve(1).is_err() {
                debug_assert!(
                    false,
                    "allocation failed (insert pivot table name, len={})",
                    used_names.len()
                );
                return Err(DuplicateSheetError::AllocationFailure(
                    "insert pivot table name",
                ));
            }
            used_names.push(candidate.clone());
            return Ok(candidate);
        }
        i = i.wrapping_add(1);
    }
}

fn pivot_destination_is_on_sheet(
    destination: &PivotDestination,
    sheet_id: WorksheetId,
    sheet_name: &str,
) -> bool {
    match destination {
        PivotDestination::Cell { sheet_id: id, .. }
        | PivotDestination::Range { sheet_id: id, .. } => *id == sheet_id,
        PivotDestination::CellName {
            sheet_name: name, ..
        }
        | PivotDestination::RangeName {
            sheet_name: name, ..
        } => crate::sheet_name::sheet_name_eq_case_insensitive(name, sheet_name),
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
            #[serde(default)]
            date_system: DateSystem,
            #[serde(
                default = "default_codepage",
                alias = "text_codepage",
                alias = "textCodepage"
            )]
            codepage: u16,
            #[serde(default)]
            theme: ThemePalette,
            #[serde(default)]
            workbook_protection: WorkbookProtection,
            #[serde(default)]
            defined_names: Vec<DefinedName>,
            #[serde(default, rename = "pivotTables", alias = "pivot_tables")]
            pivot_tables: Vec<PivotTableModel>,
            #[serde(default, rename = "pivotCaches", alias = "pivot_caches")]
            pivot_caches: Vec<PivotCacheModel>,
            #[serde(default, rename = "pivotCharts", alias = "pivot_charts")]
            pivot_charts: Vec<PivotChartModel>,
            #[serde(default)]
            slicers: Vec<SlicerModel>,
            #[serde(default)]
            timelines: Vec<TimelineModel>,
            #[serde(default)]
            print_settings: WorkbookPrintSettings,
            #[serde(default)]
            view: Option<WorkbookView>,
        }

        let helper = Helper::deserialize(deserializer)?;

        if helper.schema_version > crate::SCHEMA_VERSION {
            return Err(D::Error::custom(format!(
                "unsupported schema_version {} (max supported: {})",
                helper.schema_version,
                crate::SCHEMA_VERSION
            )));
        }

        let sheets = helper.sheets;
        let defined_names = helper.defined_names;
        let pivot_tables = helper.pivot_tables;
        let pivot_caches = helper.pivot_caches;
        let pivot_charts = helper.pivot_charts;
        let slicers = helper.slicers;
        let timelines = helper.timelines;

        let next_sheet_id = sheets
            .iter()
            .map(|s| s.id)
            .max()
            .unwrap_or(0)
            .wrapping_add(1);

        let next_defined_name_id = defined_names
            .iter()
            .map(|n| n.id)
            .max()
            .unwrap_or(0)
            .wrapping_add(1);

        let mut print_settings = helper.print_settings;
        for sheet_settings in &mut print_settings.sheets {
            if let Some(sheet) = sheets.iter().find(|s| {
                crate::sheet_name::sheet_name_eq_case_insensitive(
                    &s.name,
                    &sheet_settings.sheet_name,
                )
            }) {
                sheet_settings.sheet_name = sheet.name.clone();
            }
        }

        let mut view = helper.view.unwrap_or_default();
        if let Some(active) = view.active_sheet_id {
            if sheets.iter().all(|s| s.id != active) {
                view.active_sheet_id = None;
            }
        }

        let mut workbook = Workbook {
            schema_version: helper.schema_version,
            id: helper.id,
            sheets,
            styles: helper.styles,
            images: helper.images,
            calc_settings: helper.calc_settings,
            date_system: helper.date_system,
            codepage: helper.codepage,
            theme: helper.theme,
            workbook_protection: helper.workbook_protection,
            defined_names,
            pivot_tables,
            pivot_caches,
            pivot_charts,
            slicers,
            timelines,
            print_settings,
            view,
            next_sheet_id,
            next_defined_name_id,
        };

        // Ensure deterministic ordering for serialization and UX.
        workbook.sort_print_settings_by_sheet_order();

        Ok(workbook)
    }
}
