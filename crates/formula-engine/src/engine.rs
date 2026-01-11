use crate::eval::{
    compile_canonical_expr, parse_a1, CellAddr, CompiledExpr, Expr, FormulaParseError, Parser,
    RangeRef, SheetReference,
};
use crate::editing::{
    CellChange, CellSnapshot, EditError, EditOp, EditResult, FormulaRewrite, MovedRange,
};
use crate::editing::rewrite::{
    rewrite_formula_for_copy_delta, rewrite_formula_for_range_map, rewrite_formula_for_structural_edit,
    GridRange, RangeMapEdit, StructuralEdit,
};
use crate::graph::{CellDeps, DependencyGraph as CalcGraph, Precedent, SheetRange};
use crate::locale::{canonicalize_formula, canonicalize_formula_with_style, FormulaLocale};
use crate::value::{Array, ErrorKind, Value};
use crate::calc_settings::{CalcSettings, CalculationMode};
use crate::iterative;
use formula_model::{CellId, CellRef, Range, Table};
use rayon::prelude::*;
use std::cmp::max;
use std::collections::{BTreeSet, HashMap, HashSet, VecDeque};
use std::sync::Arc;
use thiserror::Error;

pub type SheetId = usize;

#[derive(Debug, Error)]
pub enum EngineError {
    #[error(transparent)]
    Address(#[from] crate::eval::AddressParseError),
    #[error(transparent)]
    Parse(#[from] FormulaParseError),
    #[error(transparent)]
    AstParse(#[from] crate::ParseError),
    #[error(transparent)]
    AstSerialize(#[from] crate::SerializeError),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RecalcMode {
    SingleThreaded,
    MultiThreaded,
}

/// Scope for a defined name / named range.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum NameScope<'a> {
    Workbook,
    Sheet(&'a str),
}

/// A defined name (named range) definition.
#[derive(Debug, Clone, PartialEq)]
pub enum NameDefinition {
    /// A constant scalar value (number/text/bool/error).
    Constant(Value),
    /// A reference/range definition (typically something like `Sheet1!$A$1:$B$3`).
    Reference(String),
    /// A formula definition stored as a canonical formula string (may evaluate to a scalar or reference).
    Formula(String),
}

#[derive(Debug, Clone)]
struct DefinedName {
    definition: NameDefinition,
    compiled: Option<CompiledExpr>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub(crate) struct CellKey {
    pub(crate) sheet: SheetId,
    pub(crate) addr: CellAddr,
}

#[derive(Debug, Clone)]
struct Cell {
    value: Value,
    formula: Option<String>,
    ast: Option<CompiledExpr>,
    volatile: bool,
    thread_safe: bool,
}

impl Default for Cell {
    fn default() -> Self {
        Self {
            value: Value::Blank,
            formula: None,
            ast: None,
            volatile: false,
            thread_safe: true,
        }
    }
}

#[derive(Debug, Default, Clone)]
struct Sheet {
    cells: HashMap<CellAddr, Cell>,
    tables: Vec<Table>,
    names: HashMap<String, DefinedName>,
}

#[derive(Debug, Default, Clone)]
struct Workbook {
    sheets: Vec<Sheet>,
    sheet_names: Vec<String>,
    sheet_name_to_id: HashMap<String, SheetId>,
    names: HashMap<String, DefinedName>,
}

impl Workbook {
    fn sheet_key(name: &str) -> String {
        // Excel treats sheet names as case-insensitive. Use a normalized lookup key while
        // preserving the original display name.
        name.to_ascii_uppercase()
    }

    fn ensure_sheet(&mut self, name: &str) -> SheetId {
        let key = Self::sheet_key(name);
        if let Some(id) = self.sheet_name_to_id.get(&key).copied() {
            return id;
        }
        let id = self.sheets.len();
        self.sheets.push(Sheet {
            cells: HashMap::new(),
            tables: Vec::new(),
            names: HashMap::new(),
        });
        self.sheet_names.push(name.to_string());
        self.sheet_name_to_id.insert(key, id);
        id
    }

    fn sheet_id(&self, name: &str) -> Option<SheetId> {
        let key = Self::sheet_key(name);
        self.sheet_name_to_id.get(&key).copied()
    }

    fn get_cell(&self, key: CellKey) -> Option<&Cell> {
        self.sheets.get(key.sheet)?.cells.get(&key.addr)
    }

    fn get_or_create_cell_mut(&mut self, key: CellKey) -> &mut Cell {
        self.sheets[key.sheet].cells.entry(key.addr).or_default()
    }

    fn get_cell_value(&self, key: CellKey) -> Value {
        self.get_cell(key)
            .map(|c| c.value.clone())
            .unwrap_or(Value::Blank)
    }

    fn set_tables(&mut self, sheet: SheetId, tables: Vec<Table>) {
        if let Some(s) = self.sheets.get_mut(sheet) {
            s.tables = tables;
        }
    }
}

/// Expanded dependency view used for UX/auditing.
///
/// This intentionally stores **cell-level** precedents (ranges are expanded),
/// matching Excel's precedent/dependent tracing UX and making it easy to explain
/// why a cell is dirty.
#[derive(Debug, Default)]
struct AuditGraph {
    precedents: HashMap<CellKey, HashSet<CellKey>>,
    dependents: HashMap<CellKey, HashSet<CellKey>>,
    volatile_cells: HashSet<CellKey>,
}

impl AuditGraph {
    fn set_precedents(&mut self, cell: CellKey, new_precedents: HashSet<CellKey>) {
        if let Some(old) = self.precedents.remove(&cell) {
            for p in old {
                if let Some(deps) = self.dependents.get_mut(&p) {
                    deps.remove(&cell);
                    if deps.is_empty() {
                        self.dependents.remove(&p);
                    }
                }
            }
        }

        for p in &new_precedents {
            self.dependents.entry(*p).or_default().insert(cell);
        }

        if !new_precedents.is_empty() {
            self.precedents.insert(cell, new_precedents);
        }
    }

    fn clear_cell(&mut self, cell: CellKey) {
        self.set_precedents(cell, HashSet::new());
        self.volatile_cells.remove(&cell);
    }
}

#[derive(Debug, Clone)]
struct Spill {
    end: CellAddr,
    array: Array,
}

#[derive(Debug, Clone)]
struct BlockedSpill {
    blocker: CellKey,
}

#[derive(Debug, Default, Clone)]
struct SpillState {
    by_origin: HashMap<CellKey, Spill>,
    origin_by_cell: HashMap<CellKey, CellKey>,
    /// Spill origins currently evaluating to `#SPILL!` due to a blocked spill range.
    blocked_by_origin: HashMap<CellKey, BlockedSpill>,
    /// Reverse index: blocker cell -> spill origins that should be re-evaluated when the blocker changes.
    blocked_origins_by_cell: HashMap<CellKey, HashSet<CellKey>>,
}

pub struct Engine {
    workbook: Workbook,
    external_value_provider: Option<Arc<dyn ExternalValueProvider>>,
    name_dependents: HashMap<String, HashSet<CellKey>>,
    cell_name_refs: HashMap<CellKey, HashSet<String>>,
    /// Optimized dependency graph used for incremental recalculation ordering.
    calc_graph: CalcGraph,
    /// Expanded dependency graph used for auditing/introspection (precedents/dependents queries).
    graph: AuditGraph,
    dirty: HashSet<CellKey>,
    dirty_reasons: HashMap<CellKey, CellKey>,
    calc_settings: CalcSettings,
    circular_references: HashSet<CellKey>,
    spills: SpillState,
}

impl Default for Engine {
    fn default() -> Self {
        Self::new()
    }
}

impl Engine {
    pub fn new() -> Self {
        Self {
            workbook: Workbook::default(),
            external_value_provider: None,
            name_dependents: HashMap::new(),
            cell_name_refs: HashMap::new(),
            calc_graph: CalcGraph::new(),
            graph: AuditGraph::default(),
            dirty: HashSet::new(),
            dirty_reasons: HashMap::new(),
            // Default to manual calculation to preserve historical engine behavior; callers can
            // opt into Excel-like automatic mode by setting `CalcSettings.calculation_mode`.
            calc_settings: CalcSettings {
                calculation_mode: CalculationMode::Manual,
                ..CalcSettings::default()
            },
            circular_references: HashSet::new(),
            spills: SpillState::default(),
        }
    }

    pub fn calc_settings(&self) -> &CalcSettings {
        &self.calc_settings
    }

    /// Ensure a sheet exists in the workbook.
    ///
    /// This is useful for workbook load flows where formulas may refer to other sheets
    /// that have not been populated yet; callers should create all sheets up-front
    /// before setting formulas to ensure cross-sheet references resolve correctly.
    pub fn ensure_sheet(&mut self, sheet: &str) {
        self.workbook.ensure_sheet(sheet);
    }

    pub fn set_calc_settings(&mut self, settings: CalcSettings) {
        self.calc_settings = settings;
    }

    pub fn set_external_value_provider(&mut self, provider: Option<Arc<dyn ExternalValueProvider>>) {
        self.external_value_provider = provider;
    }

    pub fn has_dirty_cells(&self) -> bool {
        !self.dirty.is_empty()
    }

    pub fn circular_reference_count(&self) -> usize {
        self.circular_references.len()
    }

    pub fn set_cell_value(
        &mut self,
        sheet: &str,
        addr: &str,
        value: impl Into<Value>,
    ) -> Result<(), EngineError> {
        let sheet_id = self.workbook.ensure_sheet(sheet);
        let addr = parse_a1(addr)?;
        let key = CellKey { sheet: sheet_id, addr };
        let cell_id = cell_id_from_key(key);

        self.clear_spill_for_cell(key);
        self.clear_blocked_spill_for_origin(key);

        // Replace any existing formula and dependencies.
        self.graph.clear_cell(key);
        self.calc_graph.remove_cell(cell_id);
        self.clear_cell_name_refs(key);
        self.dirty.remove(&key);
        self.dirty_reasons.remove(&key);

        let cell = self.workbook.get_or_create_cell_mut(key);
        cell.value = value.into();
        cell.formula = None;
        cell.ast = None;
        cell.volatile = false;
        cell.thread_safe = true;

        // Mark downstream dependents dirty.
        self.mark_dirty_dependents_with_reasons(key);
        self.calc_graph.mark_dirty(cell_id);
        self.mark_dirty_blocked_spill_origins_for_cell(key);
        self.sync_dirty_from_calc_graph();
        if self.calc_settings.calculation_mode != CalculationMode::Manual {
            self.recalculate();
        }
        Ok(())
    }

    /// Replace the set of tables for a given worksheet.
    ///
    /// Tables are needed to resolve structured references like `Table1[Col]` and `[@Col]`.
    pub fn set_sheet_tables(&mut self, sheet: &str, tables: Vec<Table>) {
        let sheet_id = self.workbook.ensure_sheet(sheet);
        self.workbook.set_tables(sheet_id, tables);

        let tables_by_sheet: Vec<Vec<Table>> =
            self.workbook.sheets.iter().map(|s| s.tables.clone()).collect();

        // Structured reference resolution can change which cells a formula depends on, so refresh
        // dependencies for all formulas.
        let mut formulas: Vec<(CellKey, CompiledExpr)> = Vec::new();
        for (sheet_id, sheet) in self.workbook.sheets.iter().enumerate() {
            for (addr, cell) in &sheet.cells {
                if let Some(ast) = cell.ast.clone() {
                    formulas.push((CellKey { sheet: sheet_id, addr: *addr }, ast));
                }
            }
        }

        for (key, ast) in formulas {
            let cell_id = cell_id_from_key(key);
            let (precedents, names, volatile, thread_safe) =
                analyze_expr(&ast, key, &tables_by_sheet, &self.workbook);
            self.graph.set_precedents(key, precedents);
            if volatile {
                self.graph.volatile_cells.insert(key);
            } else {
                self.graph.volatile_cells.remove(&key);
            }
            self.set_cell_name_refs(key, names);

            let calc_precedents =
                analyze_calc_precedents(&ast, key, &tables_by_sheet, &self.workbook);
            let mut calc_vec: Vec<Precedent> = calc_precedents.into_iter().collect();
            calc_vec.sort_by_key(|p| match p {
                Precedent::Cell(c) => (0u8, c.sheet_id, c.cell.row, c.cell.col),
                Precedent::Range(r) => (1u8, r.sheet_id, r.range.start.row, r.range.start.col),
            });
            let deps = CellDeps::new(calc_vec).volatile(volatile);
            self.calc_graph.update_cell_dependencies(cell_id, deps);

            let cell = self.workbook.get_or_create_cell_mut(key);
            cell.volatile = volatile;
            cell.thread_safe = thread_safe;

            self.dirty.insert(key);
            self.dirty_reasons.remove(&key);
            self.calc_graph.mark_dirty(cell_id);
        }
    }

    pub fn define_name(
        &mut self,
        name: &str,
        scope: NameScope<'_>,
        definition: NameDefinition,
    ) -> Result<(), EngineError> {
        let name_key = normalize_defined_name(name);
        if name_key.is_empty() {
            return Ok(());
        }

        let compiled = match &definition {
            NameDefinition::Constant(_) => None,
            NameDefinition::Reference(formula) | NameDefinition::Formula(formula) => {
                let parsed = Parser::parse(formula)?;
                Some(self.compile_name_expr(&parsed))
            }
        };

        let entry = DefinedName {
            definition,
            compiled,
        };

        match scope {
            NameScope::Workbook => {
                self.workbook.names.insert(name_key.clone(), entry);
            }
            NameScope::Sheet(sheet_name) => {
                let sheet_id = self.workbook.ensure_sheet(sheet_name);
                self.workbook.sheets[sheet_id]
                    .names
                    .insert(name_key.clone(), entry);
            }
        }

        self.refresh_cells_after_name_change(&name_key);
        if self.calc_settings.calculation_mode != CalculationMode::Manual {
            self.recalculate();
        }
        Ok(())
    }

    pub fn remove_name(&mut self, name: &str, scope: NameScope<'_>) -> Option<NameDefinition> {
        let name_key = normalize_defined_name(name);
        if name_key.is_empty() {
            return None;
        }

        let removed = match scope {
            NameScope::Workbook => self.workbook.names.remove(&name_key),
            NameScope::Sheet(sheet_name) => {
                let sheet_id = self.workbook.sheet_id(sheet_name)?;
                self.workbook.sheets.get_mut(sheet_id)?.names.remove(&name_key)
            }
        };

        let removed_def = removed.map(|n| n.definition);
        if removed_def.is_some() {
            self.refresh_cells_after_name_change(&name_key);
            if self.calc_settings.calculation_mode != CalculationMode::Manual {
                self.recalculate();
            }
        }
        removed_def
    }

    pub fn get_name(&self, name: &str, scope: NameScope<'_>) -> Option<&NameDefinition> {
        let name_key = normalize_defined_name(name);
        if name_key.is_empty() {
            return None;
        }

        match scope {
            NameScope::Workbook => self.workbook.names.get(&name_key).map(|n| &n.definition),
            NameScope::Sheet(sheet_name) => {
                let sheet_id = self.workbook.sheet_id(sheet_name)?;
                self.workbook.sheets.get(sheet_id)?.names.get(&name_key).map(|n| &n.definition)
            }
        }
    }

    pub fn set_cell_formula(
        &mut self,
        sheet: &str,
        addr: &str,
        formula: &str,
    ) -> Result<(), EngineError> {
        let sheet_id = self.workbook.ensure_sheet(sheet);
        let addr = parse_a1(addr)?;
        let key = CellKey { sheet: sheet_id, addr };
        let cell_id = cell_id_from_key(key);
        self.clear_spill_for_cell(key);
        self.clear_blocked_spill_for_origin(key);

        let parsed = crate::parse_formula(formula, crate::ParseOptions::default())?;
        let mut resolve_sheet = |name: &str| self.workbook.sheet_id(name);
        let compiled = compile_canonical_expr(&parsed.expr, sheet_id, addr, &mut resolve_sheet);
        let tables_by_sheet: Vec<Vec<Table>> =
            self.workbook.sheets.iter().map(|s| s.tables.clone()).collect();
 
        // Expanded precedents for auditing, plus volatility/thread-safety flags.
        let (precedents, names, volatile, thread_safe) =
            analyze_expr(&compiled, key, &tables_by_sheet, &self.workbook);
        self.graph.set_precedents(key, precedents);
        if volatile {
            self.graph.volatile_cells.insert(key);
        } else {
            self.graph.volatile_cells.remove(&key);
        }
        self.set_cell_name_refs(key, names);
 
        // Optimized precedents for calculation ordering (range nodes are not expanded).
        let calc_precedents =
            analyze_calc_precedents(&compiled, key, &tables_by_sheet, &self.workbook);
        let mut calc_vec: Vec<Precedent> = calc_precedents.into_iter().collect();
        calc_vec.sort_by_key(|p| match p {
            Precedent::Cell(c) => (0u8, c.sheet_id, c.cell.row, c.cell.col),
            Precedent::Range(r) => (1u8, r.sheet_id, r.range.start.row, r.range.start.col),
        });
        let deps = CellDeps::new(calc_vec).volatile(volatile);
        self.calc_graph.update_cell_dependencies(cell_id, deps);

        let cell = self.workbook.get_or_create_cell_mut(key);
        cell.formula = Some(formula.to_string());
        cell.ast = Some(compiled);
        cell.volatile = volatile;
        cell.thread_safe = thread_safe;

        // Recalculate this cell and anything depending on it.
        self.mark_dirty_including_self_with_reasons(key);
        self.calc_graph.mark_dirty(cell_id);
        self.mark_dirty_blocked_spill_origins_for_cell(key);
        self.sync_dirty_from_calc_graph();
        if self.calc_settings.calculation_mode != CalculationMode::Manual {
            self.recalculate();
        }
        Ok(())
    }

    /// Set a cell formula that was entered using a different reference style (e.g. R1C1).
    ///
    /// The engine persists formulas in canonical A1 form, so this parses the input formula with
    /// the provided [`crate::ParseOptions`] and then serializes it back to A1 based on the cell
    /// location (`addr`).
    pub fn set_cell_formula_with_options(
        &mut self,
        sheet: &str,
        addr: &str,
        formula: &str,
        mut opts: crate::ParseOptions,
    ) -> Result<(), EngineError> {
        let origin_eval = parse_a1(addr)?;
        let origin = crate::CellAddr::new(origin_eval.row, origin_eval.col);

        // Normalize any relative A1 coordinates against the destination cell so the AST is
        // origin-relative regardless of the input reference style.
        opts.normalize_relative_to = Some(origin);

        let ast = crate::parse_formula(formula, opts)?;

        let canonical = ast.to_string(crate::SerializeOptions {
            locale: crate::LocaleConfig::en_us(),
            reference_style: crate::ReferenceStyle::A1,
            // Preserve `_xlfn.` prefixes for round-trip safety when callers include them.
            include_xlfn_prefix: true,
            origin: Some(origin),
            omit_equals: false,
        })?;

        self.set_cell_formula(sheet, addr, &canonical)
    }

    /// Convenience wrapper around [`Engine::set_cell_formula_with_options`] for R1C1 formulas.
    pub fn set_cell_formula_r1c1(
        &mut self,
        sheet: &str,
        addr: &str,
        formula_r1c1: &str,
    ) -> Result<(), EngineError> {
        self.set_cell_formula_with_options(
            sheet,
            addr,
            formula_r1c1,
            crate::ParseOptions {
                locale: crate::LocaleConfig::en_us(),
                reference_style: crate::ReferenceStyle::R1C1,
                normalize_relative_to: None,
            },
        )
    }

    /// Set a cell formula that was entered in a locale-specific display format.
    ///
    /// This converts the incoming formula to canonical form before parsing and
    /// persistence. Canonical form uses English function names and `,`/`.` for
    /// separators, which matches XLSX expectations and keeps storage stable across
    /// UI locale changes.
    pub fn set_cell_formula_localized(
        &mut self,
        sheet: &str,
        addr: &str,
        localized_formula: &str,
        locale: &FormulaLocale,
    ) -> Result<(), EngineError> {
        let canonical = canonicalize_formula(localized_formula, locale)?;
        self.set_cell_formula(sheet, addr, &canonical)
    }

    /// Set a cell formula that was entered in a locale-specific display format using R1C1
    /// reference style.
    ///
    /// This converts the incoming formula to canonical form before translating R1C1 references into
    /// the persisted A1 representation.
    pub fn set_cell_formula_localized_r1c1(
        &mut self,
        sheet: &str,
        addr: &str,
        localized_formula_r1c1: &str,
        locale: &FormulaLocale,
    ) -> Result<(), EngineError> {
        let canonical = canonicalize_formula_with_style(
            localized_formula_r1c1,
            locale,
            crate::ReferenceStyle::R1C1,
        )?;
        self.set_cell_formula_r1c1(sheet, addr, &canonical)
    }

    pub fn get_cell_value(&self, sheet: &str, addr: &str) -> Value {
        let Some(sheet_id) = self.workbook.sheet_id(sheet) else {
            return Value::Blank;
        };
        let Ok(addr) = parse_a1(addr) else {
            return Value::Error(ErrorKind::Ref);
        };
        let key = CellKey { sheet: sheet_id, addr };
        if let Some(v) = self.spilled_cell_value(key) {
            return v;
        }
        self.workbook.get_cell_value(key)
    }

    /// Returns the spill range (origin inclusive) for a cell if it is an array-spill
    /// origin or belongs to a spilled range.
    pub fn spill_range(&self, sheet: &str, addr: &str) -> Option<(CellAddr, CellAddr)> {
        let sheet_id = self.workbook.sheet_id(sheet)?;
        let addr = parse_a1(addr).ok()?;
        let key = CellKey { sheet: sheet_id, addr };
        let origin = self.spill_origin_key(key)?;
        let spill = self.spills.by_origin.get(&origin)?;
        Some((origin.addr, spill.end))
    }

    /// Returns the spill origin for a cell if it is an array-spill origin or belongs
    /// to a spilled range.
    pub fn spill_origin(&self, sheet: &str, addr: &str) -> Option<(SheetId, CellAddr)> {
        let sheet_id = self.workbook.sheet_id(sheet)?;
        let addr = parse_a1(addr).ok()?;
        let key = CellKey { sheet: sheet_id, addr };
        let origin = self.spill_origin_key(key)?;
        Some((origin.sheet, origin.addr))
    }

    pub fn get_cell_formula(&self, sheet: &str, addr: &str) -> Option<&str> {
        let sheet_id = self.workbook.sheet_id(sheet)?;
        let addr = parse_a1(addr).ok()?;
        let key = CellKey { sheet: sheet_id, addr };
        self.workbook.get_cell(key)?.formula.as_deref()
    }

    /// Returns the formula for `addr` rendered in R1C1 reference style.
    ///
    /// The engine persists formulas as canonical A1 strings; this converts them on demand using the
    /// syntax-only parser/serializer.
    pub fn get_cell_formula_r1c1(&self, sheet: &str, addr: &str) -> Option<String> {
        let sheet_id = self.workbook.sheet_id(sheet)?;
        let addr = parse_a1(addr).ok()?;
        let key = CellKey { sheet: sheet_id, addr };
        let formula = self.workbook.get_cell(key)?.formula.as_deref()?;

        let origin = crate::CellAddr::new(addr.row, addr.col);
        let ast = crate::parse_formula(
            formula,
            crate::ParseOptions {
                locale: crate::LocaleConfig::en_us(),
                reference_style: crate::ReferenceStyle::A1,
                normalize_relative_to: Some(origin),
            },
        )
        .ok()?;

        ast.to_string(crate::SerializeOptions {
            locale: crate::LocaleConfig::en_us(),
            reference_style: crate::ReferenceStyle::R1C1,
            include_xlfn_prefix: true,
            origin: Some(origin),
            omit_equals: false,
        })
        .ok()
    }

    pub fn apply_operation(&mut self, op: EditOp) -> Result<EditResult, EditError> {
        let before = self.workbook.clone();
        let mut formula_rewrites = Vec::new();
        let mut moved_ranges = Vec::new();

        let sheet_names = sheet_names_by_id(&self.workbook);

        match op {
            EditOp::InsertRows { sheet, row, count } => {
                if count == 0 {
                    return Err(EditError::InvalidCount);
                }
                let sheet_id = self
                    .workbook
                    .sheet_id(&sheet)
                    .ok_or_else(|| EditError::SheetNotFound(sheet.clone()))?;
                shift_rows(&mut self.workbook.sheets[sheet_id], row, count, true);
                formula_rewrites.extend(rewrite_all_formulas_structural(
                    &mut self.workbook,
                    &sheet_names,
                    StructuralEdit::InsertRows { sheet, row, count },
                ));
            }
            EditOp::DeleteRows { sheet, row, count } => {
                if count == 0 {
                    return Err(EditError::InvalidCount);
                }
                let sheet_id = self
                    .workbook
                    .sheet_id(&sheet)
                    .ok_or_else(|| EditError::SheetNotFound(sheet.clone()))?;
                shift_rows(&mut self.workbook.sheets[sheet_id], row, count, false);
                formula_rewrites.extend(rewrite_all_formulas_structural(
                    &mut self.workbook,
                    &sheet_names,
                    StructuralEdit::DeleteRows { sheet, row, count },
                ));
            }
            EditOp::InsertCols { sheet, col, count } => {
                if count == 0 {
                    return Err(EditError::InvalidCount);
                }
                let sheet_id = self
                    .workbook
                    .sheet_id(&sheet)
                    .ok_or_else(|| EditError::SheetNotFound(sheet.clone()))?;
                shift_cols(&mut self.workbook.sheets[sheet_id], col, count, true);
                formula_rewrites.extend(rewrite_all_formulas_structural(
                    &mut self.workbook,
                    &sheet_names,
                    StructuralEdit::InsertCols { sheet, col, count },
                ));
            }
            EditOp::DeleteCols { sheet, col, count } => {
                if count == 0 {
                    return Err(EditError::InvalidCount);
                }
                let sheet_id = self
                    .workbook
                    .sheet_id(&sheet)
                    .ok_or_else(|| EditError::SheetNotFound(sheet.clone()))?;
                shift_cols(&mut self.workbook.sheets[sheet_id], col, count, false);
                formula_rewrites.extend(rewrite_all_formulas_structural(
                    &mut self.workbook,
                    &sheet_names,
                    StructuralEdit::DeleteCols { sheet, col, count },
                ));
            }
            EditOp::InsertCellsShiftRight { sheet, range } => {
                let width = range.width();
                if width == 0 {
                    return Err(EditError::InvalidRange);
                }
                let sheet_id = self
                    .workbook
                    .sheet_id(&sheet)
                    .ok_or_else(|| EditError::SheetNotFound(sheet.clone()))?;
                insert_cells_shift_right(&mut self.workbook.sheets[sheet_id], range, width);
                let edit = RangeMapEdit {
                    sheet,
                    moved_region: GridRange::new(
                        range.start.row,
                        range.start.col,
                        range.end.row,
                        u32::MAX,
                    ),
                    delta_row: 0,
                    delta_col: width as i32,
                    deleted_region: None,
                };
                formula_rewrites.extend(rewrite_all_formulas_range_map(
                    &mut self.workbook,
                    &sheet_names,
                    &edit,
                ));
            }
            EditOp::InsertCellsShiftDown { sheet, range } => {
                let height = range.height();
                if height == 0 {
                    return Err(EditError::InvalidRange);
                }
                let sheet_id = self
                    .workbook
                    .sheet_id(&sheet)
                    .ok_or_else(|| EditError::SheetNotFound(sheet.clone()))?;
                insert_cells_shift_down(&mut self.workbook.sheets[sheet_id], range, height);
                let edit = RangeMapEdit {
                    sheet,
                    moved_region: GridRange::new(
                        range.start.row,
                        range.start.col,
                        u32::MAX,
                        range.end.col,
                    ),
                    delta_row: height as i32,
                    delta_col: 0,
                    deleted_region: None,
                };
                formula_rewrites.extend(rewrite_all_formulas_range_map(
                    &mut self.workbook,
                    &sheet_names,
                    &edit,
                ));
            }
            EditOp::DeleteCellsShiftLeft { sheet, range } => {
                let width = range.width();
                if width == 0 {
                    return Err(EditError::InvalidRange);
                }
                let sheet_id = self
                    .workbook
                    .sheet_id(&sheet)
                    .ok_or_else(|| EditError::SheetNotFound(sheet.clone()))?;
                delete_cells_shift_left(&mut self.workbook.sheets[sheet_id], range, width);
                let start_col = range.end.col.saturating_add(1);
                let edit = RangeMapEdit {
                    sheet,
                    moved_region: GridRange::new(range.start.row, start_col, range.end.row, u32::MAX),
                    delta_row: 0,
                    delta_col: -(width as i32),
                    deleted_region: Some(GridRange::new(
                        range.start.row,
                        range.start.col,
                        range.end.row,
                        range.end.col,
                    )),
                };
                formula_rewrites.extend(rewrite_all_formulas_range_map(
                    &mut self.workbook,
                    &sheet_names,
                    &edit,
                ));
            }
            EditOp::DeleteCellsShiftUp { sheet, range } => {
                let height = range.height();
                if height == 0 {
                    return Err(EditError::InvalidRange);
                }
                let sheet_id = self
                    .workbook
                    .sheet_id(&sheet)
                    .ok_or_else(|| EditError::SheetNotFound(sheet.clone()))?;
                delete_cells_shift_up(&mut self.workbook.sheets[sheet_id], range, height);
                let start_row = range.end.row.saturating_add(1);
                let edit = RangeMapEdit {
                    sheet,
                    moved_region: GridRange::new(start_row, range.start.col, u32::MAX, range.end.col),
                    delta_row: -(height as i32),
                    delta_col: 0,
                    deleted_region: Some(GridRange::new(
                        range.start.row,
                        range.start.col,
                        range.end.row,
                        range.end.col,
                    )),
                };
                formula_rewrites.extend(rewrite_all_formulas_range_map(
                    &mut self.workbook,
                    &sheet_names,
                    &edit,
                ));
            }
            EditOp::MoveRange {
                sheet,
                src,
                dst_top_left,
            } => {
                let sheet_id = self
                    .workbook
                    .sheet_id(&sheet)
                    .ok_or_else(|| EditError::SheetNotFound(sheet.clone()))?;
                if src.width() == 0 || src.height() == 0 {
                    return Err(EditError::InvalidRange);
                }
                let dst = Range::new(
                    dst_top_left,
                    CellRef::new(
                        dst_top_left.row + src.height() - 1,
                        dst_top_left.col + src.width() - 1,
                    ),
                );
                if ranges_overlap(src, dst) {
                    return Err(EditError::OverlappingMove);
                }
                move_range(&mut self.workbook.sheets[sheet_id], src, dst_top_left);
                let edit = RangeMapEdit {
                    sheet: sheet.clone(),
                    moved_region: GridRange::new(src.start.row, src.start.col, src.end.row, src.end.col),
                    delta_row: dst.start.row as i32 - src.start.row as i32,
                    delta_col: dst.start.col as i32 - src.start.col as i32,
                    deleted_region: None,
                };
                formula_rewrites.extend(rewrite_all_formulas_range_map(
                    &mut self.workbook,
                    &sheet_names,
                    &edit,
                ));
                moved_ranges.push(MovedRange { sheet, from: src, to: dst });
            }
            EditOp::CopyRange {
                sheet,
                src,
                dst_top_left,
            } => {
                let sheet_id = self
                    .workbook
                    .sheet_id(&sheet)
                    .ok_or_else(|| EditError::SheetNotFound(sheet.clone()))?;
                if src.width() == 0 || src.height() == 0 {
                    return Err(EditError::InvalidRange);
                }
                copy_range(
                    &mut self.workbook.sheets[sheet_id],
                    &sheet,
                    src,
                    dst_top_left,
                    &mut formula_rewrites,
                );
            }
            EditOp::Fill { sheet, src, dst } => {
                let sheet_id = self
                    .workbook
                    .sheet_id(&sheet)
                    .ok_or_else(|| EditError::SheetNotFound(sheet.clone()))?;
                fill_range(
                    &mut self.workbook.sheets[sheet_id],
                    &sheet,
                    src,
                    dst,
                    &mut formula_rewrites,
                );
            }
        }

        self.rebuild_graph()
            .map_err(|e| EditError::Engine(e.to_string()))?;

        let sheet_names_after = sheet_names_by_id(&self.workbook);
        let changed_cells = diff_workbooks(&before, &self.workbook, &sheet_names_after);

        Ok(EditResult {
            changed_cells,
            moved_ranges,
            formula_rewrites,
        })
    }

    pub fn recalculate(&mut self) {
        self.recalculate_with_mode(RecalcMode::MultiThreaded);
    }

    pub fn recalculate_single_threaded(&mut self) {
        self.recalculate_with_mode(RecalcMode::SingleThreaded);
    }

    pub fn recalculate_multi_threaded(&mut self) {
        self.recalculate_with_mode(RecalcMode::MultiThreaded);
    }

    fn recalculate_with_mode(&mut self, mode: RecalcMode) {
        loop {
            let levels = match self.calc_graph.calc_levels_for_dirty() {
                Ok(levels) => levels,
                Err(_) => {
                    self.recalculate_with_cycles(mode);
                    return;
                }
            };

            if levels.is_empty() {
                return;
            }

            self.circular_references.clear();

            let mut snapshot = Snapshot::from_workbook(
                &self.workbook,
                &self.spills,
                self.external_value_provider.clone(),
            );
            let mut spill_dirty_roots: Vec<CellId> = Vec::new();

            for level in levels {
                let mut keys: Vec<CellKey> = level.into_iter().map(cell_key_from_id).collect();
                keys.sort_by_key(|k| (k.sheet, k.addr.row, k.addr.col));
                let has_barrier = keys.iter().any(|&k| {
                    self.workbook
                        .get_cell(k)
                        .map(|c| c.volatile || !c.thread_safe)
                        .unwrap_or(false)
                });

                let tasks: Vec<(CellKey, CompiledExpr)> = keys
                    .iter()
                    .filter_map(|&k| {
                        self.workbook
                            .get_cell(k)
                            .and_then(|c| c.ast.clone().map(|a| (k, a)))
                    })
                    .collect();

                let mut results: Vec<(CellKey, Value)> =
                    if mode == RecalcMode::MultiThreaded && !has_barrier {
                        tasks
                            .par_iter()
                            .map(|(k, expr)| {
                                let ctx = crate::eval::EvalContext {
                                    current_sheet: k.sheet,
                                    current_cell: k.addr,
                                };
                                let evaluator = crate::eval::Evaluator::new(&snapshot, ctx);
                                (*k, evaluator.eval_formula(expr))
                            })
                            .collect()
                    } else {
                        tasks
                            .iter()
                            .map(|(k, expr)| {
                                let ctx = crate::eval::EvalContext {
                                    current_sheet: k.sheet,
                                    current_cell: k.addr,
                                };
                                let evaluator = crate::eval::Evaluator::new(&snapshot, ctx);
                                (*k, evaluator.eval_formula(expr))
                            })
                            .collect()
                    };

                results.sort_by_key(|(k, _)| (k.sheet, k.addr.row, k.addr.col));

                for (k, v) in results {
                    self.apply_eval_result(k, v, &mut snapshot, &mut spill_dirty_roots);
                }
            }

            self.calc_graph.clear_dirty();
            self.dirty.clear();
            self.dirty_reasons.clear();

            if spill_dirty_roots.is_empty() {
                return;
            }

            // Spills can change which cells are considered inputs vs computed spill outputs.
            // When the spill footprint changes (new cells gain/lose values), mark the affected
            // coordinates dirty so any dependents recalculate with the updated spill state.
            for cell in spill_dirty_roots.drain(..) {
                self.calc_graph.mark_dirty(cell);
            }
        }
    }

    fn recalculate_with_cycles(&mut self, _mode: RecalcMode) {
        let mut impacted_ids: HashSet<CellId> = self.calc_graph.dirty_cells().into_iter().collect();
        impacted_ids.extend(self.calc_graph.volatile_cells());

        if impacted_ids.is_empty() {
            return;
        }

        self.circular_references.clear();

        let mut impacted: Vec<CellKey> = impacted_ids
            .into_iter()
            .map(cell_key_from_id)
            .collect();
        impacted.sort_by_key(|k| (k.sheet, k.addr.row, k.addr.col));

        let impacted_set: HashSet<CellKey> = impacted.iter().copied().collect();
        let mut edges: HashMap<CellKey, Vec<CellKey>> = HashMap::new();
        for &cell in &impacted {
            let Some(deps) = self.graph.dependents.get(&cell) else {
                continue;
            };
            let mut out: Vec<CellKey> = deps
                .iter()
                .copied()
                .filter(|d| impacted_set.contains(d))
                .collect();
            if out.is_empty() {
                continue;
            }
            out.sort_by_key(|k| (k.sheet, k.addr.row, k.addr.col));
            edges.insert(cell, out);
        }

        let sccs = iterative::strongly_connected_components(&impacted, &edges);
        let order = iterative::topo_sort_sccs(&sccs, &edges);

        let mut snapshot = Snapshot::from_workbook(
            &self.workbook,
            &self.spills,
            self.external_value_provider.clone(),
        );
        let mut spill_dirty_roots: Vec<CellId> = Vec::new();

        for scc_idx in order {
            let mut scc = sccs[scc_idx].clone();
            scc.sort_by_key(|k| (k.sheet, k.addr.row, k.addr.col));

            let is_cycle = match scc.as_slice() {
                [] => continue,
                [only] => edges
                    .get(only)
                    .map(|deps| deps.contains(only))
                    .unwrap_or(false),
                _ => true,
            };

            if !is_cycle {
                let k = scc[0];
                let Some(expr) = self
                    .workbook
                    .get_cell(k)
                    .and_then(|c| c.ast.clone())
                else {
                    continue;
                };
                let ctx = crate::eval::EvalContext {
                    current_sheet: k.sheet,
                    current_cell: k.addr,
                };
                let evaluator = crate::eval::Evaluator::new(&snapshot, ctx);
                let v = evaluator.eval_formula(&expr);
                self.apply_eval_result(k, v, &mut snapshot, &mut spill_dirty_roots);
                continue;
            }

            for &k in &scc {
                self.circular_references.insert(k);
            }

            if !self.calc_settings.iterative.enabled {
                for &k in &scc {
                    let v = Value::Number(0.0);
                    self.apply_eval_result(k, v, &mut snapshot, &mut spill_dirty_roots);
                }
                continue;
            }

            let max_iters = max(1, self.calc_settings.iterative.max_iterations) as usize;
            let tol = self.calc_settings.iterative.max_change.max(0.0);

            for _ in 0..max_iters {
                let mut max_delta: f64 = 0.0;
                for &k in &scc {
                    let Some(expr) = self
                        .workbook
                        .get_cell(k)
                        .and_then(|c| c.ast.clone())
                    else {
                        continue;
                    };
                    let old = snapshot.values.get(&k).cloned().unwrap_or(Value::Blank);
                    let ctx = crate::eval::EvalContext {
                        current_sheet: k.sheet,
                        current_cell: k.addr,
                    };
                    let evaluator = crate::eval::Evaluator::new(&snapshot, ctx);
                    let new_val = evaluator.eval_formula(&expr);
                    max_delta = max_delta.max(value_delta(&old, &new_val));
                    self.apply_eval_result(k, new_val, &mut snapshot, &mut spill_dirty_roots);
                }

                if max_delta <= tol {
                    break;
                }
            }
        }

        self.calc_graph.clear_dirty();
        self.dirty.clear();
        self.dirty_reasons.clear();
    }

    fn apply_eval_result(
        &mut self,
        key: CellKey,
        value: Value,
        snapshot: &mut Snapshot,
        spill_dirty_roots: &mut Vec<CellId>,
    ) {
        // Clear any previously tracked spill blockage for this origin before applying the new
        // evaluation result.
        self.clear_blocked_spill_for_origin(key);

        match value {
            Value::Array(array) => {
                if array.rows == 0 || array.cols == 0 {
                    self.apply_eval_result(key, Value::Error(ErrorKind::Calc), snapshot, spill_dirty_roots);
                    return;
                }

                let mut spill_too_big = || {
                    let cleared = self.clear_spill_for_origin(key);
                    for cleared_key in cleared {
                        snapshot.values.remove(&cleared_key);
                        spill_dirty_roots.push(cell_id_from_key(cleared_key));
                        self.append_blocked_spill_dirty_roots(cleared_key, spill_dirty_roots);
                    }

                    let cell = self.workbook.get_or_create_cell_mut(key);
                    cell.value = Value::Error(ErrorKind::Spill);
                    snapshot.values.insert(key, Value::Error(ErrorKind::Spill));
                };

                let row_delta = match u32::try_from(array.rows.saturating_sub(1)) {
                    Ok(v) => v,
                    Err(_) => {
                        spill_too_big();
                        return;
                    }
                };
                let col_delta = match u32::try_from(array.cols.saturating_sub(1)) {
                    Ok(v) => v,
                    Err(_) => {
                        spill_too_big();
                        return;
                    }
                };

                let Some(end_row) = key.addr.row.checked_add(row_delta) else {
                    spill_too_big();
                    return;
                };
                let Some(end_col) = key.addr.col.checked_add(col_delta) else {
                    spill_too_big();
                    return;
                };

                let end = CellAddr {
                    row: end_row,
                    col: end_col,
                };

                // Fast path: if the spill shape is unchanged, update the stored array and
                // overwrite spill cell values in the snapshot without reshaping dependency nodes.
                if let Some(existing) = self.spills.by_origin.get_mut(&key) {
                    if existing.end == end {
                        existing.array = array.clone();

                        let top_left = array.top_left();
                        let cell = self.workbook.get_or_create_cell_mut(key);
                        cell.value = top_left.clone();
                        snapshot.values.insert(key, top_left);

                        for r in 0..array.rows {
                            for c in 0..array.cols {
                                if r == 0 && c == 0 {
                                    continue;
                                }
                                let addr = CellAddr {
                                    row: key.addr.row + r as u32,
                                    col: key.addr.col + c as u32,
                                };
                                let spill_key = CellKey { sheet: key.sheet, addr };
                                if let Some(v) = array.get(r, c).cloned() {
                                    snapshot.values.insert(spill_key, v);
                                }
                            }
                        }
                        return;
                    }
                }

                // Spill footprint change: clear the previous spill range (if any) before attempting
                // to write the new one.
                let cleared = self.clear_spill_for_origin(key);
                for cleared_key in cleared {
                    snapshot.values.remove(&cleared_key);
                    spill_dirty_roots.push(cell_id_from_key(cleared_key));
                    self.append_blocked_spill_dirty_roots(cleared_key, spill_dirty_roots);
                }

                if let Some(blocker) = self.spill_blocker(key, &array) {
                    self.record_blocked_spill(key, blocker);
                    let cell = self.workbook.get_or_create_cell_mut(key);
                    cell.value = Value::Error(ErrorKind::Spill);
                    snapshot.values.insert(key, Value::Error(ErrorKind::Spill));
                    return;
                }

                self.apply_new_spill(key, end, array, snapshot, spill_dirty_roots);
            }
            other => {
                let cleared = self.clear_spill_for_origin(key);
                for cleared_key in cleared {
                    snapshot.values.remove(&cleared_key);
                    spill_dirty_roots.push(cell_id_from_key(cleared_key));
                    self.append_blocked_spill_dirty_roots(cleared_key, spill_dirty_roots);
                }

                let cell = self.workbook.get_or_create_cell_mut(key);
                cell.value = other.clone();
                snapshot.values.insert(key, other);
            }
        }
    }

    fn spill_origin_key(&self, key: CellKey) -> Option<CellKey> {
        if self.spills.by_origin.contains_key(&key) {
            return Some(key);
        }
        self.spills.origin_by_cell.get(&key).copied()
    }

    fn spilled_cell_value(&self, key: CellKey) -> Option<Value> {
        let origin = self.spills.origin_by_cell.get(&key).copied()?;
        let spill = self.spills.by_origin.get(&origin)?;
        let row_off = key.addr.row.checked_sub(origin.addr.row)? as usize;
        let col_off = key.addr.col.checked_sub(origin.addr.col)? as usize;
        spill.array.get(row_off, col_off).cloned()
    }

    fn clear_blocked_spill_for_origin(&mut self, origin: CellKey) {
        let Some(blocked) = self.spills.blocked_by_origin.remove(&origin) else {
            return;
        };

        if let Some(origins) = self.spills.blocked_origins_by_cell.get_mut(&blocked.blocker) {
            origins.remove(&origin);
            if origins.is_empty() {
                self.spills.blocked_origins_by_cell.remove(&blocked.blocker);
            }
        }
    }

    fn record_blocked_spill(&mut self, origin: CellKey, blocker: CellKey) {
        self.spills
            .blocked_by_origin
            .insert(origin, BlockedSpill { blocker });
        self.spills
            .blocked_origins_by_cell
            .entry(blocker)
            .or_default()
            .insert(origin);
    }

    fn mark_dirty_blocked_spill_origins_for_cell(&mut self, cell: CellKey) {
        let Some(origins) = self.spills.blocked_origins_by_cell.get(&cell) else {
            return;
        };

        // Clone so we can freely mutate dirty bookkeeping while iterating.
        let origins: Vec<CellKey> = origins.iter().copied().collect();
        for origin in origins {
            let origin_id = cell_id_from_key(origin);
            self.calc_graph.mark_dirty(origin_id);

            if self.dirty.insert(origin) {
                self.dirty_reasons.insert(origin, cell);
            }
            self.mark_dirty_dependents_with_reasons(origin);
        }
    }

    fn append_blocked_spill_dirty_roots(&self, cell: CellKey, out: &mut Vec<CellId>) {
        let Some(origins) = self.spills.blocked_origins_by_cell.get(&cell) else {
            return;
        };
        for &origin in origins {
            out.push(cell_id_from_key(origin));
        }
    }

    fn clear_spill_for_cell(&mut self, key: CellKey) {
        let origin = match self.spill_origin_key(key) {
            Some(origin) => origin,
            None => return,
        };

        let cleared = self.clear_spill_for_origin(origin);
        for cleared_key in cleared {
            self.calc_graph.mark_dirty(cell_id_from_key(cleared_key));
            self.mark_dirty_blocked_spill_origins_for_cell(cleared_key);
        }

        // If a user edits any cell in a spill range, the origin needs to be re-evaluated
        // to either re-spill (if the blockage was removed) or surface #SPILL! (if blocked).
        if origin != key {
            self.calc_graph.mark_dirty(cell_id_from_key(origin));
        }
    }

    fn clear_spill_for_origin(&mut self, origin: CellKey) -> Vec<CellKey> {
        let Some(spill) = self.spills.by_origin.remove(&origin) else {
            return Vec::new();
        };

        let mut cleared = Vec::new();
        for r in 0..spill.array.rows {
            for c in 0..spill.array.cols {
                if r == 0 && c == 0 {
                    continue;
                }
                let addr = CellAddr {
                    row: origin.addr.row + r as u32,
                    col: origin.addr.col + c as u32,
                };
                let key = CellKey {
                    sheet: origin.sheet,
                    addr,
                };
                self.spills.origin_by_cell.remove(&key);
                self.calc_graph.remove_cell(cell_id_from_key(key));
                cleared.push(key);
            }
        }
        cleared
    }

    fn spill_blocker(&self, origin: CellKey, array: &Array) -> Option<CellKey> {
        for r in 0..array.rows {
            for c in 0..array.cols {
                if r == 0 && c == 0 {
                    continue;
                }
                let addr = CellAddr {
                    row: origin.addr.row + r as u32,
                    col: origin.addr.col + c as u32,
                };
                let key = CellKey {
                    sheet: origin.sheet,
                    addr,
                };

                // Blocked by non-empty user cell (literal or formula).
                if let Some(cell) = self.workbook.get_cell(key) {
                    if cell.formula.is_some() {
                        return Some(key);
                    }
                    if cell.value != Value::Blank {
                        return Some(key);
                    }
                }

                // Blocked by another spill.
                if let Some(other_origin) = self.spills.origin_by_cell.get(&key) {
                    if *other_origin != origin {
                        return Some(key);
                    }
                }
            }
        }
        None
    }

    fn apply_new_spill(
        &mut self,
        origin: CellKey,
        end: CellAddr,
        array: Array,
        snapshot: &mut Snapshot,
        spill_dirty_roots: &mut Vec<CellId>,
    ) {
        let top_left = array.top_left();

        let cell = self.workbook.get_or_create_cell_mut(origin);
        cell.value = top_left.clone();
        snapshot.values.insert(origin, top_left);

        self.spills.by_origin.insert(origin, Spill { end, array: array.clone() });

        let origin_id = cell_id_from_key(origin);
        for r in 0..array.rows {
            for c in 0..array.cols {
                if r == 0 && c == 0 {
                    continue;
                }
                let addr = CellAddr {
                    row: origin.addr.row + r as u32,
                    col: origin.addr.col + c as u32,
                };
                let key = CellKey {
                    sheet: origin.sheet,
                    addr,
                };
                self.spills.origin_by_cell.insert(key, origin);

                if let Some(v) = array.get(r, c).cloned() {
                    snapshot.values.insert(key, v);
                }

                // Register spill cells as formula nodes that depend on the origin so they participate in
                // calculation ordering and dirty marking.
                let deps = CellDeps::new(vec![Precedent::Cell(origin_id)]);
                self.calc_graph.update_cell_dependencies(cell_id_from_key(key), deps);

                spill_dirty_roots.push(cell_id_from_key(key));
            }
        }
    }

    fn compile_name_expr(&mut self, expr: &Expr<String>) -> CompiledExpr {
        let mut map = |sref: &SheetReference<String>| match sref {
            SheetReference::Current => SheetReference::Current,
            SheetReference::Sheet(name) => self
                .workbook
                .sheet_id(name)
                .map(SheetReference::Sheet)
                .unwrap_or_else(|| SheetReference::External(name.clone())),
            SheetReference::External(wb) => SheetReference::External(wb.clone()),
        };
        expr.map_sheets(&mut map)
    }

    fn clear_cell_name_refs(&mut self, cell: CellKey) {
        let Some(names) = self.cell_name_refs.remove(&cell) else {
            return;
        };
        for name in names {
            let should_remove = self
                .name_dependents
                .get_mut(&name)
                .map(|deps| {
                    deps.remove(&cell);
                    deps.is_empty()
                })
                .unwrap_or(false);
            if should_remove {
                self.name_dependents.remove(&name);
            }
        }
    }

    fn set_cell_name_refs(&mut self, cell: CellKey, names: HashSet<String>) {
        self.clear_cell_name_refs(cell);
        if names.is_empty() {
            return;
        }
        self.cell_name_refs.insert(cell, names.clone());
        for name in names {
            self.name_dependents.entry(name).or_default().insert(cell);
        }
    }

    fn refresh_cells_after_name_change(&mut self, name: &str) {
        let Some(cells) = self.name_dependents.get(name).cloned() else {
            return;
        };

        let tables_by_sheet: Vec<Vec<Table>> =
            self.workbook.sheets.iter().map(|s| s.tables.clone()).collect();

        for key in cells {
            let Some(ast) = self
                .workbook
                .get_cell(key)
                .and_then(|c| c.ast.clone())
            else {
                continue;
            };

            let cell_id = cell_id_from_key(key);

            let (precedents, names, volatile, thread_safe) =
                analyze_expr(&ast, key, &tables_by_sheet, &self.workbook);
            self.graph.set_precedents(key, precedents);
            if volatile {
                self.graph.volatile_cells.insert(key);
            } else {
                self.graph.volatile_cells.remove(&key);
            }
            self.set_cell_name_refs(key, names);

            let calc_precedents =
                analyze_calc_precedents(&ast, key, &tables_by_sheet, &self.workbook);
            let mut calc_vec: Vec<Precedent> = calc_precedents.into_iter().collect();
            calc_vec.sort_by_key(|p| match p {
                Precedent::Cell(c) => (0u8, c.sheet_id, c.cell.row, c.cell.col),
                Precedent::Range(r) => (1u8, r.sheet_id, r.range.start.row, r.range.start.col),
            });
            let deps = CellDeps::new(calc_vec).volatile(volatile);
            self.calc_graph.update_cell_dependencies(cell_id, deps);

            let cell = self.workbook.get_or_create_cell_mut(key);
            cell.volatile = volatile;
            cell.thread_safe = thread_safe;

            self.mark_dirty_including_self_with_reasons(key);
            self.calc_graph.mark_dirty(cell_id);
        }
    }
    fn rebuild_graph(&mut self) -> Result<(), EngineError> {
        let sheet_names = sheet_names_by_id(&self.workbook);
        let mut formulas: Vec<(String, CellAddr, String)> = Vec::new();
        for (sheet_id, sheet) in self.workbook.sheets.iter().enumerate() {
            let Some(sheet_name) = sheet_names.get(&sheet_id).cloned() else {
                continue;
            };
            for (addr, cell) in &sheet.cells {
                if let Some(formula) = &cell.formula {
                    formulas.push((sheet_name.clone(), *addr, formula.clone()));
                }
            }
        }

        self.graph = AuditGraph::default();
        self.calc_graph = CalcGraph::new();
        self.name_dependents.clear();
        self.cell_name_refs.clear();
        self.dirty.clear();
        self.dirty_reasons.clear();
        self.spills = SpillState::default();

        for (sheet_name, addr, formula) in formulas {
            let addr_a1 = cell_addr_to_a1(addr);
            self.set_cell_formula(&sheet_name, &addr_a1, &formula)?;
        }
        Ok(())
    }

    /// Returns whether a cell is currently marked dirty (needs recalculation).
    pub fn is_dirty(&self, sheet: &str, addr: &str) -> bool {
        let Some(sheet_id) = self.workbook.sheet_id(sheet) else {
            return false;
        };
        let Ok(addr) = parse_a1(addr) else {
            return false;
        };
        self.dirty.contains(&CellKey { sheet: sheet_id, addr })
    }

    /// Direct precedents (cells referenced by the formula in `cell`).
    pub fn precedents(
        &self,
        sheet: &str,
        addr: &str,
    ) -> Result<Vec<(SheetId, CellAddr)>, EngineError> {
        self.precedents_impl(sheet, addr, false)
    }

    /// Transitive precedents (all cells that can influence `cell`).
    pub fn precedents_transitive(
        &self,
        sheet: &str,
        addr: &str,
    ) -> Result<Vec<(SheetId, CellAddr)>, EngineError> {
        self.precedents_impl(sheet, addr, true)
    }

    /// Direct dependents (cells whose formulas reference `cell`).
    pub fn dependents(
        &self,
        sheet: &str,
        addr: &str,
    ) -> Result<Vec<(SheetId, CellAddr)>, EngineError> {
        self.dependents_impl(sheet, addr, false)
    }

    /// Transitive dependents (all downstream cells that are affected by `cell`).
    pub fn dependents_transitive(
        &self,
        sheet: &str,
        addr: &str,
    ) -> Result<Vec<(SheetId, CellAddr)>, EngineError> {
        self.dependents_impl(sheet, addr, true)
    }

    /// Returns a dependency path explaining why `cell` is currently dirty.
    ///
    /// The returned vector is ordered from the root cause (usually an edited
    /// input cell) to the provided `cell`.
    pub fn dirty_dependency_path(
        &self,
        sheet: &str,
        addr: &str,
    ) -> Option<Vec<(SheetId, CellAddr)>> {
        let sheet_id = self.workbook.sheet_id(sheet)?;
        let addr = parse_a1(addr).ok()?;
        let key = CellKey { sheet: sheet_id, addr };
        if !self.dirty.contains(&key) {
            return None;
        }

        let mut path = vec![key];
        let mut current = key;
        let mut guard = 0usize;
        while let Some(prev) = self.dirty_reasons.get(&current).copied() {
            path.push(prev);
            current = prev;
            guard += 1;
            if guard > 10_000 {
                break;
            }
        }
        path.reverse();
        Some(path.into_iter().map(|k| (k.sheet, k.addr)).collect())
    }

    /// Deterministically evaluates a cell's formula while capturing a per-node trace.
    ///
    /// This is intended for on-demand debugging and does **not** mutate engine state.
    pub fn debug_evaluate(
        &self,
        sheet: &str,
        addr: &str,
    ) -> Result<crate::debug::DebugEvaluation, EngineError> {
        let Some(sheet_id) = self.workbook.sheet_id(sheet) else {
            return Err(EngineError::Parse(FormulaParseError::UnexpectedToken(format!(
                "unknown sheet '{sheet}'"
            ))));
        };
        let addr = parse_a1(addr)?;
        let key = CellKey { sheet: sheet_id, addr };
        let cell = self.workbook.get_cell(key);
        let Some(formula) = cell.and_then(|c| c.formula.as_deref()) else {
            return Err(EngineError::Parse(FormulaParseError::UnexpectedToken(
                "cell has no formula".to_string(),
            )));
        };

        let snapshot = Snapshot::from_workbook(
            &self.workbook,
            &self.spills,
            self.external_value_provider.clone(),
        );
        let ctx = crate::eval::EvalContext {
            current_sheet: sheet_id,
            current_cell: addr,
        };

        // Parse with spans, compile sheet references without mutating the workbook,
        // then evaluate with tracing.
        let parsed = crate::debug::parse_spanned_formula(formula)?;
        let mut map = |sref: &SheetReference<String>| match sref {
            SheetReference::Current => SheetReference::Current,
            SheetReference::Sheet(name) => self
                .workbook
                .sheet_id(name)
                .map(SheetReference::Sheet)
                .unwrap_or_else(|| SheetReference::External(name.clone())),
            SheetReference::External(wb) => SheetReference::External(wb.clone()),
        };
        let compiled = parsed.map_sheets(&mut map);

        let (value, trace) = crate::debug::evaluate_with_trace(&snapshot, ctx, &compiled);

        Ok(crate::debug::DebugEvaluation {
            formula: formula.to_string(),
            value,
            trace,
        })
    }

    fn precedents_impl(
        &self,
        sheet: &str,
        addr: &str,
        transitive: bool,
    ) -> Result<Vec<(SheetId, CellAddr)>, EngineError> {
        let Some(sheet_id) = self.workbook.sheet_id(sheet) else {
            return Ok(Vec::new());
        };
        let addr = parse_a1(addr)?;
        let key = CellKey { sheet: sheet_id, addr };
        let nodes = if transitive {
            collect_transitive(&self.graph.precedents, key)
        } else {
            self.graph
                .precedents
                .get(&key)
                .map(|s| sorted_cell_keys(s))
                .unwrap_or_default()
        };
        Ok(nodes.into_iter().map(|k| (k.sheet, k.addr)).collect())
    }

    fn dependents_impl(
        &self,
        sheet: &str,
        addr: &str,
        transitive: bool,
    ) -> Result<Vec<(SheetId, CellAddr)>, EngineError> {
        let Some(sheet_id) = self.workbook.sheet_id(sheet) else {
            return Ok(Vec::new());
        };
        let addr = parse_a1(addr)?;
        let key = CellKey { sheet: sheet_id, addr };
        let nodes = if transitive {
            collect_transitive(&self.graph.dependents, key)
        } else {
            self.graph
                .dependents
                .get(&key)
                .map(|s| sorted_cell_keys(s))
                .unwrap_or_default()
        };
        Ok(nodes.into_iter().map(|k| (k.sheet, k.addr)).collect())
    }

    fn mark_dirty_including_self_with_reasons(&mut self, from: CellKey) {
        self.dirty.insert(from);
        self.dirty_reasons.remove(&from);
        self.mark_dirty_dependents_with_reasons(from);
    }

    fn mark_dirty_dependents_with_reasons(&mut self, from: CellKey) {
        let mut queue: VecDeque<(CellKey, CellKey)> = VecDeque::new();
        if let Some(deps) = self.graph.dependents.get(&from) {
            for dep in sorted_cell_keys(deps) {
                queue.push_back((from, dep));
            }
        }

        while let Some((cause, cell)) = queue.pop_front() {
            if !self.dirty.insert(cell) {
                continue;
            }
            self.dirty_reasons.entry(cell).or_insert(cause);
            if let Some(deps) = self.graph.dependents.get(&cell) {
                for dep in sorted_cell_keys(deps) {
                    queue.push_back((cell, dep));
                }
            }
        }
    }

    fn sync_dirty_from_calc_graph(&mut self) {
        for id in self.calc_graph.dirty_cells() {
            self.dirty.insert(cell_key_from_id(id));
        }
    }
}

fn sheet_names_by_id(workbook: &Workbook) -> HashMap<SheetId, String> {
    workbook
        .sheet_names
        .iter()
        .cloned()
        .enumerate()
        .map(|(id, name)| (id, name))
        .collect()
}

fn cell_ref_from_addr(addr: CellAddr) -> CellRef {
    CellRef::new(addr.row, addr.col)
}

fn cell_addr_from_cell_ref(cell: CellRef) -> CellAddr {
    CellAddr { row: cell.row, col: cell.col }
}

fn cell_addr_to_a1(addr: CellAddr) -> String {
    format!("{}{}", col_to_name(addr.col), addr.row + 1)
}

fn col_to_name(col: u32) -> String {
    let mut n = col + 1;
    let mut out = Vec::<u8>::new();
    while n > 0 {
        let rem = (n - 1) % 26;
        out.push(b'A' + rem as u8);
        n = (n - 1) / 26;
    }
    out.reverse();
    String::from_utf8(out).expect("column letters are ASCII")
}

fn ranges_overlap(a: Range, b: Range) -> bool {
    !(a.end.row < b.start.row
        || a.start.row > b.end.row
        || a.end.col < b.start.col
        || a.start.col > b.end.col)
}

fn shift_rows(sheet: &mut Sheet, row: u32, count: u32, insert: bool) {
    let del_end = row.saturating_add(count.saturating_sub(1));
    let mut new_cells = HashMap::with_capacity(sheet.cells.len());
    for (addr, cell) in std::mem::take(&mut sheet.cells) {
        if insert {
            if addr.row >= row {
                new_cells.insert(
                    CellAddr {
                        row: addr.row + count,
                        col: addr.col,
                    },
                    cell,
                );
            } else {
                new_cells.insert(addr, cell);
            }
            continue;
        }

        if addr.row < row {
            new_cells.insert(addr, cell);
        } else if addr.row > del_end {
            new_cells.insert(
                CellAddr {
                    row: addr.row - count,
                    col: addr.col,
                },
                cell,
            );
        }
    }
    sheet.cells = new_cells;
}

fn shift_cols(sheet: &mut Sheet, col: u32, count: u32, insert: bool) {
    let del_end = col.saturating_add(count.saturating_sub(1));
    let mut new_cells = HashMap::with_capacity(sheet.cells.len());
    for (addr, cell) in std::mem::take(&mut sheet.cells) {
        if insert {
            if addr.col >= col {
                new_cells.insert(
                    CellAddr {
                        row: addr.row,
                        col: addr.col + count,
                    },
                    cell,
                );
            } else {
                new_cells.insert(addr, cell);
            }
            continue;
        }

        if addr.col < col {
            new_cells.insert(addr, cell);
        } else if addr.col > del_end {
            new_cells.insert(
                CellAddr {
                    row: addr.row,
                    col: addr.col - count,
                },
                cell,
            );
        }
    }
    sheet.cells = new_cells;
}

fn insert_cells_shift_right(sheet: &mut Sheet, range: Range, width: u32) {
    let mut new_cells = HashMap::with_capacity(sheet.cells.len());
    for (addr, cell) in std::mem::take(&mut sheet.cells) {
        if addr.row >= range.start.row && addr.row <= range.end.row && addr.col >= range.start.col {
            new_cells.insert(
                CellAddr {
                    row: addr.row,
                    col: addr.col + width,
                },
                cell,
            );
        } else {
            new_cells.insert(addr, cell);
        }
    }
    sheet.cells = new_cells;
}

fn insert_cells_shift_down(sheet: &mut Sheet, range: Range, height: u32) {
    let mut new_cells = HashMap::with_capacity(sheet.cells.len());
    for (addr, cell) in std::mem::take(&mut sheet.cells) {
        if addr.col >= range.start.col && addr.col <= range.end.col && addr.row >= range.start.row {
            new_cells.insert(
                CellAddr {
                    row: addr.row + height,
                    col: addr.col,
                },
                cell,
            );
        } else {
            new_cells.insert(addr, cell);
        }
    }
    sheet.cells = new_cells;
}

fn delete_cells_shift_left(sheet: &mut Sheet, range: Range, width: u32) {
    let mut new_cells = HashMap::with_capacity(sheet.cells.len());
    for (addr, cell) in std::mem::take(&mut sheet.cells) {
        if addr.row >= range.start.row && addr.row <= range.end.row {
            if addr.col >= range.start.col && addr.col <= range.end.col {
                continue;
            }
            if addr.col > range.end.col {
                new_cells.insert(
                    CellAddr {
                        row: addr.row,
                        col: addr.col - width,
                    },
                    cell,
                );
            } else {
                new_cells.insert(addr, cell);
            }
        } else {
            new_cells.insert(addr, cell);
        }
    }
    sheet.cells = new_cells;
}

fn delete_cells_shift_up(sheet: &mut Sheet, range: Range, height: u32) {
    let mut new_cells = HashMap::with_capacity(sheet.cells.len());
    for (addr, cell) in std::mem::take(&mut sheet.cells) {
        if addr.col >= range.start.col && addr.col <= range.end.col {
            if addr.row >= range.start.row && addr.row <= range.end.row {
                continue;
            }
            if addr.row > range.end.row {
                new_cells.insert(
                    CellAddr {
                        row: addr.row - height,
                        col: addr.col,
                    },
                    cell,
                );
            } else {
                new_cells.insert(addr, cell);
            }
        } else {
            new_cells.insert(addr, cell);
        }
    }
    sheet.cells = new_cells;
}

fn move_range(sheet: &mut Sheet, src: Range, dst_top_left: CellRef) {
    let dst_top_left_addr = cell_addr_from_cell_ref(dst_top_left);
    let dst = Range::new(
        dst_top_left,
        CellRef::new(
            dst_top_left.row + src.height() - 1,
            dst_top_left.col + src.width() - 1,
        ),
    );

    let mut extracted: Vec<(CellRef, Option<Cell>)> = Vec::new();
    for cell in src.iter() {
        extracted.push((cell, sheet.cells.remove(&cell_addr_from_cell_ref(cell))));
    }

    for cell in dst.iter() {
        sheet.cells.remove(&cell_addr_from_cell_ref(cell));
    }

    for (cell, value) in extracted {
        let Some(value) = value else { continue };
        let dr = cell.row - src.start.row;
        let dc = cell.col - src.start.col;
        sheet.cells.insert(
            CellAddr {
                row: dst_top_left_addr.row + dr,
                col: dst_top_left_addr.col + dc,
            },
            value,
        );
    }
}

fn copy_range(
    sheet: &mut Sheet,
    sheet_name: &str,
    src: Range,
    dst_top_left: CellRef,
    formula_rewrites: &mut Vec<FormulaRewrite>,
) {
    let dst = Range::new(
        dst_top_left,
        CellRef::new(
            dst_top_left.row + src.height() - 1,
            dst_top_left.col + src.width() - 1,
        ),
    );
    let delta_row = dst.start.row as i32 - src.start.row as i32;
    let delta_col = dst.start.col as i32 - src.start.col as i32;

    let mut extracted: Vec<(CellRef, Option<Cell>)> = Vec::new();
    for cell in src.iter() {
        extracted.push((cell, sheet.cells.get(&cell_addr_from_cell_ref(cell)).cloned()));
    }

    for cell in dst.iter() {
        sheet.cells.remove(&cell_addr_from_cell_ref(cell));
    }

    for (cell, value) in extracted {
        let Some(mut value) = value else { continue };
        let dr = cell.row - src.start.row;
        let dc = cell.col - src.start.col;
        let target = CellRef::new(dst.start.row + dr, dst.start.col + dc);

        if let Some(formula) = &value.formula {
            let (new_formula, _) =
                rewrite_formula_for_copy_delta(formula, sheet_name, delta_row, delta_col);
            if &new_formula != formula {
                formula_rewrites.push(FormulaRewrite {
                    sheet: sheet_name.to_string(),
                    cell: target,
                    before: formula.clone(),
                    after: new_formula.clone(),
                });
            }
            value.formula = Some(new_formula);
        }

        sheet.cells.insert(cell_addr_from_cell_ref(target), value);
    }
}

fn fill_range(
    sheet: &mut Sheet,
    sheet_name: &str,
    src: Range,
    dst: Range,
    formula_rewrites: &mut Vec<FormulaRewrite>,
) {
    let height = src.height() as i32;
    let width = src.width() as i32;
    if height <= 0 || width <= 0 {
        return;
    }

    for cell in dst.iter() {
        if src.contains(cell) {
            continue;
        }
        sheet.cells.remove(&cell_addr_from_cell_ref(cell));

        let rel_row = cell.row as i32 - src.start.row as i32;
        let rel_col = cell.col as i32 - src.start.col as i32;
        let src_row = src.start.row + rel_row.rem_euclid(height) as u32;
        let src_col = src.start.col + rel_col.rem_euclid(width) as u32;
        let src_cell = CellRef::new(src_row, src_col);

        let Some(mut value) = sheet.cells.get(&cell_addr_from_cell_ref(src_cell)).cloned() else {
            continue;
        };
        if let Some(formula) = &value.formula {
            let delta_row = cell.row as i32 - src_cell.row as i32;
            let delta_col = cell.col as i32 - src_cell.col as i32;
            let (new_formula, _) =
                rewrite_formula_for_copy_delta(formula, sheet_name, delta_row, delta_col);
            if &new_formula != formula {
                formula_rewrites.push(FormulaRewrite {
                    sheet: sheet_name.to_string(),
                    cell,
                    before: formula.clone(),
                    after: new_formula.clone(),
                });
            }
            value.formula = Some(new_formula);
        }
        sheet.cells.insert(cell_addr_from_cell_ref(cell), value);
    }
}

fn rewrite_all_formulas_structural(
    workbook: &mut Workbook,
    sheet_names: &HashMap<SheetId, String>,
    edit: StructuralEdit,
) -> Vec<FormulaRewrite> {
    let mut rewrites = Vec::new();
    for (sheet_id, sheet) in workbook.sheets.iter_mut().enumerate() {
        let Some(ctx_sheet) = sheet_names.get(&sheet_id) else { continue };
        for (addr, cell) in sheet.cells.iter_mut() {
            let Some(formula) = &cell.formula else { continue };
            let (new_formula, changed) =
                rewrite_formula_for_structural_edit(formula, ctx_sheet, &edit);
            if changed {
                rewrites.push(FormulaRewrite {
                    sheet: ctx_sheet.clone(),
                    cell: cell_ref_from_addr(*addr),
                    before: formula.clone(),
                    after: new_formula.clone(),
                });
                cell.formula = Some(new_formula);
            }
        }
    }
    rewrites
}

fn rewrite_all_formulas_range_map(
    workbook: &mut Workbook,
    sheet_names: &HashMap<SheetId, String>,
    edit: &RangeMapEdit,
) -> Vec<FormulaRewrite> {
    let mut rewrites = Vec::new();
    for (sheet_id, sheet) in workbook.sheets.iter_mut().enumerate() {
        let Some(ctx_sheet) = sheet_names.get(&sheet_id) else { continue };
        for (addr, cell) in sheet.cells.iter_mut() {
            let Some(formula) = &cell.formula else { continue };
            let (new_formula, changed) = rewrite_formula_for_range_map(formula, ctx_sheet, edit);
            if changed {
                rewrites.push(FormulaRewrite {
                    sheet: ctx_sheet.clone(),
                    cell: cell_ref_from_addr(*addr),
                    before: formula.clone(),
                    after: new_formula.clone(),
                });
                cell.formula = Some(new_formula);
            }
        }
    }
    rewrites
}

fn diff_workbooks(
    before: &Workbook,
    after: &Workbook,
    sheet_names: &HashMap<SheetId, String>,
) -> Vec<CellChange> {
    let mut out = Vec::new();
    let max_sheets = max(before.sheets.len(), after.sheets.len());
    for sheet_id in 0..max_sheets {
        let sheet_name = sheet_names
            .get(&sheet_id)
            .cloned()
            .unwrap_or_else(|| format!("Sheet{sheet_id}"));
        let before_sheet = before.sheets.get(sheet_id);
        let after_sheet = after.sheets.get(sheet_id);
        let mut addrs: BTreeSet<CellAddr> = BTreeSet::new();
        if let Some(sheet) = before_sheet {
            addrs.extend(sheet.cells.keys().copied());
        }
        if let Some(sheet) = after_sheet {
            addrs.extend(sheet.cells.keys().copied());
        }
        for addr in addrs {
            let before_cell = before_sheet.and_then(|s| s.cells.get(&addr));
            let after_cell = after_sheet.and_then(|s| s.cells.get(&addr));
            let before_snap = before_cell.map(cell_snapshot);
            let after_snap = after_cell.map(cell_snapshot);
            if before_snap == after_snap {
                continue;
            }
            out.push(CellChange {
                sheet: sheet_name.clone(),
                cell: cell_ref_from_addr(addr),
                before: before_snap,
                after: after_snap,
            });
        }
    }
    out
}

fn cell_snapshot(cell: &Cell) -> CellSnapshot {
    CellSnapshot {
        value: cell.value.clone(),
        formula: cell.formula.clone(),
    }
}

fn sorted_cell_keys(set: &HashSet<CellKey>) -> Vec<CellKey> {
    let mut out: Vec<CellKey> = set.iter().copied().collect();
    out.sort_by_key(|k| (k.sheet, k.addr.row, k.addr.col));
    out
}

fn collect_transitive(map: &HashMap<CellKey, HashSet<CellKey>>, start: CellKey) -> Vec<CellKey> {
    let mut visited: HashSet<CellKey> = HashSet::new();
    let mut out: Vec<CellKey> = Vec::new();
    let mut queue = VecDeque::new();

    visited.insert(start);
    queue.push_back(start);

    while let Some(cell) = queue.pop_front() {
        let neighbors = map.get(&cell).map(sorted_cell_keys).unwrap_or_default();
        for n in neighbors {
            if visited.insert(n) {
                out.push(n);
                queue.push_back(n);
            }
        }
    }

    out.sort_by_key(|k| (k.sheet, k.addr.row, k.addr.col));
    out
}

struct Snapshot {
    sheets: HashSet<SheetId>,
    sheet_names_by_id: Vec<String>,
    values: HashMap<CellKey, Value>,
    tables: Vec<Vec<Table>>,
    workbook_names: HashMap<String, crate::eval::ResolvedName>,
    sheet_names: Vec<HashMap<String, crate::eval::ResolvedName>>,
    external_value_provider: Option<Arc<dyn ExternalValueProvider>>,
}

impl Snapshot {
    fn from_workbook(
        workbook: &Workbook,
        spills: &SpillState,
        external_value_provider: Option<Arc<dyn ExternalValueProvider>>,
    ) -> Self {
        let sheets: HashSet<SheetId> = (0..workbook.sheets.len()).collect();
        let sheet_names_by_id = workbook.sheet_names.clone();
        let mut values = HashMap::new();
        for (sheet_id, sheet) in workbook.sheets.iter().enumerate() {
            for (addr, cell) in &sheet.cells {
                values.insert(CellKey { sheet: sheet_id, addr: *addr }, cell.value.clone());
            }
        }

        // Overlay spilled values so formula evaluation can observe dynamic array results even
        // when the workbook map doesn't contain explicit cell records.
        for (origin, spill) in &spills.by_origin {
            for r in 0..spill.array.rows {
                for c in 0..spill.array.cols {
                    if r == 0 && c == 0 {
                        continue;
                    }
                    let addr = CellAddr {
                        row: origin.addr.row + r as u32,
                        col: origin.addr.col + c as u32,
                    };
                    let key = CellKey {
                        sheet: origin.sheet,
                        addr,
                    };
                    if let Some(v) = spill.array.get(r, c).cloned() {
                        values.insert(key, v);
                    }
                }
            }
        }
        let tables = workbook.sheets.iter().map(|s| s.tables.clone()).collect();

        let mut workbook_names = HashMap::new();
        for (name, def) in &workbook.names {
            workbook_names.insert(name.clone(), name_to_resolved(def));
        }

        let mut sheet_names = Vec::with_capacity(workbook.sheets.len());
        for sheet in &workbook.sheets {
            let mut names = HashMap::new();
            for (name, def) in &sheet.names {
                names.insert(name.clone(), name_to_resolved(def));
            }
            sheet_names.push(names);
        }

        fn name_to_resolved(def: &DefinedName) -> crate::eval::ResolvedName {
            match &def.definition {
                NameDefinition::Constant(v) => crate::eval::ResolvedName::Constant(v.clone()),
                NameDefinition::Reference(_) | NameDefinition::Formula(_) => crate::eval::ResolvedName::Expr(
                    def.compiled
                        .clone()
                        .expect("non-constant defined name must have compiled expression"),
                ),
            }
        }

        Self {
            sheets,
            sheet_names_by_id,
            values,
            tables,
            workbook_names,
            sheet_names,
            external_value_provider,
        }
    }
}

impl crate::eval::ValueResolver for Snapshot {
    fn sheet_exists(&self, sheet_id: usize) -> bool {
        self.sheets.contains(&sheet_id)
    }

    fn get_cell_value(&self, sheet_id: usize, addr: CellAddr) -> Value {
        if let Some(v) = self.values.get(&CellKey { sheet: sheet_id, addr }) {
            return v.clone();
        }

        if let Some(provider) = &self.external_value_provider {
            if let Some(sheet_name) = self.sheet_names_by_id.get(sheet_id) {
                if let Some(v) = provider.get(sheet_name, addr) {
                    return v;
                }
            }
        }

        Value::Blank
    }

    fn resolve_structured_ref(
        &self,
        ctx: crate::eval::EvalContext,
        sref: &crate::structured_refs::StructuredRef,
    ) -> Option<(usize, CellAddr, CellAddr)> {
        crate::structured_refs::resolve_structured_ref(&self.tables, ctx.current_sheet, ctx.current_cell, sref).ok()
    }

    fn resolve_name(&self, sheet_id: usize, name: &str) -> Option<crate::eval::ResolvedName> {
        let key = name.trim().to_ascii_uppercase();
        if let Some(map) = self.sheet_names.get(sheet_id) {
            if let Some(def) = map.get(&key) {
                return Some(def.clone());
            }
        }
        self.workbook_names.get(&key).cloned()
    }
}

fn resolve_defined_name<'a>(
    workbook: &'a Workbook,
    sheet_id: SheetId,
    name_key: &str,
) -> Option<&'a DefinedName> {
    workbook
        .sheets
        .get(sheet_id)
        .and_then(|s| s.names.get(name_key))
        .or_else(|| workbook.names.get(name_key))
}

pub trait ExternalValueProvider: Send + Sync {
    fn get(&self, sheet: &str, addr: CellAddr) -> Option<Value>;
}

fn analyze_expr(
    expr: &CompiledExpr,
    current_cell: CellKey,
    tables_by_sheet: &[Vec<Table>],
    workbook: &Workbook,
) -> (HashSet<CellKey>, HashSet<String>, bool, bool) {
    let mut precedents = HashSet::new();
    let mut names = HashSet::new();
    let mut volatile = false;
    let mut thread_safe = true;
    let mut visiting_names = HashSet::new();
    walk_expr(
        expr,
        current_cell,
        tables_by_sheet,
        workbook,
        &mut precedents,
        &mut names,
        &mut volatile,
        &mut thread_safe,
        &mut visiting_names,
    );
    (precedents, names, volatile, thread_safe)
}

const MAX_AUDIT_RANGE_EXPANSION_CELLS: u64 = 10_000;

fn walk_expr(
    expr: &CompiledExpr,
    current_cell: CellKey,
    tables_by_sheet: &[Vec<Table>],
    workbook: &Workbook,
    precedents: &mut HashSet<CellKey>,
    names: &mut HashSet<String>,
    volatile: &mut bool,
    thread_safe: &mut bool,
    visiting_names: &mut HashSet<(SheetId, String)>,
) {
    match expr {
        Expr::CellRef(r) => {
            if let Some(sheet) = resolve_sheet(&r.sheet, current_cell.sheet) {
                precedents.insert(CellKey { sheet, addr: r.addr });
            }
        }
        Expr::RangeRef(RangeRef { sheet, start, end }) => {
            if let Some(sheet) = resolve_sheet(sheet, current_cell.sheet) {
                let (r1, r2) = if start.row <= end.row {
                    (start.row, end.row)
                } else {
                    (end.row, start.row)
                };
                let (c1, c2) = if start.col <= end.col {
                    (start.col, end.col)
                } else {
                    (end.col, start.col)
                };

                let height = (r2 - r1 + 1) as u64;
                let width = (c2 - c1 + 1) as u64;
                let cell_count = height.saturating_mul(width);

                if cell_count <= MAX_AUDIT_RANGE_EXPANSION_CELLS {
                    for row in r1..=r2 {
                        for col in c1..=c2 {
                            precedents.insert(CellKey {
                                sheet,
                                addr: CellAddr { row, col },
                            });
                        }
                    }
                } else if let Some(sheet_cells) = workbook.sheets.get(sheet) {
                    // Avoid catastrophic expansion for full row/col references (e.g. `A:A`).
                    // For auditing, include only cells that currently exist in the sparse workbook.
                    for addr in sheet_cells.cells.keys() {
                        if addr.row >= r1 && addr.row <= r2 && addr.col >= c1 && addr.col <= c2 {
                            precedents.insert(CellKey { sheet, addr: *addr });
                        }
                    }
                }
            }
        }
        Expr::StructuredRef(sref) => {
            if let Ok((sheet_id, start, end)) = crate::structured_refs::resolve_structured_ref(
                tables_by_sheet,
                current_cell.sheet,
                current_cell.addr,
                sref,
            ) {
                let (r1, r2) = if start.row <= end.row {
                    (start.row, end.row)
                } else {
                    (end.row, start.row)
                };
                let (c1, c2) = if start.col <= end.col {
                    (start.col, end.col)
                } else {
                    (end.col, start.col)
                };

                let height = (r2 - r1 + 1) as u64;
                let width = (c2 - c1 + 1) as u64;
                let cell_count = height.saturating_mul(width);

                if cell_count <= MAX_AUDIT_RANGE_EXPANSION_CELLS {
                    for row in r1..=r2 {
                        for col in c1..=c2 {
                            precedents.insert(CellKey {
                                sheet: sheet_id,
                                addr: CellAddr { row, col },
                            });
                        }
                    }
                } else if let Some(sheet_cells) = workbook.sheets.get(sheet_id) {
                    for addr in sheet_cells.cells.keys() {
                        if addr.row >= r1 && addr.row <= r2 && addr.col >= c1 && addr.col <= c2 {
                            precedents.insert(CellKey { sheet: sheet_id, addr: *addr });
                        }
                    }
                }
            }
        }
        Expr::NameRef(nref) => {
            let Some(sheet) = resolve_sheet(&nref.sheet, current_cell.sheet) else {
                return;
            };
            let name_key = normalize_defined_name(&nref.name);
            if name_key.is_empty() {
                return;
            }
            names.insert(name_key.clone());
            let visit_key = (sheet, name_key.clone());
            if !visiting_names.insert(visit_key.clone()) {
                // Cycle in the name definition graph. Stop expanding to avoid infinite recursion;
                // evaluation will surface `#NAME?` via the runtime recursion guard.
                return;
            }
            if let Some(def) = resolve_defined_name(workbook, sheet, &name_key) {
                if let Some(expr) = def.compiled.as_ref() {
                    walk_expr(
                        expr,
                        CellKey {
                            sheet,
                            addr: current_cell.addr,
                        },
                        tables_by_sheet,
                        workbook,
                        precedents,
                        names,
                        volatile,
                        thread_safe,
                        visiting_names,
                    );
                }
            }
            visiting_names.remove(&visit_key);
        }
        Expr::Unary { expr, .. } | Expr::Postfix { expr, .. } => walk_expr(
            expr,
            current_cell,
            tables_by_sheet,
            workbook,
            precedents,
            names,
            volatile,
            thread_safe,
            visiting_names,
        ),
        Expr::Binary { left, right, .. } | Expr::Compare { left, right, .. } => {
            walk_expr(
                left,
                current_cell,
                tables_by_sheet,
                workbook,
                precedents,
                names,
                volatile,
                thread_safe,
                visiting_names,
            );
            walk_expr(
                right,
                current_cell,
                tables_by_sheet,
                workbook,
                precedents,
                names,
                volatile,
                thread_safe,
                visiting_names,
            );
        }
        Expr::FunctionCall { name, args, .. } => {
            if let Some(spec) = crate::functions::lookup_function(name) {
                if spec.volatility == crate::functions::Volatility::Volatile {
                    *volatile = true;
                }
                if spec.thread_safety == crate::functions::ThreadSafety::NotThreadSafe {
                    *thread_safe = false;
                }
            } else {
                // Placeholder: treat unknown/UDFs as non-thread-safe.
                *thread_safe = false;
            }
            for a in args {
                walk_expr(
                    a,
                    current_cell,
                    tables_by_sheet,
                    workbook,
                    precedents,
                    names,
                    volatile,
                    thread_safe,
                    visiting_names,
                );
            }
        }
        Expr::ImplicitIntersection(inner) => {
            walk_expr(
                inner,
                current_cell,
                tables_by_sheet,
                workbook,
                precedents,
                names,
                volatile,
                thread_safe,
                visiting_names,
            )
        }
        Expr::Number(_)
        | Expr::Text(_)
        | Expr::Bool(_)
        | Expr::Blank
        | Expr::Error(_) => {}
    }
}

fn analyze_calc_precedents(
    expr: &CompiledExpr,
    current_cell: CellKey,
    tables_by_sheet: &[Vec<Table>],
    workbook: &Workbook,
) -> HashSet<Precedent> {
    let mut out = HashSet::new();
    let mut visiting_names = HashSet::new();
    walk_calc_expr(
        expr,
        current_cell,
        tables_by_sheet,
        workbook,
        &mut out,
        &mut visiting_names,
    );
    out
}

fn walk_calc_expr(
    expr: &CompiledExpr,
    current_cell: CellKey,
    tables_by_sheet: &[Vec<Table>],
    workbook: &Workbook,
    precedents: &mut HashSet<Precedent>,
    visiting_names: &mut HashSet<(SheetId, String)>,
) {
    match expr {
        Expr::CellRef(r) => {
            if let Some(sheet) = resolve_sheet(&r.sheet, current_cell.sheet) {
                precedents.insert(Precedent::Cell(CellId::new(
                    sheet_id_for_graph(sheet),
                    r.addr.row,
                    r.addr.col,
                )));
            }
        }
        Expr::RangeRef(RangeRef { sheet, start, end }) => {
            if let Some(sheet) = resolve_sheet(sheet, current_cell.sheet) {
                let range = Range::new(
                    CellRef::new(start.row, start.col),
                    CellRef::new(end.row, end.col),
                );
                precedents.insert(Precedent::Range(SheetRange::new(
                    sheet_id_for_graph(sheet),
                    range,
                )));
            }
        }
        Expr::StructuredRef(sref) => {
            if let Ok((sheet_id, start, end)) = crate::structured_refs::resolve_structured_ref(
                tables_by_sheet,
                current_cell.sheet,
                current_cell.addr,
                sref,
            ) {
                let range = Range::new(
                    CellRef::new(start.row, start.col),
                    CellRef::new(end.row, end.col),
                );
                precedents.insert(Precedent::Range(SheetRange::new(
                    sheet_id_for_graph(sheet_id),
                    range,
                )));
            }
        }
        Expr::NameRef(nref) => {
            let Some(sheet) = resolve_sheet(&nref.sheet, current_cell.sheet) else {
                return;
            };
            let name_key = normalize_defined_name(&nref.name);
            if name_key.is_empty() {
                return;
            }
            let visit_key = (sheet, name_key.clone());
            if !visiting_names.insert(visit_key.clone()) {
                return;
            }
            if let Some(def) = resolve_defined_name(workbook, sheet, &name_key) {
                if let Some(expr) = def.compiled.as_ref() {
                    walk_calc_expr(
                        expr,
                        CellKey {
                            sheet,
                            addr: current_cell.addr,
                        },
                        tables_by_sheet,
                        workbook,
                        precedents,
                        visiting_names,
                    );
                }
            }
            visiting_names.remove(&visit_key);
        }
        Expr::Unary { expr, .. } | Expr::Postfix { expr, .. } => {
            walk_calc_expr(expr, current_cell, tables_by_sheet, workbook, precedents, visiting_names)
        }
        Expr::Binary { left, right, .. } | Expr::Compare { left, right, .. } => {
            walk_calc_expr(left, current_cell, tables_by_sheet, workbook, precedents, visiting_names);
            walk_calc_expr(right, current_cell, tables_by_sheet, workbook, precedents, visiting_names);
        }
        Expr::FunctionCall { args, .. } => {
            for a in args {
                walk_calc_expr(a, current_cell, tables_by_sheet, workbook, precedents, visiting_names);
            }
        }
        Expr::ImplicitIntersection(inner) => {
            walk_calc_expr(inner, current_cell, tables_by_sheet, workbook, precedents, visiting_names)
        }
        Expr::Number(_) | Expr::Text(_) | Expr::Bool(_) | Expr::Blank | Expr::Error(_) => {}
    }
}

fn resolve_sheet(sheet: &SheetReference<usize>, current_sheet: SheetId) -> Option<SheetId> {
    match sheet {
        SheetReference::Current => Some(current_sheet),
        SheetReference::Sheet(id) => Some(*id),
        SheetReference::External(_) => None,
    }
}

fn sheet_id_for_graph(sheet: SheetId) -> u32 {
    sheet.try_into().expect("sheet id exceeds u32")
}

fn cell_id_from_key(key: CellKey) -> CellId {
    CellId::new(sheet_id_for_graph(key.sheet), key.addr.row, key.addr.col)
}

fn cell_key_from_id(id: CellId) -> CellKey {
    CellKey {
        sheet: usize::try_from(id.sheet_id).expect("sheet id exceeds usize"),
        addr: CellAddr {
            row: id.cell.row,
            col: id.cell.col,
        },
    }
}

fn value_delta(old: &Value, new: &Value) -> f64 {
    match (numeric_value(old), numeric_value(new)) {
        (Some(a), Some(b)) => (a - b).abs(),
        _ if old == new => 0.0,
        _ => f64::INFINITY,
    }
}

fn numeric_value(value: &Value) -> Option<f64> {
    match value {
        Value::Number(n) => Some(*n),
        Value::Blank => Some(0.0),
        Value::Bool(b) => Some(if *b { 1.0 } else { 0.0 }),
        Value::Text(_) | Value::Error(_) | Value::Array(_) | Value::Spill { .. } => None,
    }
}

fn normalize_defined_name(name: &str) -> String {
    name.trim().to_ascii_uppercase()
}
