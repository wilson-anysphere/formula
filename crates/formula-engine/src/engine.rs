use crate::bytecode;
use crate::calc_settings::{CalcSettings, CalculationMode};
use crate::date::ExcelDateSystem;
use crate::editing::rewrite::{
    rewrite_formula_for_copy_delta, rewrite_formula_for_range_map_with_resolver,
    rewrite_formula_for_structural_edit_with_resolver, GridRange, RangeMapEdit, StructuralEdit,
};
use crate::editing::{
    CellChange, CellSnapshot, EditError, EditOp, EditResult, FormulaRewrite, MovedRange,
};
use crate::eval::{
    compile_canonical_expr, lower_ast, parse_a1, CellAddr, CompiledExpr, Expr, FormulaParseError,
    RangeRef, SheetReference, ValueResolver,
};
use crate::graph::{CellDeps, DependencyGraph as CalcGraph, Precedent, SheetRange};
use crate::iterative;
use crate::locale::{canonicalize_formula, canonicalize_formula_with_style, FormulaLocale};
use crate::value::{Array, ErrorKind, Value};
use formula_model::{CellId, CellRef, Range, Table};
#[cfg(all(feature = "parallel", not(target_arch = "wasm32")))]
use rayon::prelude::*;
use std::cell::RefCell;
use std::cmp::max;
use std::collections::{BTreeMap, BTreeSet, HashMap, HashSet, VecDeque};
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

#[derive(Debug, Clone, PartialEq)]
pub struct RecalcValueChange {
    pub sheet: String,
    pub addr: CellAddr,
    pub value: Value,
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
enum CompiledFormula {
    Ast(CompiledExpr),
    Bytecode(BytecodeFormula),
}

impl CompiledFormula {
    fn ast(&self) -> &CompiledExpr {
        match self {
            CompiledFormula::Ast(expr) => expr,
            CompiledFormula::Bytecode(bc) => &bc.ast,
        }
    }
}

#[derive(Debug, Clone)]
struct BytecodeFormula {
    ast: CompiledExpr,
    program: Arc<bytecode::Program>,
}

#[derive(Debug, Clone)]
struct Cell {
    value: Value,
    formula: Option<String>,
    compiled: Option<CompiledFormula>,
    volatile: bool,
    thread_safe: bool,
    dynamic_deps: bool,
}

impl Default for Cell {
    fn default() -> Self {
        Self {
            value: Value::Blank,
            formula: None,
            compiled: None,
            volatile: false,
            thread_safe: true,
            dynamic_deps: false,
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

    fn set_tables(&mut self, sheet: SheetId, tables: Vec<Table>) {
        if let Some(s) = self.sheets.get_mut(sheet) {
            s.tables = tables;
        }
    }
}

/// A node returned from auditing/introspection APIs (precedents/dependents).
///
/// This can represent either a single cell or a rectangular range reference without
/// expanding it into per-cell nodes (which is prohibitive for `A:A`, `1:1`, etc).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PrecedentNode {
    Cell {
        sheet: SheetId,
        addr: CellAddr,
    },
    Range {
        sheet: SheetId,
        start: CellAddr,
        end: CellAddr,
    },
    /// Dynamic array spill footprint for a spilled formula (origin -> footprint).
    ///
    /// Dynamic array evaluation is not implemented yet, but the auditing API reserves a
    /// node type so the UI can represent spill relationships without expanding every
    /// cell in the footprint.
    SpillRange {
        sheet: SheetId,
        origin: CellAddr,
        start: CellAddr,
        end: CellAddr,
    },
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum DirtyReason {
    Cell(CellKey),
    ViaRange { from: CellKey, range: PrecedentNode },
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
    bytecode_cache: bytecode::BytecodeCache,
    external_value_provider: Option<Arc<dyn ExternalValueProvider>>,
    name_dependents: HashMap<String, HashSet<CellKey>>,
    cell_name_refs: HashMap<CellKey, HashSet<String>>,
    /// Optimized dependency graph used for incremental recalculation ordering.
    calc_graph: CalcGraph,
    dirty: HashSet<CellKey>,
    dirty_reasons: HashMap<CellKey, DirtyReason>,
    calc_settings: CalcSettings,
    date_system: ExcelDateSystem,
    circular_references: HashSet<CellKey>,
    spills: SpillState,
    next_recalc_id: u64,
}

#[derive(Default)]
struct RecalcValueChangeCollector {
    before: HashMap<CellKey, Value>,
    after: BTreeMap<CellKey, Value>,
}

impl RecalcValueChangeCollector {
    fn record(&mut self, key: CellKey, before: Value, after: Value) {
        if before == after {
            return;
        }
        self.before.entry(key).or_insert(before);
        self.after.insert(key, after);
    }

    fn into_sorted_changes(self, workbook: &Workbook) -> Vec<RecalcValueChange> {
        let mut out = Vec::new();
        for (key, after) in self.after {
            let before = self
                .before
                .get(&key)
                .expect("recalc change must record before value");
            if *before == after {
                continue;
            }
            let sheet = workbook
                .sheet_names
                .get(key.sheet)
                .cloned()
                .unwrap_or_default();
            out.push(RecalcValueChange {
                sheet,
                addr: key.addr,
                value: after,
            });
        }
        out
    }
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
            bytecode_cache: bytecode::BytecodeCache::new(),
            external_value_provider: None,
            name_dependents: HashMap::new(),
            cell_name_refs: HashMap::new(),
            calc_graph: CalcGraph::new(),
            dirty: HashSet::new(),
            dirty_reasons: HashMap::new(),
            // Default to manual calculation to preserve historical engine behavior; callers can
            // opt into Excel-like automatic mode by setting `CalcSettings.calculation_mode`.
            calc_settings: CalcSettings {
                calculation_mode: CalculationMode::Manual,
                ..CalcSettings::default()
            },
            date_system: ExcelDateSystem::EXCEL_1900,
            circular_references: HashSet::new(),
            spills: SpillState::default(),
            next_recalc_id: 0,
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

    pub fn set_external_value_provider(
        &mut self,
        provider: Option<Arc<dyn ExternalValueProvider>>,
    ) {
        self.external_value_provider = provider;
    }

    pub fn bytecode_program_count(&self) -> usize {
        self.bytecode_cache.program_count()
    }

    pub fn set_date_system(&mut self, system: ExcelDateSystem) {
        if self.date_system == system {
            return;
        }
        self.date_system = system;

        for (sheet_id, sheet) in self.workbook.sheets.iter().enumerate() {
            for (addr, cell) in &sheet.cells {
                if cell.compiled.is_some() {
                    let key = CellKey {
                        sheet: sheet_id,
                        addr: *addr,
                    };
                    self.dirty.insert(key);
                    self.dirty_reasons.remove(&key);
                    self.calc_graph.mark_dirty(cell_id_from_key(key));
                }
            }
        }

        self.sync_dirty_from_calc_graph();
        if self.calc_settings.calculation_mode != CalculationMode::Manual {
            self.recalculate();
        }
    }

    pub fn date_system(&self) -> ExcelDateSystem {
        self.date_system
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
        let key = CellKey {
            sheet: sheet_id,
            addr,
        };
        let cell_id = cell_id_from_key(key);

        self.clear_spill_for_cell(key);
        self.clear_blocked_spill_for_origin(key);

        // Replace any existing formula and dependencies.
        self.calc_graph.remove_cell(cell_id);
        self.clear_cell_name_refs(key);
        self.dirty.remove(&key);
        self.dirty_reasons.remove(&key);

        let cell = self.workbook.get_or_create_cell_mut(key);
        cell.value = value.into();
        cell.formula = None;
        cell.compiled = None;
        cell.volatile = false;
        cell.thread_safe = true;
        cell.dynamic_deps = false;

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

    /// Clears a cell's stored value/formula so it behaves as if it does not exist.
    ///
    /// This is distinct from setting a cell to [`Value::Blank`]: clearing removes the
    /// corresponding entry from the sheet's sparse cell map, preserving sparsity and
    /// avoiding explicit blank entries for large cleared ranges.
    pub fn clear_cell(&mut self, sheet: &str, addr: &str) -> Result<(), EngineError> {
        let addr = parse_a1(addr)?;
        let Some(sheet_id) = self.workbook.sheet_id(sheet) else {
            return Ok(());
        };
        let key = CellKey {
            sheet: sheet_id,
            addr,
        };
        let cell_id = cell_id_from_key(key);

        self.clear_spill_for_cell(key);
        self.clear_blocked_spill_for_origin(key);

        // Remove any existing formula and dependencies.
        self.calc_graph.remove_cell(cell_id);
        self.clear_cell_name_refs(key);
        self.dirty.remove(&key);
        self.dirty_reasons.remove(&key);

        if let Some(sheet) = self.workbook.sheets.get_mut(sheet_id) {
            sheet.cells.remove(&addr);
        }

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

        let tables_by_sheet: Vec<Vec<Table>> = self
            .workbook
            .sheets
            .iter()
            .map(|s| s.tables.clone())
            .collect();

        // Structured reference resolution can change which cells a formula depends on, so refresh
        // dependencies for all formulas.
        let mut formulas: Vec<(CellKey, CompiledExpr)> = Vec::new();
        for (sheet_id, sheet) in self.workbook.sheets.iter().enumerate() {
            for (addr, cell) in &sheet.cells {
                if let Some(compiled) = cell.compiled.as_ref() {
                    formulas.push((
                        CellKey {
                            sheet: sheet_id,
                            addr: *addr,
                        },
                        compiled.ast().clone(),
                    ));
                }
            }
        }

        for (key, ast) in formulas {
            let cell_id = cell_id_from_key(key);
            let (names, volatile, thread_safe, dynamic_deps) =
                analyze_expr_flags(&ast, key, &tables_by_sheet, &self.workbook);
            self.set_cell_name_refs(key, names);

            let calc_precedents =
                analyze_calc_precedents(&ast, key, &tables_by_sheet, &self.workbook, &self.spills);
            let mut calc_vec: Vec<Precedent> = calc_precedents.into_iter().collect();
            calc_vec.sort_by_key(|p| match p {
                Precedent::Cell(c) => (0u8, c.sheet_id, c.cell.row, c.cell.col, 0u32, 0u32),
                Precedent::Range(r) => (
                    1u8,
                    r.sheet_id,
                    r.range.start.row,
                    r.range.start.col,
                    r.range.end.row,
                    r.range.end.col,
                ),
            });
            let deps = CellDeps::new(calc_vec).volatile(volatile);
            self.calc_graph.update_cell_dependencies(cell_id, deps);

            let cell = self.workbook.get_or_create_cell_mut(key);
            cell.volatile = volatile;
            cell.thread_safe = thread_safe;
            cell.dynamic_deps = dynamic_deps;

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
                let ast = crate::parse_formula(
                    formula,
                    crate::ParseOptions {
                        locale: crate::LocaleConfig::en_us(),
                        reference_style: crate::ReferenceStyle::A1,
                        normalize_relative_to: None,
                    },
                )?;
                let parsed = lower_ast(&ast, None);
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
                self.workbook
                    .sheets
                    .get_mut(sheet_id)?
                    .names
                    .remove(&name_key)
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
                self.workbook
                    .sheets
                    .get(sheet_id)?
                    .names
                    .get(&name_key)
                    .map(|n| &n.definition)
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
        let key = CellKey {
            sheet: sheet_id,
            addr,
        };
        let cell_id = cell_id_from_key(key);
        self.clear_spill_for_cell(key);
        self.clear_blocked_spill_for_origin(key);

        let origin = crate::CellAddr::new(addr.row, addr.col);
        let parsed = crate::parse_formula(
            formula,
            crate::ParseOptions {
                locale: crate::LocaleConfig::en_us(),
                reference_style: crate::ReferenceStyle::A1,
                normalize_relative_to: Some(origin),
            },
        )?;
        let mut resolve_sheet = |name: &str| self.workbook.sheet_id(name);
        let compiled = compile_canonical_expr(&parsed.expr, sheet_id, addr, &mut resolve_sheet);
        let tables_by_sheet: Vec<Vec<Table>> = self
            .workbook
            .sheets
            .iter()
            .map(|s| s.tables.clone())
            .collect();
        let (names, volatile, thread_safe, dynamic_deps) =
            analyze_expr_flags(&compiled, key, &tables_by_sheet, &self.workbook);
        self.set_cell_name_refs(key, names);

        // Optimized precedents for calculation ordering (range nodes are not expanded).
        let calc_precedents = analyze_calc_precedents(
            &compiled,
            key,
            &tables_by_sheet,
            &self.workbook,
            &self.spills,
        );
        let mut calc_vec: Vec<Precedent> = calc_precedents.into_iter().collect();
        calc_vec.sort_by_key(|p| match p {
            Precedent::Cell(c) => (0u8, c.sheet_id, c.cell.row, c.cell.col, 0u32, 0u32),
            Precedent::Range(r) => (
                1u8,
                r.sheet_id,
                r.range.start.row,
                r.range.start.col,
                r.range.end.row,
                r.range.end.col,
            ),
        });
        let deps = CellDeps::new(calc_vec).volatile(volatile);
        self.calc_graph.update_cell_dependencies(cell_id, deps);

        let compiled_formula =
            match self.try_compile_bytecode(&parsed.expr, key, volatile, thread_safe) {
                Some(program) => CompiledFormula::Bytecode(BytecodeFormula {
                    ast: compiled.clone(),
                    program,
                }),
                None => CompiledFormula::Ast(compiled),
            };

        let cell = self.workbook.get_or_create_cell_mut(key);
        cell.formula = Some(formula.to_string());
        cell.compiled = Some(compiled_formula);
        cell.volatile = volatile;
        cell.thread_safe = thread_safe;
        cell.dynamic_deps = dynamic_deps;

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
        let key = CellKey {
            sheet: sheet_id,
            addr,
        };
        if let Some(v) = self.spilled_cell_value(key) {
            return v;
        }
        if let Some(cell) = self.workbook.get_cell(key) {
            return cell.value.clone();
        }

        if let Some(provider) = &self.external_value_provider {
            // Use the workbook's canonical display name to keep provider lookups stable even when
            // callers pass a different sheet-name casing.
            if let Some(sheet_name) = self.workbook.sheet_names.get(sheet_id) {
                if let Some(v) = provider.get(sheet_name, addr) {
                    return v;
                }
            }
        }

        Value::Blank
    }

    /// Returns the spill range (origin inclusive) for a cell if it is an array-spill
    /// origin or belongs to a spilled range.
    pub fn spill_range(&self, sheet: &str, addr: &str) -> Option<(CellAddr, CellAddr)> {
        let sheet_id = self.workbook.sheet_id(sheet)?;
        let addr = parse_a1(addr).ok()?;
        let key = CellKey {
            sheet: sheet_id,
            addr,
        };
        let origin = self.spill_origin_key(key)?;
        let spill = self.spills.by_origin.get(&origin)?;
        Some((origin.addr, spill.end))
    }

    /// Returns the spill origin for a cell if it is an array-spill origin or belongs
    /// to a spilled range.
    pub fn spill_origin(&self, sheet: &str, addr: &str) -> Option<(SheetId, CellAddr)> {
        let sheet_id = self.workbook.sheet_id(sheet)?;
        let addr = parse_a1(addr).ok()?;
        let key = CellKey {
            sheet: sheet_id,
            addr,
        };
        let origin = self.spill_origin_key(key)?;
        Some((origin.sheet, origin.addr))
    }

    pub fn get_cell_formula(&self, sheet: &str, addr: &str) -> Option<&str> {
        let sheet_id = self.workbook.sheet_id(sheet)?;
        let addr = parse_a1(addr).ok()?;
        let key = CellKey {
            sheet: sheet_id,
            addr,
        };
        self.workbook.get_cell(key)?.formula.as_deref()
    }

    /// Returns the formula for `addr` rendered in R1C1 reference style.
    ///
    /// The engine persists formulas as canonical A1 strings; this converts them on demand using the
    /// syntax-only parser/serializer.
    pub fn get_cell_formula_r1c1(&self, sheet: &str, addr: &str) -> Option<String> {
        let sheet_id = self.workbook.sheet_id(sheet)?;
        let addr = parse_a1(addr).ok()?;
        let key = CellKey {
            sheet: sheet_id,
            addr,
        };
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
                let edit = StructuralEdit::InsertRows {
                    sheet: sheet.clone(),
                    row,
                    count,
                };
                self.rewrite_defined_names_structural(&sheet_names, &edit)
                    .map_err(|e| EditError::Engine(e.to_string()))?;
                formula_rewrites.extend(rewrite_all_formulas_structural(
                    &mut self.workbook,
                    &sheet_names,
                    edit,
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
                let edit = StructuralEdit::DeleteRows {
                    sheet: sheet.clone(),
                    row,
                    count,
                };
                self.rewrite_defined_names_structural(&sheet_names, &edit)
                    .map_err(|e| EditError::Engine(e.to_string()))?;
                formula_rewrites.extend(rewrite_all_formulas_structural(
                    &mut self.workbook,
                    &sheet_names,
                    edit,
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
                let edit = StructuralEdit::InsertCols {
                    sheet: sheet.clone(),
                    col,
                    count,
                };
                self.rewrite_defined_names_structural(&sheet_names, &edit)
                    .map_err(|e| EditError::Engine(e.to_string()))?;
                formula_rewrites.extend(rewrite_all_formulas_structural(
                    &mut self.workbook,
                    &sheet_names,
                    edit,
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
                let edit = StructuralEdit::DeleteCols {
                    sheet: sheet.clone(),
                    col,
                    count,
                };
                self.rewrite_defined_names_structural(&sheet_names, &edit)
                    .map_err(|e| EditError::Engine(e.to_string()))?;
                formula_rewrites.extend(rewrite_all_formulas_structural(
                    &mut self.workbook,
                    &sheet_names,
                    edit,
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
                self.rewrite_defined_names_range_map(&sheet_names, &edit)
                    .map_err(|e| EditError::Engine(e.to_string()))?;
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
                self.rewrite_defined_names_range_map(&sheet_names, &edit)
                    .map_err(|e| EditError::Engine(e.to_string()))?;
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
                    moved_region: GridRange::new(
                        range.start.row,
                        start_col,
                        range.end.row,
                        u32::MAX,
                    ),
                    delta_row: 0,
                    delta_col: -(width as i32),
                    deleted_region: Some(GridRange::new(
                        range.start.row,
                        range.start.col,
                        range.end.row,
                        range.end.col,
                    )),
                };
                self.rewrite_defined_names_range_map(&sheet_names, &edit)
                    .map_err(|e| EditError::Engine(e.to_string()))?;
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
                    moved_region: GridRange::new(
                        start_row,
                        range.start.col,
                        u32::MAX,
                        range.end.col,
                    ),
                    delta_row: -(height as i32),
                    delta_col: 0,
                    deleted_region: Some(GridRange::new(
                        range.start.row,
                        range.start.col,
                        range.end.row,
                        range.end.col,
                    )),
                };
                self.rewrite_defined_names_range_map(&sheet_names, &edit)
                    .map_err(|e| EditError::Engine(e.to_string()))?;
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
                    moved_region: GridRange::new(
                        src.start.row,
                        src.start.col,
                        src.end.row,
                        src.end.col,
                    ),
                    delta_row: dst.start.row as i32 - src.start.row as i32,
                    delta_col: dst.start.col as i32 - src.start.col as i32,
                    deleted_region: None,
                };
                self.rewrite_defined_names_range_map(&sheet_names, &edit)
                    .map_err(|e| EditError::Engine(e.to_string()))?;
                formula_rewrites.extend(rewrite_all_formulas_range_map(
                    &mut self.workbook,
                    &sheet_names,
                    &edit,
                ));
                moved_ranges.push(MovedRange {
                    sheet,
                    from: src,
                    to: dst,
                });
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
        #[cfg(all(feature = "parallel", not(target_arch = "wasm32")))]
        {
            self.recalculate_with_mode(RecalcMode::MultiThreaded);
        }
        #[cfg(any(target_arch = "wasm32", not(feature = "parallel")))]
        {
            self.recalculate_with_mode(RecalcMode::SingleThreaded);
        }
    }

    pub fn recalculate_single_threaded(&mut self) {
        self.recalculate_with_mode(RecalcMode::SingleThreaded);
    }

    pub fn recalculate_multi_threaded(&mut self) {
        #[cfg(all(feature = "parallel", not(target_arch = "wasm32")))]
        {
            self.recalculate_with_mode(RecalcMode::MultiThreaded);
        }
        #[cfg(any(target_arch = "wasm32", not(feature = "parallel")))]
        {
            self.recalculate_with_mode(RecalcMode::SingleThreaded);
        }
    }

    fn recalculate_with_mode(&mut self, mode: RecalcMode) {
        self.recalculate_with_mode_and_value_changes(mode, None);
    }

    pub fn recalculate_with_value_changes(&mut self, mode: RecalcMode) -> Vec<RecalcValueChange> {
        let mut changes = RecalcValueChangeCollector::default();
        self.recalculate_with_mode_and_value_changes(mode, Some(&mut changes));
        changes.into_sorted_changes(&self.workbook)
    }

    pub fn recalculate_with_value_changes_single_threaded(&mut self) -> Vec<RecalcValueChange> {
        self.recalculate_with_value_changes(RecalcMode::SingleThreaded)
    }

    pub fn recalculate_with_value_changes_multi_threaded(&mut self) -> Vec<RecalcValueChange> {
        self.recalculate_with_value_changes(RecalcMode::MultiThreaded)
    }

    fn recalculate_with_mode_and_value_changes(
        &mut self,
        mode: RecalcMode,
        mut value_changes: Option<&mut RecalcValueChangeCollector>,
    ) {
        let date_system = self.date_system;
        // Spill recalculation can introduce new dirty cells (spill outputs becoming
        // computed/cleared). These should be resolved as part of the same recalc "tick",
        // sharing a single `RecalcContext` so volatile functions remain stable.
        let mut recalc_ctx: Option<crate::eval::RecalcContext> = None;
        loop {
            let levels = match self.calc_graph.calc_levels_for_dirty() {
                Ok(levels) => levels,
                Err(_) => {
                    self.recalculate_with_cycles(mode, value_changes);
                    return;
                }
            };

            if levels.is_empty() {
                return;
            }

            let recalc_ctx = recalc_ctx.get_or_insert_with(|| self.begin_recalc_context());

            let (spill_dirty_roots, dynamic_dirty_roots) = self.recalculate_levels(
                levels,
                mode,
                recalc_ctx,
                date_system,
                value_changes.as_deref_mut(),
            );
            if spill_dirty_roots.is_empty() && dynamic_dirty_roots.is_empty() {
                return;
            }

            // Spills can change which cells are considered inputs vs computed spill outputs.
            // When the spill footprint changes (new cells gain/lose values), mark the affected
            // coordinates dirty so any dependents recalculate with the updated spill state.
            for cell in spill_dirty_roots.into_iter().chain(dynamic_dirty_roots) {
                self.calc_graph.mark_dirty(cell);
            }
        }
    }

    fn recalculate_levels(
        &mut self,
        levels: Vec<Vec<CellId>>,
        mode: RecalcMode,
        recalc_ctx: &crate::eval::RecalcContext,
        date_system: ExcelDateSystem,
        mut value_changes: Option<&mut RecalcValueChangeCollector>,
    ) -> (Vec<CellId>, Vec<CellId>) {
        self.circular_references.clear();

        let mut snapshot = Snapshot::from_workbook(
            &self.workbook,
            &self.spills,
            self.external_value_provider.clone(),
        );
        let mut spill_dirty_roots: Vec<CellId> = Vec::new();
        let mut dynamic_dirty_roots: Vec<CellId> = Vec::new();

        for level in levels {
            let mut keys: Vec<CellKey> = level.into_iter().map(cell_key_from_id).collect();
            keys.sort_by_key(|k| (k.sheet, k.addr.row, k.addr.col));

            let mut parallel_tasks: Vec<(CellKey, CompiledFormula)> = Vec::new();
            let mut serial_tasks: Vec<(CellKey, CompiledFormula)> = Vec::new();
            let mut dynamic_tasks: Vec<(CellKey, CompiledFormula)> = Vec::new();

            for &k in &keys {
                let Some(cell) = self.workbook.get_cell(k) else {
                    continue;
                };
                let Some(compiled) = cell.compiled.clone() else {
                    continue;
                };

                if cell.dynamic_deps {
                    dynamic_tasks.push((k, compiled));
                } else if cell.thread_safe {
                    parallel_tasks.push((k, compiled));
                } else {
                    serial_tasks.push((k, compiled));
                }
            }

            let mut all_tasks: Vec<(CellKey, CompiledFormula)> =
                Vec::with_capacity(parallel_tasks.len() + serial_tasks.len() + dynamic_tasks.len());
            all_tasks.extend(parallel_tasks.iter().cloned());
            all_tasks.extend(serial_tasks.iter().cloned());
            all_tasks.extend(dynamic_tasks.iter().cloned());

            let sheet_count = self.workbook.sheets.len();
            let column_cache = BytecodeColumnCache::build(sheet_count, &snapshot, &all_tasks);
            let empty_cols: HashMap<i32, BytecodeColumn> = HashMap::new();

            let mut results: Vec<(CellKey, Value)> =
                Vec::with_capacity(parallel_tasks.len() + serial_tasks.len());
            let eval_parallel_tasks_serial = |results: &mut Vec<(CellKey, Value)>| {
                let mut vm = bytecode::Vm::with_capacity(32);
                for (k, compiled) in &parallel_tasks {
                    let ctx = crate::eval::EvalContext {
                        current_sheet: k.sheet,
                        current_cell: k.addr,
                    };
                    let value = match compiled {
                        CompiledFormula::Ast(expr) => {
                            let evaluator = crate::eval::Evaluator::new_with_date_system(
                                &snapshot,
                                ctx,
                                recalc_ctx,
                                date_system,
                            );
                            evaluator.eval_formula(expr)
                        }
                        CompiledFormula::Bytecode(bc) => {
                            let cols = column_cache.by_sheet.get(k.sheet).unwrap_or(&empty_cols);
                            let slice_mode = slice_mode_for_program(&bc.program);
                            let grid = EngineBytecodeGrid {
                                snapshot: &snapshot,
                                sheet: k.sheet,
                                cols,
                                slice_mode,
                            };
                            let base = bytecode::CellCoord {
                                row: k.addr.row as i32,
                                col: k.addr.col as i32,
                            };
                            let v = vm.eval(&bc.program, &grid, base);
                            bytecode_value_to_engine(v)
                        }
                    };
                    results.push((*k, value));
                }
            };

            if mode == RecalcMode::MultiThreaded {
                #[cfg(all(feature = "parallel", not(target_arch = "wasm32")))]
                {
                    results.extend(
                        parallel_tasks
                            .par_iter()
                            .map_init(
                                || bytecode::Vm::with_capacity(32),
                                |vm, (k, compiled)| {
                                    let ctx = crate::eval::EvalContext {
                                        current_sheet: k.sheet,
                                        current_cell: k.addr,
                                    };
                                    match compiled {
                                        CompiledFormula::Ast(expr) => {
                                            let evaluator =
                                                crate::eval::Evaluator::new_with_date_system(
                                                    &snapshot,
                                                    ctx,
                                                    recalc_ctx,
                                                    date_system,
                                                );
                                            (*k, evaluator.eval_formula(expr))
                                        }
                                        CompiledFormula::Bytecode(bc) => {
                                            let cols = column_cache
                                                .by_sheet
                                                .get(k.sheet)
                                                .unwrap_or(&empty_cols);
                                            let slice_mode = slice_mode_for_program(&bc.program);
                                            let grid = EngineBytecodeGrid {
                                                snapshot: &snapshot,
                                                sheet: k.sheet,
                                                cols,
                                                slice_mode,
                                            };
                                            let base = bytecode::CellCoord {
                                                row: k.addr.row as i32,
                                                col: k.addr.col as i32,
                                            };
                                            let v = vm.eval(&bc.program, &grid, base);
                                            (*k, bytecode_value_to_engine(v))
                                        }
                                    }
                                },
                            )
                            .collect::<Vec<_>>(),
                    );
                }

                #[cfg(not(all(feature = "parallel", not(target_arch = "wasm32"))))]
                {
                    eval_parallel_tasks_serial(&mut results);
                }
            } else {
                eval_parallel_tasks_serial(&mut results);
            }

            // Non-thread-safe tasks are always serialized.
            let mut vm = bytecode::Vm::with_capacity(32);
            for (k, compiled) in &serial_tasks {
                let ctx = crate::eval::EvalContext {
                    current_sheet: k.sheet,
                    current_cell: k.addr,
                };
                let value = match compiled {
                    CompiledFormula::Ast(expr) => {
                        let evaluator = crate::eval::Evaluator::new_with_date_system(
                            &snapshot,
                            ctx,
                            recalc_ctx,
                            date_system,
                        );
                        evaluator.eval_formula(expr)
                    }
                    CompiledFormula::Bytecode(bc) => {
                        let cols = column_cache.by_sheet.get(k.sheet).unwrap_or(&empty_cols);
                        let slice_mode = slice_mode_for_program(&bc.program);
                        let grid = EngineBytecodeGrid {
                            snapshot: &snapshot,
                            sheet: k.sheet,
                            cols,
                            slice_mode,
                        };
                        let base = bytecode::CellCoord {
                            row: k.addr.row as i32,
                            col: k.addr.col as i32,
                        };
                        let v = vm.eval(&bc.program, &grid, base);
                        bytecode_value_to_engine(v)
                    }
                };
                results.push((*k, value));
            }

            results.sort_by_key(|(k, _)| (k.sheet, k.addr.row, k.addr.col));

            for (k, v) in results {
                self.apply_eval_result(
                    k,
                    v,
                    &mut snapshot,
                    &mut spill_dirty_roots,
                    value_changes.as_deref_mut(),
                );
            }

            // Dynamic-reference formulas (e.g. INDIRECT/OFFSET) must be evaluated serially so we
            // can trace their runtime precedents and update the dependency graph deterministically.
            for (k, compiled) in &dynamic_tasks {
                let ctx = crate::eval::EvalContext {
                    current_sheet: k.sheet,
                    current_cell: k.addr,
                };

                let trace = RefCell::new(crate::eval::DependencyTrace::default());

                let value = match compiled {
                    CompiledFormula::Ast(expr) => {
                        let evaluator = crate::eval::Evaluator::new_with_date_system(
                            &snapshot,
                            ctx,
                            recalc_ctx,
                            date_system,
                        )
                        .with_dependency_trace(&trace);
                        evaluator.eval_formula(expr)
                    }
                    CompiledFormula::Bytecode(bc) => {
                        // Dynamic dependency tracing is only supported for AST formulas. Fallback
                        // to bytecode evaluation without dependency updates.
                        let cols = column_cache.by_sheet.get(k.sheet).unwrap_or(&empty_cols);
                        let slice_mode = slice_mode_for_program(&bc.program);
                        let grid = EngineBytecodeGrid {
                            snapshot: &snapshot,
                            sheet: k.sheet,
                            cols,
                            slice_mode,
                        };
                        let base = bytecode::CellCoord {
                            row: k.addr.row as i32,
                            col: k.addr.col as i32,
                        };
                        let v = vm.eval(&bc.program, &grid, base);
                        bytecode_value_to_engine(v)
                    }
                };

                self.apply_eval_result(
                    *k,
                    value,
                    &mut snapshot,
                    &mut spill_dirty_roots,
                    value_changes.as_deref_mut(),
                );

                let CompiledFormula::Ast(expr) = compiled else {
                    continue;
                };

                let cell_id = cell_id_from_key(*k);
                let old_precedents: HashSet<Precedent> =
                    self.calc_graph.precedents_of(cell_id).into_iter().collect();

                let mut new_precedents: HashSet<Precedent> = analyze_calc_precedents(
                    expr,
                    *k,
                    &snapshot.tables,
                    &self.workbook,
                    &self.spills,
                );
                for reference in trace.borrow().precedents() {
                    let sheet_id = sheet_id_for_graph(reference.sheet_id);
                    if reference.start == reference.end {
                        new_precedents.insert(Precedent::Cell(CellId::new(
                            sheet_id,
                            reference.start.row,
                            reference.start.col,
                        )));
                    } else {
                        let range = Range::new(
                            CellRef::new(reference.start.row, reference.start.col),
                            CellRef::new(reference.end.row, reference.end.col),
                        );
                        new_precedents.insert(Precedent::Range(SheetRange::new(sheet_id, range)));
                    }
                }

                if new_precedents != old_precedents {
                    let mut vec: Vec<Precedent> = new_precedents.into_iter().collect();
                    vec.sort_by_key(|p| match p {
                        Precedent::Cell(c) => (0u8, c.sheet_id, c.cell.row, c.cell.col, 0u32, 0u32),
                        Precedent::Range(r) => (
                            1u8,
                            r.sheet_id,
                            r.range.start.row,
                            r.range.start.col,
                            r.range.end.row,
                            r.range.end.col,
                        ),
                    });

                    let is_volatile = self
                        .workbook
                        .get_cell(*k)
                        .map(|c| c.volatile)
                        .unwrap_or(false);
                    self.calc_graph.update_cell_dependencies(
                        cell_id,
                        CellDeps::new(vec).volatile(is_volatile),
                    );
                    dynamic_dirty_roots.push(cell_id);
                }
            }
        }

        self.calc_graph.clear_dirty();
        self.dirty.clear();
        self.dirty_reasons.clear();

        (spill_dirty_roots, dynamic_dirty_roots)
    }

    fn recalculate_with_cycles(
        &mut self,
        _mode: RecalcMode,
        mut value_changes: Option<&mut RecalcValueChangeCollector>,
    ) {
        let mut impacted_ids: HashSet<CellId> = self.calc_graph.dirty_cells().into_iter().collect();
        impacted_ids.extend(self.calc_graph.volatile_cells());

        if impacted_ids.is_empty() {
            return;
        }

        let recalc_ctx = self.begin_recalc_context();

        self.circular_references.clear();

        let mut impacted: Vec<CellKey> = impacted_ids.into_iter().map(cell_key_from_id).collect();
        impacted.sort_by_key(|k| (k.sheet, k.addr.row, k.addr.col));

        let impacted_set: HashSet<CellKey> = impacted.iter().copied().collect();
        let mut edges: HashMap<CellKey, Vec<CellKey>> = HashMap::new();
        for &cell in &impacted {
            let cell_id = cell_id_from_key(cell);
            let mut out: Vec<CellKey> = self
                .calc_graph
                .direct_dependents(cell_id)
                .into_iter()
                .map(cell_key_from_id)
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
        let date_system = self.date_system;

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
                    .and_then(|c| c.compiled.as_ref().map(|compiled| compiled.ast().clone()))
                else {
                    continue;
                };
                let ctx = crate::eval::EvalContext {
                    current_sheet: k.sheet,
                    current_cell: k.addr,
                };
                let evaluator = crate::eval::Evaluator::new_with_date_system(
                    &snapshot,
                    ctx,
                    &recalc_ctx,
                    date_system,
                );
                let v = evaluator.eval_formula(&expr);
                self.apply_eval_result(
                    k,
                    v,
                    &mut snapshot,
                    &mut spill_dirty_roots,
                    value_changes.as_deref_mut(),
                );
                continue;
            }

            for &k in &scc {
                self.circular_references.insert(k);
            }

            if !self.calc_settings.iterative.enabled {
                for &k in &scc {
                    let v = Value::Number(0.0);
                    self.apply_eval_result(
                        k,
                        v,
                        &mut snapshot,
                        &mut spill_dirty_roots,
                        value_changes.as_deref_mut(),
                    );
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
                        .and_then(|c| c.compiled.as_ref().map(|compiled| compiled.ast().clone()))
                    else {
                        continue;
                    };
                    let old = snapshot.values.get(&k).cloned().unwrap_or(Value::Blank);
                    let ctx = crate::eval::EvalContext {
                        current_sheet: k.sheet,
                        current_cell: k.addr,
                    };
                    let evaluator = crate::eval::Evaluator::new_with_date_system(
                        &snapshot,
                        ctx,
                        &recalc_ctx,
                        date_system,
                    );
                    let new_val = evaluator.eval_formula(&expr);
                    max_delta = max_delta.max(value_delta(&old, &new_val));
                    self.apply_eval_result(
                        k,
                        new_val,
                        &mut snapshot,
                        &mut spill_dirty_roots,
                        value_changes.as_deref_mut(),
                    );
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
        mut value_changes: Option<&mut RecalcValueChangeCollector>,
    ) {
        // Clear any previously tracked spill blockage for this origin before applying the new
        // evaluation result.
        self.clear_blocked_spill_for_origin(key);

        match value {
            Value::Array(array) => {
                if array.rows == 0 || array.cols == 0 {
                    self.apply_eval_result(
                        key,
                        Value::Error(ErrorKind::Calc),
                        snapshot,
                        spill_dirty_roots,
                        value_changes,
                    );
                    return;
                }

                let mut spill_too_big = || {
                    let cleared = self.clear_spill_for_origin(key);
                    snapshot.spill_end_by_origin.remove(&key);
                    for cleared_key in cleared {
                        if let Some(changes) = value_changes.as_deref_mut() {
                            let before =
                                snapshot.get_cell_value(cleared_key.sheet, cleared_key.addr);
                            snapshot.values.remove(&cleared_key);
                            let after =
                                snapshot.get_cell_value(cleared_key.sheet, cleared_key.addr);
                            changes.record(cleared_key, before, after);
                        } else {
                            snapshot.values.remove(&cleared_key);
                        }
                        snapshot.spill_origin_by_cell.remove(&cleared_key);
                        spill_dirty_roots.push(cell_id_from_key(cleared_key));
                        self.append_blocked_spill_dirty_roots(cleared_key, spill_dirty_roots);
                    }

                    let after = Value::Error(ErrorKind::Spill);
                    if let Some(changes) = value_changes.as_deref_mut() {
                        let before = snapshot.get_cell_value(key.sheet, key.addr);
                        changes.record(key, before, after.clone());
                    }
                    let cell = self.workbook.get_or_create_cell_mut(key);
                    cell.value = after.clone();
                    snapshot.values.insert(key, after);
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
                        if let Some(changes) = value_changes.as_deref_mut() {
                            let before = snapshot.get_cell_value(key.sheet, key.addr);
                            changes.record(key, before, top_left.clone());
                        }
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
                                let spill_key = CellKey {
                                    sheet: key.sheet,
                                    addr,
                                };
                                if let Some(v) = array.get(r, c).cloned() {
                                    if let Some(changes) = value_changes.as_deref_mut() {
                                        let before = snapshot
                                            .get_cell_value(spill_key.sheet, spill_key.addr);
                                        changes.record(spill_key, before, v.clone());
                                    }
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
                snapshot.spill_end_by_origin.remove(&key);
                for cleared_key in cleared {
                    if let Some(changes) = value_changes.as_deref_mut() {
                        let before = snapshot.get_cell_value(cleared_key.sheet, cleared_key.addr);
                        snapshot.values.remove(&cleared_key);
                        let after = snapshot.get_cell_value(cleared_key.sheet, cleared_key.addr);
                        changes.record(cleared_key, before, after);
                    } else {
                        snapshot.values.remove(&cleared_key);
                    }
                    snapshot.spill_origin_by_cell.remove(&cleared_key);
                    spill_dirty_roots.push(cell_id_from_key(cleared_key));
                    self.append_blocked_spill_dirty_roots(cleared_key, spill_dirty_roots);
                }

                if let Some(blocker) = self.spill_blocker(key, &array) {
                    self.record_blocked_spill(key, blocker);
                    let after = Value::Error(ErrorKind::Spill);
                    if let Some(changes) = value_changes.as_deref_mut() {
                        let before = snapshot.get_cell_value(key.sheet, key.addr);
                        changes.record(key, before, after.clone());
                    }
                    let cell = self.workbook.get_or_create_cell_mut(key);
                    cell.value = after.clone();
                    snapshot.values.insert(key, after);
                    return;
                }

                self.apply_new_spill(key, end, array, snapshot, spill_dirty_roots, value_changes);
            }
            other => {
                let cleared = self.clear_spill_for_origin(key);
                snapshot.spill_end_by_origin.remove(&key);
                for cleared_key in cleared {
                    if let Some(changes) = value_changes.as_deref_mut() {
                        let before = snapshot.get_cell_value(cleared_key.sheet, cleared_key.addr);
                        snapshot.values.remove(&cleared_key);
                        let after = snapshot.get_cell_value(cleared_key.sheet, cleared_key.addr);
                        changes.record(cleared_key, before, after);
                    } else {
                        snapshot.values.remove(&cleared_key);
                    }
                    snapshot.spill_origin_by_cell.remove(&cleared_key);
                    spill_dirty_roots.push(cell_id_from_key(cleared_key));
                    self.append_blocked_spill_dirty_roots(cleared_key, spill_dirty_roots);
                }

                if let Some(changes) = value_changes.as_deref_mut() {
                    let before = snapshot.get_cell_value(key.sheet, key.addr);
                    changes.record(key, before, other.clone());
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

        if let Some(origins) = self
            .spills
            .blocked_origins_by_cell
            .get_mut(&blocked.blocker)
        {
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
                self.dirty_reasons.insert(origin, DirtyReason::Cell(cell));
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
                } else if let Some(provider) = &self.external_value_provider {
                    if let Some(sheet_name) = self.workbook.sheet_names.get(origin.sheet) {
                        if let Some(v) = provider.get(sheet_name, addr) {
                            if v != Value::Blank {
                                return Some(key);
                            }
                        }
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
        mut value_changes: Option<&mut RecalcValueChangeCollector>,
    ) {
        let top_left = array.top_left();

        if let Some(changes) = value_changes.as_deref_mut() {
            let before = snapshot.get_cell_value(origin.sheet, origin.addr);
            changes.record(origin, before, top_left.clone());
        }

        let cell = self.workbook.get_or_create_cell_mut(origin);
        cell.value = top_left.clone();
        snapshot.values.insert(origin, top_left);

        self.spills.by_origin.insert(
            origin,
            Spill {
                end,
                array: array.clone(),
            },
        );
        snapshot.spill_end_by_origin.insert(origin, end);

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
                snapshot.spill_origin_by_cell.insert(key, origin);

                if let Some(v) = array.get(r, c).cloned() {
                    if let Some(changes) = value_changes.as_deref_mut() {
                        let before = snapshot.get_cell_value(key.sheet, key.addr);
                        changes.record(key, before, v.clone());
                    }
                    snapshot.values.insert(key, v);
                }

                // Register spill cells as formula nodes that depend on the origin so they participate in
                // calculation ordering and dirty marking.
                let deps = CellDeps::new(vec![Precedent::Cell(origin_id)]);
                self.calc_graph
                    .update_cell_dependencies(cell_id_from_key(key), deps);

                spill_dirty_roots.push(cell_id_from_key(key));
            }
        }
    }

    fn try_compile_bytecode(
        &self,
        expr: &crate::Expr,
        key: CellKey,
        volatile: bool,
        thread_safe: bool,
    ) -> Option<Arc<bytecode::Program>> {
        if volatile || !thread_safe {
            return None;
        }

        let origin_ast = crate::CellAddr::new(key.addr.row, key.addr.col);
        let origin = bytecode::CellCoord {
            row: key.addr.row as i32,
            col: key.addr.col as i32,
        };
        let mut resolve_sheet = |name: &str| self.workbook.sheet_id(name);
        let expr =
            bytecode::lower_canonical_expr(expr, origin_ast, key.sheet, &mut resolve_sheet).ok()?;
        if !bytecode_expr_is_eligible(&expr) {
            return None;
        }
        if !bytecode_expr_within_grid_limits(&expr, origin) {
            return None;
        }
        Some(self.bytecode_cache.get_or_compile(&expr))
    }

    fn begin_recalc_context(&mut self) -> crate::eval::RecalcContext {
        let id = self.next_recalc_id;
        self.next_recalc_id = self.next_recalc_id.wrapping_add(1);
        crate::eval::RecalcContext::new(id)
    }

    fn compile_name_expr(&self, expr: &Expr<String>) -> CompiledExpr {
        let mut map = |sref: &SheetReference<String>| match sref {
            SheetReference::Current => SheetReference::Current,
            SheetReference::Sheet(name) => self
                .workbook
                .sheet_id(name)
                .map(SheetReference::Sheet)
                .unwrap_or_else(|| SheetReference::External(name.clone())),
            SheetReference::SheetRange(start, end) => {
                let start_id = self.workbook.sheet_id(start);
                let end_id = self.workbook.sheet_id(end);
                match (start_id, end_id) {
                    (Some(a), Some(b)) => SheetReference::SheetRange(a, b),
                    _ => SheetReference::External(format!("{start}:{end}")),
                }
            }
            SheetReference::External(wb) => SheetReference::External(wb.clone()),
        };
        expr.map_sheets(&mut map)
    }

    fn rewrite_defined_names_structural(
        &mut self,
        sheet_names: &HashMap<SheetId, String>,
        edit: &StructuralEdit,
    ) -> Result<(), EngineError> {
        let edit_sheet = match edit {
            StructuralEdit::InsertRows { sheet, .. }
            | StructuralEdit::DeleteRows { sheet, .. }
            | StructuralEdit::InsertCols { sheet, .. }
            | StructuralEdit::DeleteCols { sheet, .. } => sheet.as_str(),
        };

        let mut updates: Vec<(Option<SheetId>, String, NameDefinition, CompiledExpr)> = Vec::new();

        for (name, def) in &self.workbook.names {
            let Some((new_def, compiled)) =
                rewrite_defined_name_structural(self, def, edit_sheet, edit)?
            else {
                continue;
            };
            updates.push((None, name.clone(), new_def, compiled));
        }

        for (sheet_id, sheet) in self.workbook.sheets.iter().enumerate() {
            let Some(ctx_sheet) = sheet_names.get(&sheet_id) else {
                continue;
            };
            for (name, def) in &sheet.names {
                let Some((new_def, compiled)) =
                    rewrite_defined_name_structural(self, def, ctx_sheet, edit)?
                else {
                    continue;
                };
                updates.push((Some(sheet_id), name.clone(), new_def, compiled));
            }
        }

        for (scope_sheet, name, new_def, compiled) in updates {
            match scope_sheet {
                None => {
                    if let Some(def) = self.workbook.names.get_mut(&name) {
                        def.definition = new_def;
                        def.compiled = Some(compiled);
                    }
                }
                Some(sheet_id) => {
                    if let Some(def) = self.workbook.sheets[sheet_id].names.get_mut(&name) {
                        def.definition = new_def;
                        def.compiled = Some(compiled);
                    }
                }
            }
        }

        Ok(())
    }

    fn rewrite_defined_names_range_map(
        &mut self,
        sheet_names: &HashMap<SheetId, String>,
        edit: &RangeMapEdit,
    ) -> Result<(), EngineError> {
        let mut updates: Vec<(Option<SheetId>, String, NameDefinition, CompiledExpr)> = Vec::new();

        for (name, def) in &self.workbook.names {
            let Some((new_def, compiled)) =
                rewrite_defined_name_range_map(self, def, &edit.sheet, edit)?
            else {
                continue;
            };
            updates.push((None, name.clone(), new_def, compiled));
        }

        for (sheet_id, sheet) in self.workbook.sheets.iter().enumerate() {
            let Some(ctx_sheet) = sheet_names.get(&sheet_id) else {
                continue;
            };
            for (name, def) in &sheet.names {
                let Some((new_def, compiled)) =
                    rewrite_defined_name_range_map(self, def, ctx_sheet, edit)?
                else {
                    continue;
                };
                updates.push((Some(sheet_id), name.clone(), new_def, compiled));
            }
        }

        for (scope_sheet, name, new_def, compiled) in updates {
            match scope_sheet {
                None => {
                    if let Some(def) = self.workbook.names.get_mut(&name) {
                        def.definition = new_def;
                        def.compiled = Some(compiled);
                    }
                }
                Some(sheet_id) => {
                    if let Some(def) = self.workbook.sheets[sheet_id].names.get_mut(&name) {
                        def.definition = new_def;
                        def.compiled = Some(compiled);
                    }
                }
            }
        }

        Ok(())
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

        let tables_by_sheet: Vec<Vec<Table>> = self
            .workbook
            .sheets
            .iter()
            .map(|s| s.tables.clone())
            .collect();

        for key in cells {
            let Some(ast) = self
                .workbook
                .get_cell(key)
                .and_then(|c| c.compiled.as_ref().map(|compiled| compiled.ast().clone()))
            else {
                continue;
            };

            let cell_id = cell_id_from_key(key);

            let (names, volatile, thread_safe, dynamic_deps) =
                analyze_expr_flags(&ast, key, &tables_by_sheet, &self.workbook);
            self.set_cell_name_refs(key, names);

            let calc_precedents =
                analyze_calc_precedents(&ast, key, &tables_by_sheet, &self.workbook, &self.spills);
            let mut calc_vec: Vec<Precedent> = calc_precedents.into_iter().collect();
            calc_vec.sort_by_key(|p| match p {
                Precedent::Cell(c) => (0u8, c.sheet_id, c.cell.row, c.cell.col, 0u32, 0u32),
                Precedent::Range(r) => (
                    1u8,
                    r.sheet_id,
                    r.range.start.row,
                    r.range.start.col,
                    r.range.end.row,
                    r.range.end.col,
                ),
            });
            let deps = CellDeps::new(calc_vec).volatile(volatile);
            self.calc_graph.update_cell_dependencies(cell_id, deps);

            let cell = self.workbook.get_or_create_cell_mut(key);
            cell.volatile = volatile;
            cell.thread_safe = thread_safe;
            cell.dynamic_deps = dynamic_deps;

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
        self.dirty.contains(&CellKey {
            sheet: sheet_id,
            addr,
        })
    }

    /// Direct precedents (cells and ranges referenced by the formula in `cell`).
    pub fn precedents(&self, sheet: &str, addr: &str) -> Result<Vec<PrecedentNode>, EngineError> {
        self.precedents_impl(sheet, addr, false)
    }

    /// Transitive precedents (all precedents that can influence `cell`).
    pub fn precedents_transitive(
        &self,
        sheet: &str,
        addr: &str,
    ) -> Result<Vec<PrecedentNode>, EngineError> {
        self.precedents_impl(sheet, addr, true)
    }

    /// Returns a cell-level view of `precedents`, expanding any ranges until `limit` cells have been
    /// produced.
    ///
    /// This is intended for UI tracing/highlighting. The returned list is deterministically ordered
    /// by `(sheet, row, col)`.
    pub fn precedents_expanded(
        &self,
        sheet: &str,
        addr: &str,
        limit: usize,
    ) -> Result<Vec<(SheetId, CellAddr)>, EngineError> {
        let nodes = self.precedents(sheet, addr)?;
        Ok(expand_nodes_to_cells(&nodes, limit))
    }

    /// Direct dependents (cells whose formulas reference `cell`).
    pub fn dependents(&self, sheet: &str, addr: &str) -> Result<Vec<PrecedentNode>, EngineError> {
        self.dependents_impl(sheet, addr, false)
    }

    /// Transitive dependents (all downstream cells that are affected by `cell`).
    pub fn dependents_transitive(
        &self,
        sheet: &str,
        addr: &str,
    ) -> Result<Vec<PrecedentNode>, EngineError> {
        self.dependents_impl(sheet, addr, true)
    }

    /// Returns a cell-level view of `dependents`, expanding any range nodes until `limit` cells have
    /// been produced.
    #[must_use]
    pub fn dependents_expanded(
        &self,
        sheet: &str,
        addr: &str,
        limit: usize,
    ) -> Result<Vec<(SheetId, CellAddr)>, EngineError> {
        let nodes = self.dependents(sheet, addr)?;
        Ok(expand_nodes_to_cells(&nodes, limit))
    }

    /// Returns a dependency path explaining why `cell` is currently dirty.
    ///
    /// The returned vector is ordered from the root cause (usually an edited input cell) to the
    /// provided `cell`.
    pub fn dirty_dependency_path(&self, sheet: &str, addr: &str) -> Option<Vec<PrecedentNode>> {
        let sheet_id = self.workbook.sheet_id(sheet)?;
        let addr = parse_a1(addr).ok()?;
        let key = CellKey {
            sheet: sheet_id,
            addr,
        };
        if !self.dirty.contains(&key) {
            return None;
        }

        let mut path = vec![PrecedentNode::Cell {
            sheet: key.sheet,
            addr: key.addr,
        }];
        let mut current = key;
        let mut guard = 0usize;
        while let Some(reason) = self.dirty_reasons.get(&current).copied() {
            match reason {
                DirtyReason::Cell(prev) => {
                    path.push(PrecedentNode::Cell {
                        sheet: prev.sheet,
                        addr: prev.addr,
                    });
                    current = prev;
                }
                DirtyReason::ViaRange { from, range } => {
                    path.push(range);
                    path.push(PrecedentNode::Cell {
                        sheet: from.sheet,
                        addr: from.addr,
                    });
                    current = from;
                }
            }
            guard += 1;
            if guard > 10_000 {
                break;
            }
        }
        path.reverse();
        Some(path)
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
            return Err(EngineError::Parse(FormulaParseError::UnexpectedToken(
                format!("unknown sheet '{sheet}'"),
            )));
        };
        let addr = parse_a1(addr)?;
        let key = CellKey {
            sheet: sheet_id,
            addr,
        };
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
            SheetReference::SheetRange(start, end) => {
                let start_id = self.workbook.sheet_id(start);
                let end_id = self.workbook.sheet_id(end);
                match (start_id, end_id) {
                    (Some(a), Some(b)) => SheetReference::SheetRange(a, b),
                    _ => SheetReference::External(format!("{start}:{end}")),
                }
            }
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
    ) -> Result<Vec<PrecedentNode>, EngineError> {
        let Some(sheet_id) = self.workbook.sheet_id(sheet) else {
            return Ok(Vec::new());
        };
        let addr = parse_a1(addr)?;
        let key = CellKey {
            sheet: sheet_id,
            addr,
        };
        if transitive {
            return Ok(self.precedents_transitive_nodes(key));
        }

        let cell_id = cell_id_from_key(key);
        let mut out: Vec<PrecedentNode> = self
            .calc_graph
            .precedents_of(cell_id)
            .into_iter()
            .map(precedent_to_node)
            .collect();
        sort_and_dedup_nodes(&mut out);
        Ok(out)
    }

    fn dependents_impl(
        &self,
        sheet: &str,
        addr: &str,
        transitive: bool,
    ) -> Result<Vec<PrecedentNode>, EngineError> {
        let Some(sheet_id) = self.workbook.sheet_id(sheet) else {
            return Ok(Vec::new());
        };
        let addr = parse_a1(addr)?;
        let key = CellKey {
            sheet: sheet_id,
            addr,
        };
        if transitive {
            return Ok(self.dependents_transitive_nodes(key));
        }

        let cell_id = cell_id_from_key(key);
        let mut out: Vec<PrecedentNode> = self
            .calc_graph
            .direct_dependents(cell_id)
            .into_iter()
            .map(cell_key_from_id)
            .map(|key| PrecedentNode::Cell {
                sheet: key.sheet,
                addr: key.addr,
            })
            .collect();
        sort_and_dedup_nodes(&mut out);
        Ok(out)
    }

    fn mark_dirty_including_self_with_reasons(&mut self, from: CellKey) {
        self.dirty.insert(from);
        self.dirty_reasons.remove(&from);
        self.mark_dirty_dependents_with_reasons(from);
    }

    fn mark_dirty_dependents_with_reasons(&mut self, from: CellKey) {
        let mut queue: VecDeque<(DirtyReason, CellKey)> = VecDeque::new();

        let from_id = cell_id_from_key(from);
        for edge in self.calc_graph.dependents_of(from_id) {
            let dep = cell_key_from_id(edge.dependent);
            let reason = match edge.kind {
                crate::graph::DependentEdgeKind::DirectCell => DirtyReason::Cell(from),
                crate::graph::DependentEdgeKind::Range(range) => DirtyReason::ViaRange {
                    from,
                    range: sheet_range_to_node(range),
                },
            };
            queue.push_back((reason, dep));
        }

        while let Some((reason, cell)) = queue.pop_front() {
            if !self.dirty.insert(cell) {
                continue;
            }
            self.dirty_reasons.entry(cell).or_insert(reason);

            let cell_id = cell_id_from_key(cell);
            for edge in self.calc_graph.dependents_of(cell_id) {
                let dep = cell_key_from_id(edge.dependent);
                let next_reason = match edge.kind {
                    crate::graph::DependentEdgeKind::DirectCell => DirtyReason::Cell(cell),
                    crate::graph::DependentEdgeKind::Range(range) => DirtyReason::ViaRange {
                        from: cell,
                        range: sheet_range_to_node(range),
                    },
                };
                queue.push_back((next_reason, dep));
            }
        }
    }

    fn dependents_transitive_nodes(&self, start: CellKey) -> Vec<PrecedentNode> {
        let mut visited: HashSet<CellKey> = HashSet::new();
        let mut out: Vec<CellKey> = Vec::new();
        let mut queue: VecDeque<CellKey> = VecDeque::new();

        visited.insert(start);
        queue.push_back(start);

        while let Some(cell) = queue.pop_front() {
            let cell_id = cell_id_from_key(cell);
            for dep_id in self.calc_graph.direct_dependents(cell_id) {
                let dep = cell_key_from_id(dep_id);
                if visited.insert(dep) {
                    out.push(dep);
                    queue.push_back(dep);
                }
            }
        }

        out.sort_by_key(|k| (k.sheet, k.addr.row, k.addr.col));
        out.into_iter()
            .map(|k| PrecedentNode::Cell {
                sheet: k.sheet,
                addr: k.addr,
            })
            .collect()
    }

    fn precedents_transitive_nodes(&self, start: CellKey) -> Vec<PrecedentNode> {
        let start_node = PrecedentNode::Cell {
            sheet: start.sheet,
            addr: start.addr,
        };

        let mut visited: HashSet<PrecedentNode> = HashSet::new();
        let mut out: Vec<PrecedentNode> = Vec::new();
        let mut queue: VecDeque<PrecedentNode> = VecDeque::new();

        visited.insert(start_node);
        queue.push_back(start_node);

        while let Some(node) = queue.pop_front() {
            let neighbors: Vec<PrecedentNode> = match node {
                PrecedentNode::Cell { sheet, addr } => {
                    let key = CellKey { sheet, addr };
                    let cell_id = cell_id_from_key(key);
                    self.calc_graph
                        .precedents_of(cell_id)
                        .into_iter()
                        .map(precedent_to_node)
                        .collect()
                }
                PrecedentNode::Range { sheet, start, end } => {
                    let range = Range::new(cell_ref_from_addr(start), cell_ref_from_addr(end));
                    let sheet_range = SheetRange::new(sheet_id_for_graph(sheet), range);
                    self.calc_graph
                        .formula_cells_in_range(sheet_range)
                        .into_iter()
                        .map(|id| {
                            let key = cell_key_from_id(id);
                            PrecedentNode::Cell {
                                sheet: key.sheet,
                                addr: key.addr,
                            }
                        })
                        .collect()
                }
                PrecedentNode::SpillRange { sheet, origin, .. } => vec![PrecedentNode::Cell {
                    sheet,
                    addr: origin,
                }],
            };

            for n in neighbors {
                if visited.insert(n) {
                    out.push(n);
                    queue.push_back(n);
                }
            }
        }

        sort_and_dedup_nodes(&mut out);
        out
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
    CellAddr {
        row: cell.row,
        col: cell.col,
    }
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
        extracted.push((
            cell,
            sheet.cells.get(&cell_addr_from_cell_ref(cell)).cloned(),
        ));
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
            let origin = crate::CellAddr::new(target.row, target.col);
            let (new_formula, _) =
                rewrite_formula_for_copy_delta(formula, sheet_name, origin, delta_row, delta_col);
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
            let origin = crate::CellAddr::new(cell.row, cell.col);
            let (new_formula, _) =
                rewrite_formula_for_copy_delta(formula, sheet_name, origin, delta_row, delta_col);
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
    let mut sheet_ids: HashMap<String, SheetId> = HashMap::new();
    for (sheet_id, name) in sheet_names {
        sheet_ids.insert(name.to_ascii_lowercase(), *sheet_id);
    }

    let mut rewrites = Vec::new();
    for (sheet_id, sheet) in workbook.sheets.iter_mut().enumerate() {
        let Some(ctx_sheet) = sheet_names.get(&sheet_id) else {
            continue;
        };
        for (addr, cell) in sheet.cells.iter_mut() {
            let Some(formula) = &cell.formula else {
                continue;
            };
            let origin = crate::CellAddr::new(addr.row, addr.col);
            let (new_formula, changed) = rewrite_formula_for_structural_edit_with_resolver(
                formula,
                ctx_sheet,
                origin,
                &edit,
                |name| sheet_ids.get(&name.to_ascii_lowercase()).copied(),
            );
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
    let mut sheet_ids: HashMap<String, SheetId> = HashMap::new();
    for (sheet_id, name) in sheet_names {
        sheet_ids.insert(name.to_ascii_lowercase(), *sheet_id);
    }

    let mut rewrites = Vec::new();
    for (sheet_id, sheet) in workbook.sheets.iter_mut().enumerate() {
        let Some(ctx_sheet) = sheet_names.get(&sheet_id) else {
            continue;
        };
        for (addr, cell) in sheet.cells.iter_mut() {
            let Some(formula) = &cell.formula else {
                continue;
            };
            let origin = crate::CellAddr::new(addr.row, addr.col);
            let (new_formula, changed) = rewrite_formula_for_range_map_with_resolver(
                formula,
                ctx_sheet,
                origin,
                edit,
                |name| sheet_ids.get(&name.to_ascii_lowercase()).copied(),
            );
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

fn sheet_id_from_graph(sheet: u32) -> SheetId {
    usize::try_from(sheet).expect("sheet id exceeds usize")
}

fn sheet_range_to_node(range: SheetRange) -> PrecedentNode {
    PrecedentNode::Range {
        sheet: sheet_id_from_graph(range.sheet_id),
        start: CellAddr {
            row: range.range.start.row,
            col: range.range.start.col,
        },
        end: CellAddr {
            row: range.range.end.row,
            col: range.range.end.col,
        },
    }
}

fn precedent_to_node(precedent: Precedent) -> PrecedentNode {
    match precedent {
        Precedent::Cell(cell) => PrecedentNode::Cell {
            sheet: sheet_id_from_graph(cell.sheet_id),
            addr: CellAddr {
                row: cell.cell.row,
                col: cell.cell.col,
            },
        },
        Precedent::Range(range) => sheet_range_to_node(range),
    }
}

fn precedent_node_sort_key(node: PrecedentNode) -> (u8, SheetId, u32, u32, u32, u32, u32, u32) {
    match node {
        PrecedentNode::Cell { sheet, addr } => (0, sheet, addr.row, addr.col, 0, 0, 0, 0),
        PrecedentNode::Range { sheet, start, end } => {
            (1, sheet, start.row, start.col, end.row, end.col, 0, 0)
        }
        PrecedentNode::SpillRange {
            sheet,
            origin,
            start,
            end,
        } => (
            2, sheet, origin.row, origin.col, start.row, start.col, end.row, end.col,
        ),
    }
}

fn sort_and_dedup_nodes(nodes: &mut Vec<PrecedentNode>) {
    nodes.sort_by_key(|n| precedent_node_sort_key(*n));
    nodes.dedup();
}

fn normalize_range(start: CellAddr, end: CellAddr) -> (CellAddr, CellAddr) {
    let start_row = start.row.min(end.row);
    let end_row = start.row.max(end.row);
    let start_col = start.col.min(end.col);
    let end_col = start.col.max(end.col);
    (
        CellAddr {
            row: start_row,
            col: start_col,
        },
        CellAddr {
            row: end_row,
            col: end_col,
        },
    )
}

fn expand_nodes_to_cells(nodes: &[PrecedentNode], limit: usize) -> Vec<(SheetId, CellAddr)> {
    #[derive(Debug, Clone)]
    enum Stream {
        Single {
            sheet: SheetId,
            addr: CellAddr,
            done: bool,
        },
        Range {
            sheet: SheetId,
            start: CellAddr,
            end: CellAddr,
            cur: CellAddr,
            done: bool,
        },
    }

    impl Stream {
        fn from_node(node: PrecedentNode) -> Self {
            match node {
                PrecedentNode::Cell { sheet, addr } => Stream::Single {
                    sheet,
                    addr,
                    done: false,
                },
                PrecedentNode::Range { sheet, start, end } => {
                    let (start, end) = normalize_range(start, end);
                    Stream::Range {
                        sheet,
                        start,
                        end,
                        cur: start,
                        done: false,
                    }
                }
                PrecedentNode::SpillRange {
                    sheet, start, end, ..
                } => {
                    let (start, end) = normalize_range(start, end);
                    Stream::Range {
                        sheet,
                        start,
                        end,
                        cur: start,
                        done: false,
                    }
                }
            }
        }

        fn peek(&self) -> Option<(SheetId, CellAddr)> {
            match self {
                Stream::Single { sheet, addr, done } => (!*done).then_some((*sheet, *addr)),
                Stream::Range {
                    sheet, cur, done, ..
                } => (!*done).then_some((*sheet, *cur)),
            }
        }

        fn advance(&mut self) {
            match self {
                Stream::Single { done, .. } => *done = true,
                Stream::Range {
                    start,
                    end,
                    cur,
                    done,
                    ..
                } => {
                    if *done {
                        return;
                    }
                    if cur.row == end.row && cur.col == end.col {
                        *done = true;
                        return;
                    }
                    if cur.col < end.col {
                        cur.col += 1;
                    } else {
                        cur.col = start.col;
                        cur.row += 1;
                        if cur.row > end.row {
                            *done = true;
                        }
                    }
                }
            }
        }
    }

    if limit == 0 || nodes.is_empty() {
        return Vec::new();
    }

    let mut nodes: Vec<PrecedentNode> = nodes.to_vec();
    sort_and_dedup_nodes(&mut nodes);

    let mut streams: Vec<Stream> = nodes.into_iter().map(Stream::from_node).collect();
    let mut heap: std::collections::BinaryHeap<std::cmp::Reverse<(SheetId, u32, u32, usize)>> =
        std::collections::BinaryHeap::new();

    for (idx, stream) in streams.iter().enumerate() {
        if let Some((sheet, addr)) = stream.peek() {
            heap.push(std::cmp::Reverse((sheet, addr.row, addr.col, idx)));
        }
    }

    let mut seen: HashSet<CellKey> = HashSet::new();
    let mut out: Vec<(SheetId, CellAddr)> = Vec::new();
    out.reserve(limit.min(1024));

    while out.len() < limit {
        let Some(std::cmp::Reverse((sheet, row, col, idx))) = heap.pop() else {
            break;
        };

        let addr = CellAddr { row, col };
        if seen.insert(CellKey { sheet, addr }) {
            out.push((sheet, addr));
        }

        let stream = streams
            .get_mut(idx)
            .expect("heap indices are valid stream indices");
        stream.advance();
        if let Some((sheet, addr)) = stream.peek() {
            heap.push(std::cmp::Reverse((sheet, addr.row, addr.col, idx)));
        }
    }

    out
}

struct Snapshot {
    sheets: HashSet<SheetId>,
    sheet_names_by_id: Vec<String>,
    values: HashMap<CellKey, Value>,
    spill_end_by_origin: HashMap<CellKey, CellAddr>,
    spill_origin_by_cell: HashMap<CellKey, CellKey>,
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
                values.insert(
                    CellKey {
                        sheet: sheet_id,
                        addr: *addr,
                    },
                    cell.value.clone(),
                );
            }
        }

        // Overlay spilled values so formula evaluation can observe dynamic array results even
        // when the workbook map doesn't contain explicit cell records.
        let mut spill_end_by_origin = HashMap::new();
        for (origin, spill) in &spills.by_origin {
            spill_end_by_origin.insert(*origin, spill.end);
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
        let spill_origin_by_cell = spills.origin_by_cell.clone();
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
                NameDefinition::Reference(_) | NameDefinition::Formula(_) => {
                    crate::eval::ResolvedName::Expr(
                        def.compiled
                            .clone()
                            .expect("non-constant defined name must have compiled expression"),
                    )
                }
            }
        }

        Self {
            sheets,
            sheet_names_by_id,
            values,
            spill_end_by_origin,
            spill_origin_by_cell,
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
        if let Some(v) = self.values.get(&CellKey {
            sheet: sheet_id,
            addr,
        }) {
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

    fn sheet_id(&self, name: &str) -> Option<usize> {
        self.sheet_names_by_id
            .iter()
            .position(|candidate| candidate.eq_ignore_ascii_case(name))
    }

    fn iter_sheet_cells(&self, sheet_id: usize) -> Option<Box<dyn Iterator<Item = CellAddr> + '_>> {
        if !self.sheet_exists(sheet_id) {
            return None;
        }
        Some(Box::new(self.values.keys().filter_map(move |k| {
            if k.sheet == sheet_id {
                Some(k.addr)
            } else {
                None
            }
        })))
    }

    fn resolve_structured_ref(
        &self,
        ctx: crate::eval::EvalContext,
        sref: &crate::structured_refs::StructuredRef,
    ) -> Option<Vec<(usize, CellAddr, CellAddr)>> {
        crate::structured_refs::resolve_structured_ref(
            &self.tables,
            ctx.current_sheet,
            ctx.current_cell,
            sref,
        )
        .ok()
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

    fn spill_origin(&self, sheet_id: usize, addr: CellAddr) -> Option<CellAddr> {
        let key = CellKey {
            sheet: sheet_id,
            addr,
        };
        if self.spill_end_by_origin.contains_key(&key) {
            return Some(addr);
        }
        self.spill_origin_by_cell.get(&key).map(|k| k.addr)
    }

    fn spill_range(&self, sheet_id: usize, origin: CellAddr) -> Option<(CellAddr, CellAddr)> {
        let key = CellKey {
            sheet: sheet_id,
            addr: origin,
        };
        self.spill_end_by_origin.get(&key).map(|end| (origin, *end))
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

const EXCEL_MAX_ROWS_I32: i32 = 1_048_576;
const EXCEL_MAX_COLS_I32: i32 = 16_384;
const BYTECODE_MAX_RANGE_CELLS: i64 = 5_000_000;

fn engine_error_to_bytecode(err: ErrorKind) -> bytecode::ErrorKind {
    match err {
        ErrorKind::Null => bytecode::ErrorKind::Null,
        ErrorKind::Div0 => bytecode::ErrorKind::Div0,
        ErrorKind::Value => bytecode::ErrorKind::Value,
        ErrorKind::Ref => bytecode::ErrorKind::Ref,
        ErrorKind::Name => bytecode::ErrorKind::Name,
        ErrorKind::Num => bytecode::ErrorKind::Num,
        ErrorKind::NA => bytecode::ErrorKind::NA,
        ErrorKind::Spill => bytecode::ErrorKind::Spill,
        ErrorKind::Calc => bytecode::ErrorKind::Calc,
    }
}

fn bytecode_error_to_engine(err: bytecode::ErrorKind) -> ErrorKind {
    match err {
        bytecode::ErrorKind::Null => ErrorKind::Null,
        bytecode::ErrorKind::Div0 => ErrorKind::Div0,
        bytecode::ErrorKind::Value => ErrorKind::Value,
        bytecode::ErrorKind::Ref => ErrorKind::Ref,
        bytecode::ErrorKind::Name => ErrorKind::Name,
        bytecode::ErrorKind::Num => ErrorKind::Num,
        bytecode::ErrorKind::NA => ErrorKind::NA,
        bytecode::ErrorKind::Spill => ErrorKind::Spill,
        bytecode::ErrorKind::Calc => ErrorKind::Calc,
    }
}

fn engine_value_to_bytecode(value: &Value) -> bytecode::Value {
    match value {
        Value::Number(n) => bytecode::Value::Number(*n),
        Value::Bool(b) => bytecode::Value::Bool(*b),
        Value::Text(s) => bytecode::Value::Text(Arc::from(s.as_str())),
        Value::Blank => bytecode::Value::Empty,
        Value::Error(e) => bytecode::Value::Error(engine_error_to_bytecode(*e)),
        Value::Lambda(_) => bytecode::Value::Error(bytecode::ErrorKind::Calc),
        Value::Reference(_) | Value::ReferenceUnion(_) => {
            bytecode::Value::Error(bytecode::ErrorKind::Value)
        }
        Value::Array(_) | Value::Spill { .. } => bytecode::Value::Error(bytecode::ErrorKind::Spill),
    }
}

fn bytecode_value_to_engine(value: bytecode::Value) -> Value {
    match value {
        bytecode::Value::Number(n) => Value::Number(n),
        bytecode::Value::Bool(b) => Value::Bool(b),
        bytecode::Value::Text(s) => Value::Text(s.to_string()),
        bytecode::Value::Empty => Value::Blank,
        bytecode::Value::Error(e) => Value::Error(bytecode_error_to_engine(e)),
        bytecode::Value::Array(_) | bytecode::Value::Range(_) => Value::Error(ErrorKind::Spill),
    }
}

#[derive(Debug, Clone, Copy)]
enum ColumnSliceMode {
    /// Use column slices only when the column range contains numbers/blanks.
    ///
    /// This is required for functions like `SUMPRODUCT` that coerce logical/text values.
    StrictNumeric,
    /// Allow column slices even when the column contains logical/text values.
    ///
    /// For SUM/AVERAGE/MIN/MAX/COUNT/COUNTIF range args, Excel ignores logical/text values, so
    /// representing them as NaN (ignored by the SIMD kernels) is correct and enables SIMD even
    /// for common "header + data" columns.
    IgnoreNonNumeric,
}

fn slice_mode_for_program(program: &bytecode::Program) -> ColumnSliceMode {
    if program
        .funcs
        .iter()
        .any(|f| matches!(f, bytecode::ast::Function::SumProduct | bytecode::ast::Function::CountIf))
    {
        ColumnSliceMode::StrictNumeric
    } else {
        ColumnSliceMode::IgnoreNonNumeric
    }
}

#[derive(Debug, Clone)]
struct BytecodeColumnSegment {
    row_start: i32,
    values: Vec<f64>,
    blocked_rows_strict: Vec<i32>,
    blocked_rows_ignore_nonnumeric: Vec<i32>,
}

impl BytecodeColumnSegment {
    fn row_end(&self) -> i32 {
        self.row_start + self.values.len() as i32 - 1
    }
}

#[derive(Debug, Clone)]
struct BytecodeColumn {
    segments: Vec<BytecodeColumnSegment>,
}

#[derive(Debug)]
struct BytecodeColumnCache {
    by_sheet: Vec<HashMap<i32, BytecodeColumn>>,
}

impl BytecodeColumnCache {
    fn build(
        sheet_count: usize,
        snapshot: &Snapshot,
        tasks: &[(CellKey, CompiledFormula)],
    ) -> Self {
        // Collect row windows for each referenced column so the cache can build compact
        // columnar buffers. This avoids allocating/scanning from row 0 (e.g. `A900000:A900010`),
        // and also avoids spanning huge gaps when formulas reference multiple disjoint windows.
        let mut row_ranges_by_col: Vec<HashMap<i32, Vec<(i32, i32)>>> =
            vec![HashMap::new(); sheet_count];

        for (key, compiled) in tasks {
            let CompiledFormula::Bytecode(bc) = compiled else {
                continue;
            };

            if bc.program.range_refs.is_empty() {
                continue;
            }

            let base = bytecode::CellCoord {
                row: key.addr.row as i32,
                col: key.addr.col as i32,
            };
            for range in &bc.program.range_refs {
                let resolved = range.resolve(base);
                if resolved.row_start < 0
                    || resolved.col_start < 0
                    || resolved.row_end >= EXCEL_MAX_ROWS_I32
                    || resolved.col_end >= EXCEL_MAX_COLS_I32
                {
                    // Out-of-bounds ranges must evaluate via per-cell access so `#REF!` can be
                    // surfaced. Don't build a cache that would otherwise treat them as empty/NaN.
                    continue;
                }
                for col in resolved.col_start..=resolved.col_end {
                    row_ranges_by_col[key.sheet]
                        .entry(col)
                        .or_default()
                        .push((resolved.row_start, resolved.row_end));
                }
            }
        }

        fn apply_value(seg: &mut BytecodeColumnSegment, value: &Value, row: i32) {
            match value {
                Value::Number(n) => seg.values[(row - seg.row_start) as usize] = *n,
                Value::Blank => {}
                Value::Error(_)
                | Value::Reference(_)
                | Value::ReferenceUnion(_)
                | Value::Array(_)
                | Value::Lambda(_)
                | Value::Spill { .. } => {
                    seg.blocked_rows_strict.push(row);
                    seg.blocked_rows_ignore_nonnumeric.push(row);
                }
                Value::Bool(_) | Value::Text(_) => seg.blocked_rows_strict.push(row),
            }
        }

        let has_provider = snapshot.external_value_provider.is_some();
        let mut by_sheet: Vec<HashMap<i32, BytecodeColumn>> = Vec::with_capacity(sheet_count);
        for sheet_id in 0..sheet_count {
            let mut cols: HashMap<i32, BytecodeColumn> = HashMap::new();
            for (col, ranges) in row_ranges_by_col[sheet_id].iter() {
                let mut merged = ranges.clone();
                merged.sort_unstable();

                let mut segments: Vec<(i32, i32)> = Vec::new();
                let mut cur_start = merged[0].0;
                let mut cur_end = merged[0].1;
                for &(row_start, row_end) in &merged[1..] {
                    if row_start <= cur_end.saturating_add(1) {
                        cur_end = cur_end.max(row_end);
                    } else {
                        segments.push((cur_start, cur_end));
                        cur_start = row_start;
                        cur_end = row_end;
                    }
                }
                segments.push((cur_start, cur_end));

                let mut col_segments: Vec<BytecodeColumnSegment> =
                    Vec::with_capacity(segments.len());

                let sheet_name = snapshot.sheet_names_by_id.get(sheet_id).map(String::as_str);
                let provider = snapshot.external_value_provider.as_ref();

                for (row_start, row_end) in segments {
                    debug_assert!(row_start >= 0);
                    debug_assert!(row_end >= row_start);
                    let len = (row_end - row_start + 1) as usize;
                    let mut segment = BytecodeColumnSegment {
                        row_start,
                        values: vec![f64::NAN; len],
                        blocked_rows_strict: Vec::new(),
                        blocked_rows_ignore_nonnumeric: Vec::new(),
                    };

                    // When there is no external value provider, treat missing cells as blank and
                    // avoid scanning the entire row window just to do HashMap lookups. We'll fill
                    // from `snapshot.values` in a second pass after all segments are allocated.
                    if has_provider {
                        for row in row_start..=row_end {
                            let addr = CellAddr {
                                row: row as u32,
                                col: (*col) as u32,
                            };
                            if let Some(v) = snapshot.values.get(&CellKey {
                                sheet: sheet_id,
                                addr,
                            }) {
                                apply_value(&mut segment, v, row);
                                continue;
                            }

                            if let (Some(provider), Some(sheet_name)) = (provider, sheet_name) {
                                if let Some(v) = provider.get(sheet_name, addr) {
                                    apply_value(&mut segment, &v, row);
                                }
                            }
                        }
                    }

                    segment.blocked_rows_strict.sort_unstable();
                    segment.blocked_rows_strict.dedup();
                    segment.blocked_rows_ignore_nonnumeric.sort_unstable();
                    segment.blocked_rows_ignore_nonnumeric.dedup();

                    col_segments.push(segment);
                }

                cols.insert(
                    *col,
                    BytecodeColumn {
                        segments: col_segments,
                    },
                );
            }
            by_sheet.push(cols);
        }

        if !has_provider {
            for (key, value) in &snapshot.values {
                if matches!(value, Value::Blank) {
                    continue;
                }
                let Some(sheet_cols) = by_sheet.get_mut(key.sheet) else {
                    continue;
                };
                let col = key.addr.col as i32;
                let Some(column) = sheet_cols.get_mut(&col) else {
                    continue;
                };
                let row = key.addr.row as i32;
                let idx = column.segments.partition_point(|seg| seg.row_end() < row);
                if idx >= column.segments.len() {
                    continue;
                }
                let seg = &mut column.segments[idx];
                if row < seg.row_start || row > seg.row_end() {
                    continue;
                }
                apply_value(seg, value, row);
            }

            for sheet_cols in &mut by_sheet {
                for col in sheet_cols.values_mut() {
                    for seg in &mut col.segments {
                        seg.blocked_rows_strict.sort_unstable();
                        seg.blocked_rows_strict.dedup();
                        seg.blocked_rows_ignore_nonnumeric.sort_unstable();
                        seg.blocked_rows_ignore_nonnumeric.dedup();
                    }
                }
            }
        }

        Self { by_sheet }
    }
}

fn has_blocked_row(blocked_rows: &[i32], row_start: i32, row_end: i32) -> bool {
    if blocked_rows.is_empty() {
        return false;
    }
    let idx = blocked_rows.partition_point(|r| *r < row_start);
    blocked_rows.get(idx).is_some_and(|r| *r <= row_end)
}

struct EngineBytecodeGrid<'a> {
    snapshot: &'a Snapshot,
    sheet: SheetId,
    cols: &'a HashMap<i32, BytecodeColumn>,
    slice_mode: ColumnSliceMode,
}

impl bytecode::grid::Grid for EngineBytecodeGrid<'_> {
    fn get_value(&self, coord: bytecode::CellCoord) -> bytecode::Value {
        if coord.row < 0
            || coord.col < 0
            || coord.row >= EXCEL_MAX_ROWS_I32
            || coord.col >= EXCEL_MAX_COLS_I32
        {
            return bytecode::Value::Error(bytecode::ErrorKind::Ref);
        }
        let addr = CellAddr {
            row: coord.row as u32,
            col: coord.col as u32,
        };
        self.snapshot
            .values
            .get(&CellKey {
                sheet: self.sheet,
                addr,
            })
            .map(engine_value_to_bytecode)
            .or_else(|| {
                let provider = self.snapshot.external_value_provider.as_ref()?;
                let sheet_name = self.snapshot.sheet_names_by_id.get(self.sheet)?;
                provider
                    .get(sheet_name, addr)
                    .as_ref()
                    .map(engine_value_to_bytecode)
            })
            .unwrap_or(bytecode::Value::Empty)
    }

    fn column_slice(&self, col: i32, row_start: i32, row_end: i32) -> Option<&[f64]> {
        if col < 0
            || col >= EXCEL_MAX_COLS_I32
            || row_start < 0
            || row_end < 0
            || row_start > row_end
            || row_end >= EXCEL_MAX_ROWS_I32
        {
            return None;
        }
        let data = self.cols.get(&col)?;
        let idx = data
            .segments
            .partition_point(|seg| seg.row_end() < row_start);
        let seg = data.segments.get(idx)?;
        if row_start < seg.row_start || row_end > seg.row_end() {
            return None;
        }

        let blocked_rows = match self.slice_mode {
            ColumnSliceMode::StrictNumeric => &seg.blocked_rows_strict,
            ColumnSliceMode::IgnoreNonNumeric => &seg.blocked_rows_ignore_nonnumeric,
        };
        if has_blocked_row(blocked_rows, row_start, row_end) {
            return None;
        }

        let start = (row_start - seg.row_start) as usize;
        let end = (row_end - seg.row_start) as usize;
        if end >= seg.values.len() {
            return None;
        }
        Some(&seg.values[start..=end])
    }

    fn bounds(&self) -> (i32, i32) {
        (EXCEL_MAX_ROWS_I32, EXCEL_MAX_COLS_I32)
    }
}

fn bytecode_expr_is_eligible(expr: &bytecode::Expr) -> bool {
    bytecode_expr_is_eligible_inner(expr, false)
}

fn bytecode_expr_within_grid_limits(expr: &bytecode::Expr, origin: bytecode::CellCoord) -> bool {
    match expr {
        bytecode::Expr::Literal(_) => true,
        bytecode::Expr::CellRef(r) => {
            let coord = r.resolve(origin);
            coord.row >= 0
                && coord.col >= 0
                && coord.row < EXCEL_MAX_ROWS_I32
                && coord.col < EXCEL_MAX_COLS_I32
        }
        bytecode::Expr::RangeRef(r) => {
            let resolved = r.resolve(origin);
            if resolved.row_start < 0
                || resolved.col_start < 0
                || resolved.row_end >= EXCEL_MAX_ROWS_I32
                || resolved.col_end >= EXCEL_MAX_COLS_I32
            {
                return false;
            }
            let cells = (resolved.rows() as i64) * (resolved.cols() as i64);
            cells <= BYTECODE_MAX_RANGE_CELLS
        }
        bytecode::Expr::Unary { expr, .. } => bytecode_expr_within_grid_limits(expr, origin),
        bytecode::Expr::Binary { left, right, .. } => {
            bytecode_expr_within_grid_limits(left, origin)
                && bytecode_expr_within_grid_limits(right, origin)
        }
        bytecode::Expr::FuncCall { args, .. } => args
            .iter()
            .all(|arg| bytecode_expr_within_grid_limits(arg, origin)),
    }
}

fn bytecode_expr_is_eligible_inner(expr: &bytecode::Expr, allow_range: bool) -> bool {
    fn parses_numeric_criteria_literal(raw: &str) -> bool {
        let raw = raw.trim();
        let rest = if let Some(r) = raw.strip_prefix(">=") {
            r
        } else if let Some(r) = raw.strip_prefix("<=") {
            r
        } else if let Some(r) = raw.strip_prefix("<>") {
            r
        } else if let Some(r) = raw.strip_prefix('>') {
            r
        } else if let Some(r) = raw.strip_prefix('<') {
            r
        } else if let Some(r) = raw.strip_prefix('=') {
            r
        } else {
            raw
        };
        let rhs = rest.trim();
        if rhs.is_empty() {
            return false;
        }
        rhs.parse::<f64>().is_ok()
    }

    match expr {
        bytecode::Expr::Literal(v) => match v {
            bytecode::Value::Number(_) | bytecode::Value::Bool(_) => true,
            bytecode::Value::Text(_) => true,
            bytecode::Value::Empty => true,
            bytecode::Value::Error(_) | bytecode::Value::Array(_) | bytecode::Value::Range(_) => {
                false
            }
        },
        bytecode::Expr::CellRef(_) => true,
        bytecode::Expr::RangeRef(_) => allow_range,
        bytecode::Expr::Unary { expr, .. } => bytecode_expr_is_eligible_inner(expr, false),
        bytecode::Expr::Binary { op, left, right } => {
            matches!(
                op,
                bytecode::ast::BinaryOp::Add
                    | bytecode::ast::BinaryOp::Sub
                    | bytecode::ast::BinaryOp::Mul
                    | bytecode::ast::BinaryOp::Div
                    | bytecode::ast::BinaryOp::Pow
                    | bytecode::ast::BinaryOp::Eq
                    | bytecode::ast::BinaryOp::Ne
                    | bytecode::ast::BinaryOp::Lt
                    | bytecode::ast::BinaryOp::Le
                    | bytecode::ast::BinaryOp::Gt
                    | bytecode::ast::BinaryOp::Ge
            ) && bytecode_expr_is_eligible_inner(left, false)
                && bytecode_expr_is_eligible_inner(right, false)
        }
        bytecode::Expr::FuncCall { func, args } => match func {
            bytecode::ast::Function::Sum
            | bytecode::ast::Function::Average
            | bytecode::ast::Function::Min
            | bytecode::ast::Function::Max
            | bytecode::ast::Function::Count => args
                .iter()
                .all(|arg| bytecode_expr_is_eligible_inner(arg, true)),
            bytecode::ast::Function::CountIf => {
                if args.len() != 2 {
                    return false;
                }
                let range_ok = matches!(args[0], bytecode::Expr::RangeRef(_) | bytecode::Expr::CellRef(_));

                // The bytecode runtime currently supports numeric COUNTIF criteria only. Reject
                // unsupported criteria shapes here so the engine falls back to the full evaluator
                // (which implements Excel-style wildcards, blanks, errors, date parsing, etc).
                let criteria_ok = match &args[1] {
                    bytecode::Expr::Literal(bytecode::Value::Number(_))
                    | bytecode::Expr::Literal(bytecode::Value::Bool(_)) => true,
                    bytecode::Expr::Literal(bytecode::Value::Text(s)) => {
                        parses_numeric_criteria_literal(s.as_ref())
                    }
                    _ => false,
                };

                range_ok && criteria_ok
            }
            bytecode::ast::Function::SumProduct => {
                if args.len() != 2 {
                    return false;
                }
                (matches!(args[0], bytecode::Expr::RangeRef(_))
                    || matches!(args[0], bytecode::Expr::CellRef(_)))
                    && (matches!(args[1], bytecode::Expr::RangeRef(_))
                        || matches!(args[1], bytecode::Expr::CellRef(_)))
            }
            bytecode::ast::Function::Abs
            | bytecode::ast::Function::Int
            | bytecode::ast::Function::Round
            | bytecode::ast::Function::RoundUp
            | bytecode::ast::Function::RoundDown
            | bytecode::ast::Function::Mod
            | bytecode::ast::Function::Sign
            | bytecode::ast::Function::Concat => args
                .iter()
                .all(|arg| bytecode_expr_is_eligible_inner(arg, false)),
            bytecode::ast::Function::Unknown(_) => false,
        },
    }
}

fn analyze_expr_flags(
    expr: &CompiledExpr,
    current_cell: CellKey,
    _tables_by_sheet: &[Vec<Table>],
    workbook: &Workbook,
) -> (HashSet<String>, bool, bool, bool) {
    let mut names = HashSet::new();
    let mut volatile = false;
    let mut thread_safe = true;
    let mut dynamic_deps = false;
    let mut visiting_names = HashSet::new();
    let mut lexical_scopes: Vec<HashSet<String>> = Vec::new();
    walk_expr_flags(
        expr,
        current_cell,
        workbook,
        &mut names,
        &mut volatile,
        &mut thread_safe,
        &mut dynamic_deps,
        &mut visiting_names,
        &mut lexical_scopes,
    );
    (names, volatile, thread_safe, dynamic_deps)
}

fn walk_expr_flags(
    expr: &CompiledExpr,
    current_cell: CellKey,
    workbook: &Workbook,
    names: &mut HashSet<String>,
    volatile: &mut bool,
    thread_safe: &mut bool,
    dynamic_deps: &mut bool,
    visiting_names: &mut HashSet<(SheetId, String)>,
    lexical_scopes: &mut Vec<HashSet<String>>,
) {
    fn name_is_local(scopes: &[HashSet<String>], name_key: &str) -> bool {
        scopes.iter().rev().any(|scope| scope.contains(name_key))
    }

    fn bare_identifier(expr: &CompiledExpr) -> Option<String> {
        match expr {
            Expr::NameRef(nref) if matches!(nref.sheet, SheetReference::Current) => {
                Some(normalize_defined_name(&nref.name))
            }
            _ => None,
        }
    }

    match expr {
        Expr::NameRef(nref) => {
            let Some(sheet) = resolve_single_sheet(&nref.sheet, current_cell.sheet) else {
                return;
            };
            let name_key = normalize_defined_name(&nref.name);
            if name_key.is_empty() {
                return;
            }

            if name_is_local(lexical_scopes, &name_key) {
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
                    walk_expr_flags(
                        expr,
                        CellKey {
                            sheet,
                            addr: current_cell.addr,
                        },
                        workbook,
                        names,
                        volatile,
                        thread_safe,
                        dynamic_deps,
                        visiting_names,
                        lexical_scopes,
                    );
                }
            }

            visiting_names.remove(&visit_key);
        }
        Expr::Unary { expr, .. } | Expr::Postfix { expr, .. } => {
            walk_expr_flags(
                expr,
                current_cell,
                workbook,
                names,
                volatile,
                thread_safe,
                dynamic_deps,
                visiting_names,
                lexical_scopes,
            );
        }
        Expr::Binary { left, right, .. } | Expr::Compare { left, right, .. } => {
            walk_expr_flags(
                left,
                current_cell,
                workbook,
                names,
                volatile,
                thread_safe,
                dynamic_deps,
                visiting_names,
                lexical_scopes,
            );
            walk_expr_flags(
                right,
                current_cell,
                workbook,
                names,
                volatile,
                thread_safe,
                dynamic_deps,
                visiting_names,
                lexical_scopes,
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
                if matches!(spec.name, "OFFSET" | "INDIRECT") {
                    *dynamic_deps = true;
                }

                match spec.name {
                    "LET" => {
                        if args.len() < 3 || args.len() % 2 == 0 {
                            return;
                        }

                        lexical_scopes.push(HashSet::new());
                        for pair in args[..args.len() - 1].chunks_exact(2) {
                            let Some(name_key) = bare_identifier(&pair[0]) else {
                                lexical_scopes.pop();
                                return;
                            };

                            walk_expr_flags(
                                &pair[1],
                                current_cell,
                                workbook,
                                names,
                                volatile,
                                thread_safe,
                                dynamic_deps,
                                visiting_names,
                                lexical_scopes,
                            );
                            lexical_scopes
                                .last_mut()
                                .expect("pushed scope")
                                .insert(name_key);
                        }

                        walk_expr_flags(
                            &args[args.len() - 1],
                            current_cell,
                            workbook,
                            names,
                            volatile,
                            thread_safe,
                            dynamic_deps,
                            visiting_names,
                            lexical_scopes,
                        );
                        lexical_scopes.pop();
                        return;
                    }
                    "LAMBDA" => {
                        if args.is_empty() {
                            return;
                        }

                        let mut scope = HashSet::new();
                        for param in &args[..args.len() - 1] {
                            let Some(name_key) = bare_identifier(param) else {
                                return;
                            };
                            if !scope.insert(name_key) {
                                return;
                            }
                        }

                        lexical_scopes.push(scope);
                        walk_expr_flags(
                            &args[args.len() - 1],
                            current_cell,
                            workbook,
                            names,
                            volatile,
                            thread_safe,
                            dynamic_deps,
                            visiting_names,
                            lexical_scopes,
                        );
                        lexical_scopes.pop();
                        return;
                    }
                    _ => {}
                }
            } else {
                let name_key = normalize_defined_name(name);
                let is_local = !name_key.is_empty() && name_is_local(lexical_scopes, &name_key);
                let mut resolved_defined_name = None;
                if !name_key.is_empty() && !is_local {
                    names.insert(name_key.clone());

                    let sheet = current_cell.sheet;
                    let visit_key = (sheet, name_key.clone());
                    if visiting_names.insert(visit_key.clone()) {
                        resolved_defined_name = resolve_defined_name(workbook, sheet, &name_key);
                        if let Some(def) = resolved_defined_name.as_ref() {
                            if let Some(expr) = def.compiled.as_ref() {
                                walk_expr_flags(
                                    expr,
                                    CellKey {
                                        sheet,
                                        addr: current_cell.addr,
                                    },
                                    workbook,
                                    names,
                                    volatile,
                                    thread_safe,
                                    dynamic_deps,
                                    visiting_names,
                                    lexical_scopes,
                                );
                            }
                        }
                        visiting_names.remove(&visit_key);
                    }
                }

                // Placeholder: treat unresolved UDFs as non-thread-safe.
                if !is_local && resolved_defined_name.is_none() {
                    *thread_safe = false;
                }
            }
            for a in args {
                walk_expr_flags(
                    a,
                    current_cell,
                    workbook,
                    names,
                    volatile,
                    thread_safe,
                    dynamic_deps,
                    visiting_names,
                    lexical_scopes,
                );
            }
        }
        Expr::ArrayLiteral { values, .. } => {
            for el in values.iter() {
                walk_expr_flags(
                    el,
                    current_cell,
                    workbook,
                    names,
                    volatile,
                    thread_safe,
                    dynamic_deps,
                    visiting_names,
                    lexical_scopes,
                );
            }
        }
        Expr::ImplicitIntersection(inner) | Expr::SpillRange(inner) => {
            walk_expr_flags(
                inner,
                current_cell,
                workbook,
                names,
                volatile,
                thread_safe,
                dynamic_deps,
                visiting_names,
                lexical_scopes,
            );
        }
        Expr::CellRef(_)
        | Expr::RangeRef(_)
        | Expr::StructuredRef(_)
        | Expr::Number(_)
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
    spills: &SpillState,
) -> HashSet<Precedent> {
    let mut out = HashSet::new();
    let mut visiting_names = HashSet::new();
    let mut lexical_scopes: Vec<HashSet<String>> = Vec::new();
    walk_calc_expr(
        expr,
        current_cell,
        tables_by_sheet,
        workbook,
        spills,
        &mut out,
        &mut visiting_names,
        &mut lexical_scopes,
    );
    out
}

fn spill_range_target_cell(expr: &CompiledExpr, current_cell: CellKey) -> Option<CellKey> {
    match expr {
        Expr::CellRef(r) => {
            resolve_single_sheet(&r.sheet, current_cell.sheet).map(|sheet| CellKey {
                sheet,
                addr: r.addr,
            })
        }
        Expr::ImplicitIntersection(inner) | Expr::SpillRange(inner) => {
            spill_range_target_cell(inner, current_cell)
        }
        _ => None,
    }
}

fn spill_range_bounds(cell: CellKey, spills: &SpillState) -> (CellKey, CellAddr) {
    let origin = if spills.by_origin.contains_key(&cell) {
        cell
    } else {
        spills.origin_by_cell.get(&cell).copied().unwrap_or(cell)
    };
    let end = spills
        .by_origin
        .get(&origin)
        .map(|spill| spill.end)
        .unwrap_or(origin.addr);
    (origin, end)
}

fn walk_calc_expr(
    expr: &CompiledExpr,
    current_cell: CellKey,
    tables_by_sheet: &[Vec<Table>],
    workbook: &Workbook,
    spills: &SpillState,
    precedents: &mut HashSet<Precedent>,
    visiting_names: &mut HashSet<(SheetId, String)>,
    lexical_scopes: &mut Vec<HashSet<String>>,
) {
    fn name_is_local(scopes: &[HashSet<String>], name_key: &str) -> bool {
        scopes.iter().rev().any(|scope| scope.contains(name_key))
    }

    fn bare_identifier(expr: &CompiledExpr) -> Option<String> {
        match expr {
            Expr::NameRef(nref) if matches!(nref.sheet, SheetReference::Current) => {
                Some(normalize_defined_name(&nref.name))
            }
            _ => None,
        }
    }

    match expr {
        Expr::CellRef(r) => {
            if let Some(sheets) = resolve_sheet_span(&r.sheet, current_cell.sheet, workbook) {
                for sheet in sheets {
                    precedents.insert(Precedent::Cell(CellId::new(
                        sheet_id_for_graph(sheet),
                        r.addr.row,
                        r.addr.col,
                    )));
                }
            }
        }
        Expr::RangeRef(RangeRef { sheet, start, end }) => {
            if let Some(sheets) = resolve_sheet_span(sheet, current_cell.sheet, workbook) {
                let range = Range::new(
                    CellRef::new(start.row, start.col),
                    CellRef::new(end.row, end.col),
                );
                for sheet_id in sheets {
                    precedents.insert(Precedent::Range(SheetRange::new(
                        sheet_id_for_graph(sheet_id),
                        range,
                    )));
                }
            }
        }
        Expr::StructuredRef(sref) => {
            if let Ok(ranges) = crate::structured_refs::resolve_structured_ref(
                tables_by_sheet,
                current_cell.sheet,
                current_cell.addr,
                sref,
            ) {
                for (sheet_id, start, end) in ranges {
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
        }
        Expr::NameRef(nref) => {
            let Some(sheet) = resolve_single_sheet(&nref.sheet, current_cell.sheet) else {
                return;
            };
            let name_key = normalize_defined_name(&nref.name);
            if name_key.is_empty() {
                return;
            }
            if name_is_local(lexical_scopes, &name_key) {
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
                        spills,
                        precedents,
                        visiting_names,
                        lexical_scopes,
                    );
                }
            }
            visiting_names.remove(&visit_key);
        }
        Expr::SpillRange(inner) => {
            if let Some(target) = spill_range_target_cell(inner, current_cell) {
                let (origin, end) = spill_range_bounds(target, spills);
                if origin.addr == end {
                    precedents.insert(Precedent::Cell(CellId::new(
                        sheet_id_for_graph(origin.sheet),
                        origin.addr.row,
                        origin.addr.col,
                    )));
                } else {
                    let range = Range::new(
                        CellRef::new(origin.addr.row, origin.addr.col),
                        CellRef::new(end.row, end.col),
                    );
                    precedents.insert(Precedent::Range(SheetRange::new(
                        sheet_id_for_graph(origin.sheet),
                        range,
                    )));
                }
            }
            walk_calc_expr(
                inner,
                current_cell,
                tables_by_sheet,
                workbook,
                spills,
                precedents,
                visiting_names,
                lexical_scopes,
            );
        }
        Expr::Unary { expr, .. } | Expr::Postfix { expr, .. } => {
            walk_calc_expr(
                expr,
                current_cell,
                tables_by_sheet,
                workbook,
                spills,
                precedents,
                visiting_names,
                lexical_scopes,
            );
        }
        Expr::Binary { left, right, .. } | Expr::Compare { left, right, .. } => {
            walk_calc_expr(
                left,
                current_cell,
                tables_by_sheet,
                workbook,
                spills,
                precedents,
                visiting_names,
                lexical_scopes,
            );
            walk_calc_expr(
                right,
                current_cell,
                tables_by_sheet,
                workbook,
                spills,
                precedents,
                visiting_names,
                lexical_scopes,
            );
        }
        Expr::FunctionCall { name, args, .. } => {
            if let Some(spec) = crate::functions::lookup_function(name) {
                match spec.name {
                    "LET" => {
                        if args.len() < 3 || args.len() % 2 == 0 {
                            return;
                        }

                        lexical_scopes.push(HashSet::new());
                        for pair in args[..args.len() - 1].chunks_exact(2) {
                            let Some(name_key) = bare_identifier(&pair[0]) else {
                                lexical_scopes.pop();
                                return;
                            };
                            walk_calc_expr(
                                &pair[1],
                                current_cell,
                                tables_by_sheet,
                                workbook,
                                spills,
                                precedents,
                                visiting_names,
                                lexical_scopes,
                            );
                            lexical_scopes
                                .last_mut()
                                .expect("pushed scope")
                                .insert(name_key);
                        }

                        walk_calc_expr(
                            &args[args.len() - 1],
                            current_cell,
                            tables_by_sheet,
                            workbook,
                            spills,
                            precedents,
                            visiting_names,
                            lexical_scopes,
                        );
                        lexical_scopes.pop();
                        return;
                    }
                    "LAMBDA" => {
                        if args.is_empty() {
                            return;
                        }

                        let mut scope = HashSet::new();
                        for param in &args[..args.len() - 1] {
                            let Some(name_key) = bare_identifier(param) else {
                                return;
                            };
                            if !scope.insert(name_key) {
                                return;
                            }
                        }

                        lexical_scopes.push(scope);
                        walk_calc_expr(
                            &args[args.len() - 1],
                            current_cell,
                            tables_by_sheet,
                            workbook,
                            spills,
                            precedents,
                            visiting_names,
                            lexical_scopes,
                        );
                        lexical_scopes.pop();
                        return;
                    }
                    _ => {}
                }
            } else {
                let name_key = normalize_defined_name(name);
                if !name_key.is_empty() && !name_is_local(lexical_scopes, &name_key) {
                    let sheet = current_cell.sheet;
                    let visit_key = (sheet, name_key.clone());
                    if visiting_names.insert(visit_key.clone()) {
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
                                    spills,
                                    precedents,
                                    visiting_names,
                                    lexical_scopes,
                                );
                            }
                        }
                        visiting_names.remove(&visit_key);
                    }
                }
            }

            for a in args {
                walk_calc_expr(
                    a,
                    current_cell,
                    tables_by_sheet,
                    workbook,
                    spills,
                    precedents,
                    visiting_names,
                    lexical_scopes,
                );
            }
        }
        Expr::ArrayLiteral { values, .. } => {
            for el in values.iter() {
                walk_calc_expr(
                    el,
                    current_cell,
                    tables_by_sheet,
                    workbook,
                    spills,
                    precedents,
                    visiting_names,
                    lexical_scopes,
                );
            }
        }
        Expr::ImplicitIntersection(inner) => {
            walk_calc_expr(
                inner,
                current_cell,
                tables_by_sheet,
                workbook,
                spills,
                precedents,
                visiting_names,
                lexical_scopes,
            );
        }
        Expr::Number(_) | Expr::Text(_) | Expr::Bool(_) | Expr::Blank | Expr::Error(_) => {}
    }
}

fn resolve_sheet(sheet: &SheetReference<usize>, current_sheet: SheetId) -> Option<SheetId> {
    match sheet {
        SheetReference::Current => Some(current_sheet),
        SheetReference::Sheet(id) => Some(*id),
        SheetReference::SheetRange(a, b) => {
            if a == b {
                Some(*a)
            } else {
                None
            }
        }
        SheetReference::External(_) => None,
    }
}

fn resolve_single_sheet(sheet: &SheetReference<usize>, current_sheet: SheetId) -> Option<SheetId> {
    resolve_sheet(sheet, current_sheet)
}

fn resolve_sheet_span(
    sheet: &SheetReference<usize>,
    current_sheet: SheetId,
    workbook: &Workbook,
) -> Option<std::ops::RangeInclusive<SheetId>> {
    match sheet {
        SheetReference::Current => Some(current_sheet..=current_sheet),
        SheetReference::Sheet(id) => Some(*id..=*id),
        SheetReference::SheetRange(a, b) => {
            let (start, end) = if a <= b { (*a, *b) } else { (*b, *a) };
            if end >= workbook.sheets.len() {
                return None;
            }
            Some(start..=end)
        }
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
        Value::Text(_)
        | Value::Error(_)
        | Value::Reference(_)
        | Value::ReferenceUnion(_)
        | Value::Array(_)
        | Value::Lambda(_)
        | Value::Spill { .. } => None,
    }
}

fn normalize_defined_name(name: &str) -> String {
    name.trim().to_ascii_uppercase()
}

fn rewrite_defined_name_structural(
    engine: &Engine,
    def: &DefinedName,
    ctx_sheet: &str,
    edit: &StructuralEdit,
) -> Result<Option<(NameDefinition, CompiledExpr)>, EngineError> {
    let origin = crate::CellAddr::new(0, 0);
    let (new_def, changed) = match &def.definition {
        NameDefinition::Constant(_) => return Ok(None),
        NameDefinition::Reference(formula) => {
            let (new_formula, changed) = rewrite_formula_for_structural_edit_with_resolver(
                formula,
                ctx_sheet,
                origin,
                edit,
                |name| engine.workbook.sheet_id(name),
            );
            (NameDefinition::Reference(new_formula), changed)
        }
        NameDefinition::Formula(formula) => {
            let (new_formula, changed) = rewrite_formula_for_structural_edit_with_resolver(
                formula,
                ctx_sheet,
                origin,
                edit,
                |name| engine.workbook.sheet_id(name),
            );
            (NameDefinition::Formula(new_formula), changed)
        }
    };

    if !changed {
        return Ok(None);
    }

    let formula = match &new_def {
        NameDefinition::Reference(f) | NameDefinition::Formula(f) => f,
        NameDefinition::Constant(_) => unreachable!("handled above"),
    };
    let ast = crate::parse_formula(
        formula,
        crate::ParseOptions {
            locale: crate::LocaleConfig::en_us(),
            reference_style: crate::ReferenceStyle::A1,
            normalize_relative_to: None,
        },
    )?;
    let parsed = lower_ast(&ast, None);
    let compiled = engine.compile_name_expr(&parsed);
    Ok(Some((new_def, compiled)))
}

fn rewrite_defined_name_range_map(
    engine: &Engine,
    def: &DefinedName,
    ctx_sheet: &str,
    edit: &RangeMapEdit,
) -> Result<Option<(NameDefinition, CompiledExpr)>, EngineError> {
    let origin = crate::CellAddr::new(0, 0);
    let (new_def, changed) = match &def.definition {
        NameDefinition::Constant(_) => return Ok(None),
        NameDefinition::Reference(formula) => {
            let (new_formula, changed) = rewrite_formula_for_range_map_with_resolver(
                formula,
                ctx_sheet,
                origin,
                edit,
                |name| engine.workbook.sheet_id(name),
            );
            (NameDefinition::Reference(new_formula), changed)
        }
        NameDefinition::Formula(formula) => {
            let (new_formula, changed) = rewrite_formula_for_range_map_with_resolver(
                formula,
                ctx_sheet,
                origin,
                edit,
                |name| engine.workbook.sheet_id(name),
            );
            (NameDefinition::Formula(new_formula), changed)
        }
    };

    if !changed {
        return Ok(None);
    }

    let formula = match &new_def {
        NameDefinition::Reference(f) | NameDefinition::Formula(f) => f,
        NameDefinition::Constant(_) => unreachable!("handled above"),
    };
    let ast = crate::parse_formula(
        formula,
        crate::ParseOptions {
            locale: crate::LocaleConfig::en_us(),
            reference_style: crate::ReferenceStyle::A1,
            normalize_relative_to: None,
        },
    )?;
    let parsed = lower_ast(&ast, None);
    let compiled = engine.compile_name_expr(&parsed);
    Ok(Some((new_def, compiled)))
}

#[cfg(test)]
mod tests {
    use super::*;
    use chrono::TimeZone;

    #[test]
    fn let_lambda_calls_are_thread_safe() {
        let mut engine = Engine::new();
        engine
            .set_cell_formula("Sheet1", "A1", "=LET(f,LAMBDA(x,x+1),f(2))")
            .unwrap();

        let sheet_id = engine.workbook.sheet_id("Sheet1").expect("sheet exists");
        let addr = parse_a1("A1").unwrap();
        let cell = engine.workbook.sheets[sheet_id]
            .cells
            .get(&addr)
            .expect("cell stored");
        assert!(
            cell.thread_safe,
            "LET/LAMBDA should be safe for parallel evaluation"
        );
    }

    #[test]
    fn lambda_parameter_calls_are_thread_safe() {
        let mut engine = Engine::new();
        engine
            .set_cell_formula(
                "Sheet1",
                "A1",
                "=LET(apply,LAMBDA(f,f(1)),apply(LAMBDA(x,x+1)))",
            )
            .unwrap();

        let sheet_id = engine.workbook.sheet_id("Sheet1").expect("sheet exists");
        let addr = parse_a1("A1").unwrap();
        let cell = engine.workbook.sheets[sheet_id]
            .cells
            .get(&addr)
            .expect("cell stored");
        assert!(
            cell.thread_safe,
            "higher-order lambda calls should be parallel-safe"
        );

        engine.recalculate_single_threaded();
        assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(2.0));
    }

    #[test]
    fn multithreaded_and_singlethreaded_match_for_volatiles_given_same_recalc_context() {
        fn setup(engine: &mut Engine) {
            engine
                .set_cell_formula("Sheet1", "A1", "=NOW()")
                .expect("set NOW()");
            engine
                .set_cell_formula("Sheet1", "A2", "=RAND()")
                .expect("set RAND()");
            engine
                .set_cell_formula("Sheet1", "A3", "=RANDBETWEEN(10, 20)")
                .expect("set RANDBETWEEN()");
            engine
                .set_cell_formula("Sheet1", "B1", "=A1+A2+A3")
                .expect("set dependent");
        }

        let mut single = Engine::new();
        setup(&mut single);
        let mut multi = Engine::new();
        setup(&mut multi);

        let recalc_ctx = crate::eval::RecalcContext {
            now_utc: chrono::Utc
                .timestamp_opt(1_700_000_000, 123_456_789)
                .single()
                .unwrap(),
            recalc_id: 42,
            number_locale: crate::value::NumberLocale::en_us(),
        };

        let levels_single = single
            .calc_graph
            .calc_levels_for_dirty()
            .expect("calc levels");
        let _ = single.recalculate_levels(
            levels_single,
            RecalcMode::SingleThreaded,
            &recalc_ctx,
            single.date_system,
            None,
        );

        let levels_multi = multi
            .calc_graph
            .calc_levels_for_dirty()
            .expect("calc levels");
        let _ = multi.recalculate_levels(
            levels_multi,
            RecalcMode::MultiThreaded,
            &recalc_ctx,
            multi.date_system,
            None,
        );

        for addr in ["A1", "A2", "A3", "B1"] {
            assert_eq!(
                multi.get_cell_value("Sheet1", addr),
                single.get_cell_value("Sheet1", addr),
                "mismatch at {addr}"
            );
        }
    }

    #[test]
    fn recalculate_with_value_changes_includes_spill_outputs() {
        let mut engine = Engine::new();
        engine
            .set_cell_formula("Sheet1", "A1", "=SEQUENCE(1,2)")
            .unwrap();

        let changes = engine.recalculate_with_value_changes(RecalcMode::SingleThreaded);
        assert_eq!(
            changes,
            vec![
                RecalcValueChange {
                    sheet: "Sheet1".to_string(),
                    addr: parse_a1("A1").unwrap(),
                    value: Value::Number(1.0),
                },
                RecalcValueChange {
                    sheet: "Sheet1".to_string(),
                    addr: parse_a1("B1").unwrap(),
                    value: Value::Number(2.0),
                },
            ]
        );
    }

    #[test]
    fn clear_cell_removes_literal_entry_from_sheet_map() {
        let mut engine = Engine::new();
        engine.set_cell_value("Sheet1", "A1", 123.0).unwrap();

        let sheet_id = engine.workbook.sheet_id("Sheet1").expect("sheet exists");
        let addr = parse_a1("A1").unwrap();
        assert!(
            engine.workbook.sheets[sheet_id].cells.contains_key(&addr),
            "literal cell should be stored"
        );

        engine.clear_cell("Sheet1", "A1").unwrap();

        assert!(
            !engine.workbook.sheets[sheet_id].cells.contains_key(&addr),
            "cleared literal cell should be removed from sparse storage"
        );
        assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Blank);
    }

    #[test]
    fn clear_cell_marks_dependents_dirty() {
        let mut engine = Engine::new();
        engine.set_cell_value("Sheet1", "A1", 2.0).unwrap();
        engine.set_cell_formula("Sheet1", "B1", "=A1*2").unwrap();
        engine.recalculate();
        assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(4.0));

        engine.clear_cell("Sheet1", "A1").unwrap();
        engine.recalculate();
        assert_eq!(
            engine.get_cell_value("Sheet1", "B1"),
            Value::Number(0.0),
            "clearing an input should propagate to dependent formulas"
        );
    }

    #[test]
    fn recalculate_with_value_changes_tracks_scalar_formula_updates() {
        let mut engine = Engine::new();
        engine.set_cell_value("Sheet1", "A2", 3.0).unwrap();
        engine.set_cell_formula("Sheet1", "A1", "=A2*2").unwrap();

        let changes = engine.recalculate_with_value_changes(RecalcMode::SingleThreaded);
        assert_eq!(
            changes,
            vec![RecalcValueChange {
                sheet: "Sheet1".to_string(),
                addr: parse_a1("A1").unwrap(),
                value: Value::Number(6.0),
            }]
        );

        engine.set_cell_value("Sheet1", "A2", 4.0).unwrap();
        let changes = engine.recalculate_with_value_changes(RecalcMode::SingleThreaded);
        assert_eq!(
            changes,
            vec![RecalcValueChange {
                sheet: "Sheet1".to_string(),
                addr: parse_a1("A1").unwrap(),
                value: Value::Number(8.0),
            }]
        );
    }

    #[test]
    fn recalculate_with_value_changes_clears_shrunk_spill_cells() {
        let mut engine = Engine::new();
        engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
        engine
            .set_cell_formula("Sheet1", "A1", "=SEQUENCE(1,A2)")
            .unwrap();

        let _ = engine.recalculate_with_value_changes(RecalcMode::SingleThreaded);

        // Shrink the spill width from 2 columns to 1. The previous spill cell (B1)
        // should be returned as a delta back to blank.
        engine.set_cell_value("Sheet1", "A2", 1.0).unwrap();
        let changes = engine.recalculate_with_value_changes(RecalcMode::SingleThreaded);
        assert_eq!(
            changes,
            vec![RecalcValueChange {
                sheet: "Sheet1".to_string(),
                addr: parse_a1("B1").unwrap(),
                value: Value::Blank,
            }]
        );
    }

    #[test]
    fn clear_cell_removes_formula_entry_from_sheet_map() {
        let mut engine = Engine::new();
        engine.set_cell_formula("Sheet1", "A1", "=1+1").unwrap();

        let sheet_id = engine.workbook.sheet_id("Sheet1").expect("sheet exists");
        let addr = parse_a1("A1").unwrap();
        assert!(
            engine.workbook.sheets[sheet_id].cells.contains_key(&addr),
            "formula cell should be stored"
        );
        assert_eq!(engine.get_cell_formula("Sheet1", "A1"), Some("=1+1"));

        engine.clear_cell("Sheet1", "A1").unwrap();

        assert!(
            !engine.workbook.sheets[sheet_id].cells.contains_key(&addr),
            "cleared formula cell should be removed from sparse storage"
        );
        assert_eq!(engine.get_cell_formula("Sheet1", "A1"), None);
        assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Blank);
    }

    #[test]
    fn bytecode_column_cache_uses_range_min_row() {
        let mut engine = Engine::new();
        engine
            .set_cell_formula("Sheet1", "B1", "=SUM(A900000:A900010)")
            .unwrap();

        let sheet_id = engine.workbook.sheet_id("Sheet1").expect("sheet exists");
        let key = CellKey {
            sheet: sheet_id,
            addr: parse_a1("B1").unwrap(),
        };
        let compiled = engine
            .workbook
            .get_cell(key)
            .and_then(|c| c.compiled.clone())
            .expect("compiled formula stored");

        let snapshot = Snapshot::from_workbook(
            &engine.workbook,
            &engine.spills,
            engine.external_value_provider.clone(),
        );
        let column_cache =
            BytecodeColumnCache::build(engine.workbook.sheets.len(), &snapshot, &[(key, compiled)]);

        let col = column_cache.by_sheet[sheet_id]
            .get(&0)
            .expect("column A is cached");
        assert_eq!(col.segments.len(), 1);
        let seg = &col.segments[0];
        assert_eq!(seg.row_start, 899_999);
        assert_eq!(seg.values.len(), 11);
    }

    #[test]
    fn bytecode_column_cache_builds_disjoint_segments() {
        let mut engine = Engine::new();
        engine
            .set_cell_formula("Sheet1", "B1", "=SUM(A1:A10)+SUM(A900000:A900010)")
            .unwrap();

        let sheet_id = engine.workbook.sheet_id("Sheet1").expect("sheet exists");
        let key = CellKey {
            sheet: sheet_id,
            addr: parse_a1("B1").unwrap(),
        };
        let compiled = engine
            .workbook
            .get_cell(key)
            .and_then(|c| c.compiled.clone())
            .expect("compiled formula stored");

        let snapshot = Snapshot::from_workbook(
            &engine.workbook,
            &engine.spills,
            engine.external_value_provider.clone(),
        );
        let column_cache =
            BytecodeColumnCache::build(engine.workbook.sheets.len(), &snapshot, &[(key, compiled)]);

        let col = column_cache.by_sheet[sheet_id]
            .get(&0)
            .expect("column A is cached");
        assert_eq!(col.segments.len(), 2);
        assert_eq!(col.segments[0].row_start, 0);
        assert_eq!(col.segments[0].values.len(), 10);
        assert_eq!(col.segments[1].row_start, 899_999);
        assert_eq!(col.segments[1].values.len(), 11);
    }

    #[test]
    fn bytecode_column_cache_populates_segment_values() {
        let mut engine = Engine::new();
        engine.set_cell_value("Sheet1", "A5", 42.0).unwrap();
        engine
            .set_cell_formula("Sheet1", "B1", "=SUM(A1:A10)")
            .unwrap();

        let sheet_id = engine.workbook.sheet_id("Sheet1").expect("sheet exists");
        let key = CellKey {
            sheet: sheet_id,
            addr: parse_a1("B1").unwrap(),
        };
        let compiled = engine
            .workbook
            .get_cell(key)
            .and_then(|c| c.compiled.clone())
            .expect("compiled formula stored");

        let snapshot = Snapshot::from_workbook(
            &engine.workbook,
            &engine.spills,
            engine.external_value_provider.clone(),
        );
        let column_cache =
            BytecodeColumnCache::build(engine.workbook.sheets.len(), &snapshot, &[(key, compiled)]);

        let col = column_cache.by_sheet[sheet_id]
            .get(&0)
            .expect("column A is cached");
        assert_eq!(col.segments.len(), 1);
        let seg = &col.segments[0];
        assert_eq!(seg.row_start, 0);
        assert_eq!(seg.values[4], 42.0);
    }

    #[test]
    fn bytecode_column_cache_ignores_out_of_bounds_ranges() {
        let mut engine = Engine::new();
        engine.ensure_sheet("Sheet1");
        let sheet_id = engine.workbook.sheet_id("Sheet1").expect("sheet exists");

        let addr = parse_a1("A1").unwrap();
        let origin = bytecode::CellCoord {
            row: addr.row as i32,
            col: addr.col as i32,
        };
        let expr = bytecode::parse_formula("=SUM(XFE1:XFE2)", origin).expect("bytecode parse");
        let program = engine.bytecode_cache.get_or_compile(&expr);

        // The bytecode column cache ignores out-of-bounds ranges, but still needs a dummy AST
        // payload to satisfy the `CompiledFormula::Bytecode` wrapper.
        let parsed = crate::parse_formula("=1", crate::ParseOptions::default()).unwrap();
        let mut resolve_sheet = |name: &str| engine.workbook.sheet_id(name);
        let ast = compile_canonical_expr(&parsed.expr, sheet_id, addr, &mut resolve_sheet);

        let key = CellKey {
            sheet: sheet_id,
            addr,
        };
        let compiled = CompiledFormula::Bytecode(BytecodeFormula { ast, program });

        let snapshot = Snapshot::from_workbook(
            &engine.workbook,
            &engine.spills,
            engine.external_value_provider.clone(),
        );
        let column_cache =
            BytecodeColumnCache::build(engine.workbook.sheets.len(), &snapshot, &[(key, compiled)]);

        assert!(column_cache.by_sheet[sheet_id].is_empty());
    }

    #[test]
    fn bytecode_compiler_skips_huge_ranges() {
        let mut engine = Engine::new();
        engine
            .set_cell_formula("Sheet1", "B1", "=SUM(A1:XFD1048576)")
            .unwrap();

        // Full-sheet ranges would require enormous columnar buffers; skip bytecode compilation
        // so evaluation uses the AST engine's sparse range handling instead.
        assert_eq!(engine.bytecode_program_count(), 0);
    }

    #[test]
    fn engine_bytecode_grid_column_slice_rejects_out_of_bounds_columns() {
        let mut engine = Engine::new();
        engine.ensure_sheet("Sheet1");

        let snapshot = Snapshot::from_workbook(
            &engine.workbook,
            &engine.spills,
            engine.external_value_provider.clone(),
        );

        let mut cols = HashMap::new();
        cols.insert(
            EXCEL_MAX_COLS_I32,
            BytecodeColumn {
                segments: vec![BytecodeColumnSegment {
                    row_start: 0,
                    values: vec![1.0],
                    blocked_rows_strict: Vec::new(),
                    blocked_rows_ignore_nonnumeric: Vec::new(),
                }],
            },
        );

        let grid = EngineBytecodeGrid {
            snapshot: &snapshot,
            sheet: 0,
            cols: &cols,
            slice_mode: ColumnSliceMode::IgnoreNonNumeric,
        };

        assert!(
            bytecode::grid::Grid::column_slice(&grid, EXCEL_MAX_COLS_I32, 0, 0).is_none(),
            "out-of-bounds columns should never be eligible for SIMD slicing"
        );
    }
}
