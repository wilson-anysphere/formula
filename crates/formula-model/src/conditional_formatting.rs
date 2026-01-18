use regex::Regex;
use serde::{Deserialize, Serialize};
use std::cmp::Ordering;
use std::collections::HashMap;

use crate::{
    A1ParseError, CellRef, CellValue, Color, ErrorValue, Fill, FillPattern, Font, Range, Style,
    StyleTable,
};

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub enum CfRuleSchema {
    /// Conditional formatting stored using the SpreadsheetML 2006 schema (Excel 2007).
    Office2007,
    /// Conditional formatting extensions stored using the x14 schema (Excel 2010+).
    X14,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub enum CellIsOperator {
    GreaterThan,
    GreaterThanOrEqual,
    LessThan,
    LessThanOrEqual,
    Equal,
    NotEqual,
    Between,
    NotBetween,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub enum CfvoType {
    Min,
    Max,
    Number,
    Percent,
    Percentile,
    Formula,
    AutoMin,
    AutoMax,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct Cfvo {
    pub type_: CfvoType,
    pub value: Option<String>,
}

impl Cfvo {
    fn rewrite_sheet_references(&mut self, old_name: &str, new_name: &str) {
        if self.type_ != CfvoType::Formula {
            return;
        }
        let Some(value) = self.value.as_mut() else {
            return;
        };
        *value = crate::rewrite_sheet_names_in_formula(value, old_name, new_name);
    }

    fn rewrite_sheet_references_internal_refs_only(&mut self, old_name: &str, new_name: &str) {
        if self.type_ != CfvoType::Formula {
            return;
        }
        let Some(value) = self.value.as_mut() else {
            return;
        };
        *value = crate::formula_rewrite::rewrite_sheet_names_in_formula_internal_refs_only(
            value, old_name, new_name,
        );
    }

    fn rewrite_table_references(&mut self, renames: &[(String, String)]) {
        if self.type_ != CfvoType::Formula {
            return;
        }
        let Some(value) = self.value.as_mut() else {
            return;
        };
        *value = crate::rewrite_table_names_in_formula(value, renames);
    }

    fn invalidate_deleted_sheet_references(&mut self, deleted_sheet: &str, sheet_order: &[String]) {
        if self.type_ != CfvoType::Formula {
            return;
        }
        let Some(value) = self.value.as_mut() else {
            return;
        };
        *value =
            crate::rewrite_deleted_sheet_references_in_formula(value, deleted_sheet, sheet_order);
    }
}

pub fn parse_argb_hex_color(s: &str) -> Option<Color> {
    let s = s.trim();
    if s.len() != 8 {
        return None;
    }
    u32::from_str_radix(s, 16).ok().map(Color::new_argb)
}

fn color_to_argb_hex(color: Color) -> String {
    format!("{:08X}", color.argb().unwrap_or(0))
}

fn lerp_color(a: Color, b: Color, t: f32) -> Color {
    let t = t.clamp(0.0, 1.0);

    let split = |argb: u32| {
        (
            (argb >> 24) & 0xFF,
            (argb >> 16) & 0xFF,
            (argb >> 8) & 0xFF,
            argb & 0xFF,
        )
    };

    let (aa, ar, ag, ab) = split(a.argb().unwrap_or(0));
    let (ba, br, bg, bb) = split(b.argb().unwrap_or(0));

    let lerp_u8 = |x: u32, y: u32| -> u32 { (x as f32 + (y as f32 - x as f32) * t).round() as u32 };

    Color::new_argb(
        (lerp_u8(aa, ba) << 24)
            | (lerp_u8(ar, br) << 16)
            | (lerp_u8(ag, bg) << 8)
            | lerp_u8(ab, bb),
    )
}

/// A partial style to apply to a cell when a conditional formatting rule matches.
///
/// This is intentionally a tri-state overlay (e.g. `bold: None` means "do not
/// touch", `Some(true)` means force bold on, `Some(false)` means force bold off).
#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
pub struct CfStyleOverride {
    pub fill: Option<Color>,
    pub font_color: Option<Color>,
    pub bold: Option<bool>,
    pub italic: Option<bool>,
}

/// Direction that data bars are drawn in (x14 extension).
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub enum DataBarDirection {
    LeftToRight,
    RightToLeft,
    /// Excel's default: follow sheet / locale direction.
    Context,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct DataBarRule {
    pub min: Cfvo,
    pub max: Cfvo,
    pub color: Option<Color>,
    // x14 extensions (Excel 2010+)
    pub min_length: Option<u8>,
    pub max_length: Option<u8>,
    pub gradient: Option<bool>,
    pub negative_fill_color: Option<Color>,
    pub axis_color: Option<Color>,
    pub direction: Option<DataBarDirection>,
}

impl DataBarRule {
    fn rewrite_sheet_references(&mut self, old_name: &str, new_name: &str) {
        self.min.rewrite_sheet_references(old_name, new_name);
        self.max.rewrite_sheet_references(old_name, new_name);
    }

    fn rewrite_sheet_references_internal_refs_only(&mut self, old_name: &str, new_name: &str) {
        self.min
            .rewrite_sheet_references_internal_refs_only(old_name, new_name);
        self.max
            .rewrite_sheet_references_internal_refs_only(old_name, new_name);
    }

    fn rewrite_table_references(&mut self, renames: &[(String, String)]) {
        self.min.rewrite_table_references(renames);
        self.max.rewrite_table_references(renames);
    }

    fn invalidate_deleted_sheet_references(&mut self, deleted_sheet: &str, sheet_order: &[String]) {
        self.min
            .invalidate_deleted_sheet_references(deleted_sheet, sheet_order);
        self.max
            .invalidate_deleted_sheet_references(deleted_sheet, sheet_order);
    }
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct ColorScaleRule {
    pub cfvos: Vec<Cfvo>,
    pub colors: Vec<Color>,
}

impl ColorScaleRule {
    fn rewrite_sheet_references(&mut self, old_name: &str, new_name: &str) {
        for cfvo in &mut self.cfvos {
            cfvo.rewrite_sheet_references(old_name, new_name);
        }
    }

    fn rewrite_sheet_references_internal_refs_only(&mut self, old_name: &str, new_name: &str) {
        for cfvo in &mut self.cfvos {
            cfvo.rewrite_sheet_references_internal_refs_only(old_name, new_name);
        }
    }

    fn rewrite_table_references(&mut self, renames: &[(String, String)]) {
        for cfvo in &mut self.cfvos {
            cfvo.rewrite_table_references(renames);
        }
    }

    fn invalidate_deleted_sheet_references(&mut self, deleted_sheet: &str, sheet_order: &[String]) {
        for cfvo in &mut self.cfvos {
            cfvo.invalidate_deleted_sheet_references(deleted_sheet, sheet_order);
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub enum IconSet {
    ThreeArrows,
    ThreeTrafficLights1,
    ThreeTrafficLights2,
    ThreeFlags,
    ThreeSymbols,
    ThreeSymbols2,
    FourArrows,
    FourArrowsGray,
    FiveArrows,
    FiveArrowsGray,
    FiveQuarters,
}

impl IconSet {
    pub fn icon_count(self) -> usize {
        match self {
            IconSet::ThreeArrows
            | IconSet::ThreeTrafficLights1
            | IconSet::ThreeTrafficLights2
            | IconSet::ThreeFlags
            | IconSet::ThreeSymbols
            | IconSet::ThreeSymbols2 => 3,
            IconSet::FourArrows | IconSet::FourArrowsGray => 4,
            IconSet::FiveArrows | IconSet::FiveArrowsGray | IconSet::FiveQuarters => 5,
        }
    }
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct IconSetRule {
    pub set: IconSet,
    pub cfvos: Vec<Cfvo>,
    pub show_value: bool,
    pub reverse: bool,
}

impl IconSetRule {
    fn rewrite_sheet_references(&mut self, old_name: &str, new_name: &str) {
        for cfvo in &mut self.cfvos {
            cfvo.rewrite_sheet_references(old_name, new_name);
        }
    }

    fn rewrite_sheet_references_internal_refs_only(&mut self, old_name: &str, new_name: &str) {
        for cfvo in &mut self.cfvos {
            cfvo.rewrite_sheet_references_internal_refs_only(old_name, new_name);
        }
    }

    fn rewrite_table_references(&mut self, renames: &[(String, String)]) {
        for cfvo in &mut self.cfvos {
            cfvo.rewrite_table_references(renames);
        }
    }

    fn invalidate_deleted_sheet_references(&mut self, deleted_sheet: &str, sheet_order: &[String]) {
        for cfvo in &mut self.cfvos {
            cfvo.invalidate_deleted_sheet_references(deleted_sheet, sheet_order);
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub enum TopBottomKind {
    Top,
    Bottom,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct TopBottomRule {
    pub kind: TopBottomKind,
    pub rank: u32,
    pub percent: bool,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct UniqueDuplicateRule {
    pub unique: bool,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub enum CfRuleKind {
    CellIs {
        operator: CellIsOperator,
        formulas: Vec<String>,
    },
    Expression {
        formula: String,
    },
    DataBar(DataBarRule),
    ColorScale(ColorScaleRule),
    IconSet(IconSetRule),
    TopBottom(TopBottomRule),
    UniqueDuplicate(UniqueDuplicateRule),
    Unsupported {
        type_name: Option<String>,
        raw_xml: String,
    },
}

impl CfRuleKind {
    pub(crate) fn rewrite_sheet_references(&mut self, old_name: &str, new_name: &str) {
        match self {
            CfRuleKind::CellIs { formulas, .. } => {
                for formula in formulas {
                    *formula = crate::rewrite_sheet_names_in_formula(formula, old_name, new_name);
                }
            }
            CfRuleKind::Expression { formula } => {
                *formula = crate::rewrite_sheet_names_in_formula(formula, old_name, new_name);
            }
            CfRuleKind::DataBar(rule) => rule.rewrite_sheet_references(old_name, new_name),
            CfRuleKind::ColorScale(rule) => rule.rewrite_sheet_references(old_name, new_name),
            CfRuleKind::IconSet(rule) => rule.rewrite_sheet_references(old_name, new_name),
            CfRuleKind::TopBottom(_)
            | CfRuleKind::UniqueDuplicate(_)
            | CfRuleKind::Unsupported { .. } => {}
        }
    }

    pub(crate) fn rewrite_sheet_references_internal_refs_only(
        &mut self,
        old_name: &str,
        new_name: &str,
    ) {
        match self {
            CfRuleKind::CellIs { formulas, .. } => {
                for formula in formulas {
                    *formula =
                        crate::formula_rewrite::rewrite_sheet_names_in_formula_internal_refs_only(
                            formula, old_name, new_name,
                        );
                }
            }
            CfRuleKind::Expression { formula } => {
                *formula =
                    crate::formula_rewrite::rewrite_sheet_names_in_formula_internal_refs_only(
                        formula, old_name, new_name,
                    );
            }
            CfRuleKind::DataBar(rule) => {
                rule.rewrite_sheet_references_internal_refs_only(old_name, new_name)
            }
            CfRuleKind::ColorScale(rule) => {
                rule.rewrite_sheet_references_internal_refs_only(old_name, new_name)
            }
            CfRuleKind::IconSet(rule) => {
                rule.rewrite_sheet_references_internal_refs_only(old_name, new_name)
            }
            CfRuleKind::TopBottom(_)
            | CfRuleKind::UniqueDuplicate(_)
            | CfRuleKind::Unsupported { .. } => {}
        }
    }

    pub(crate) fn rewrite_table_references(&mut self, renames: &[(String, String)]) {
        match self {
            CfRuleKind::CellIs { formulas, .. } => {
                for formula in formulas {
                    *formula = crate::rewrite_table_names_in_formula(formula, renames);
                }
            }
            CfRuleKind::Expression { formula } => {
                *formula = crate::rewrite_table_names_in_formula(formula, renames);
            }
            CfRuleKind::DataBar(rule) => rule.rewrite_table_references(renames),
            CfRuleKind::ColorScale(rule) => rule.rewrite_table_references(renames),
            CfRuleKind::IconSet(rule) => rule.rewrite_table_references(renames),
            CfRuleKind::TopBottom(_)
            | CfRuleKind::UniqueDuplicate(_)
            | CfRuleKind::Unsupported { .. } => {}
        }
    }

    pub(crate) fn invalidate_deleted_sheet_references(
        &mut self,
        deleted_sheet: &str,
        sheet_order: &[String],
    ) {
        match self {
            CfRuleKind::CellIs { formulas, .. } => {
                for formula in formulas {
                    *formula = crate::rewrite_deleted_sheet_references_in_formula(
                        formula,
                        deleted_sheet,
                        sheet_order,
                    );
                }
            }
            CfRuleKind::Expression { formula } => {
                *formula = crate::rewrite_deleted_sheet_references_in_formula(
                    formula,
                    deleted_sheet,
                    sheet_order,
                );
            }
            CfRuleKind::DataBar(rule) => {
                rule.invalidate_deleted_sheet_references(deleted_sheet, sheet_order)
            }
            CfRuleKind::ColorScale(rule) => {
                rule.invalidate_deleted_sheet_references(deleted_sheet, sheet_order)
            }
            CfRuleKind::IconSet(rule) => {
                rule.invalidate_deleted_sheet_references(deleted_sheet, sheet_order)
            }
            CfRuleKind::TopBottom(_)
            | CfRuleKind::UniqueDuplicate(_)
            | CfRuleKind::Unsupported { .. } => {}
        }
    }
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct CfRule {
    pub schema: CfRuleSchema,
    pub id: Option<String>,
    pub priority: u32,
    pub applies_to: Vec<Range>,
    pub dxf_id: Option<u32>,
    pub stop_if_true: bool,
    pub kind: CfRuleKind,
    /// Best-effort dependency ranges used for cache invalidation.
    pub dependencies: Vec<Range>,
}

impl CfRule {
    pub fn applies_to_cell(&self, cell: CellRef) -> bool {
        self.applies_to.iter().any(|r| r.contains(cell))
    }

    pub(crate) fn rewrite_sheet_references(&mut self, old_name: &str, new_name: &str) {
        self.kind.rewrite_sheet_references(old_name, new_name);
    }

    pub(crate) fn rewrite_sheet_references_internal_refs_only(
        &mut self,
        old_name: &str,
        new_name: &str,
    ) {
        self.kind
            .rewrite_sheet_references_internal_refs_only(old_name, new_name);
    }

    pub(crate) fn rewrite_table_references(&mut self, renames: &[(String, String)]) {
        self.kind.rewrite_table_references(renames);
    }

    pub(crate) fn invalidate_deleted_sheet_references(
        &mut self,
        deleted_sheet: &str,
        sheet_order: &[String],
    ) {
        self.kind
            .invalidate_deleted_sheet_references(deleted_sheet, sheet_order);
    }
}

pub trait CellValueProvider {
    fn get_value(&self, cell: CellRef) -> Option<CellValue>;
}

impl CellValueProvider for crate::Worksheet {
    fn get_value(&self, cell: CellRef) -> Option<CellValue> {
        Some(self.value(cell))
    }
}

pub trait FormulaEvaluator {
    fn eval(&self, formula: &str, ctx: CellRef) -> Option<CellValue>;
}

pub trait DifferentialFormatProvider {
    fn get_dxf(&self, dxf_id: u32) -> Option<CfStyleOverride>;
}

impl DifferentialFormatProvider for Vec<CfStyleOverride> {
    fn get_dxf(&self, dxf_id: u32) -> Option<CfStyleOverride> {
        self.get(dxf_id as usize).cloned()
    }
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct DataBarRender {
    pub color: Color,
    pub fill_ratio: f32,
    pub min_length: u8,
    pub max_length: u8,
    pub gradient: bool,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct IconRender {
    pub set: IconSet,
    pub index: usize,
    pub show_value: bool,
    pub reverse: bool,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
pub struct CellFormatResult {
    pub style: CfStyleOverride,
    pub data_bar: Option<DataBarRender>,
    pub icon: Option<IconRender>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct CfEvaluationResult {
    pub visible: Range,
    pub cols: u32,
    pub rows: u32,
    pub cells: Vec<CellFormatResult>,
}

impl CfEvaluationResult {
    pub fn new(visible: Range) -> Self {
        let cols = visible.width();
        let rows = visible.height();
        let len = cols as usize * rows as usize;
        Self {
            visible,
            cols,
            rows,
            cells: vec![CellFormatResult::default(); len],
        }
    }

    pub fn get(&self, cell: CellRef) -> Option<&CellFormatResult> {
        let idx = self.index(cell)?;
        self.cells.get(idx)
    }

    pub fn get_mut(&mut self, cell: CellRef) -> Option<&mut CellFormatResult> {
        let idx = self.index(cell)?;
        self.cells.get_mut(idx)
    }

    fn index(&self, cell: CellRef) -> Option<usize> {
        if !self.visible.contains(cell) {
            return None;
        }
        let row_off = cell.row - self.visible.start.row;
        let col_off = cell.col - self.visible.start.col;
        Some((row_off * self.cols + col_off) as usize)
    }
}

/// Fully resolved formatting for a cell after applying conditional formatting.
///
/// This is intended for viewport rendering: callers can resolve the base style
/// from the workbook [`StyleTable`] and apply the [`CellFormatResult`] overlay
/// from [`CfEvaluationResult`].
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct ResolvedCellFormat {
    pub style: Style,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_bar: Option<DataBarRender>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub icon: Option<IconRender>,
}

/// Apply a conditional formatting style override on top of an existing [`Style`].
pub fn apply_cf_style_override(style: &mut Style, overlay: &CfStyleOverride) {
    if let Some(fill) = overlay.fill {
        let fill_style = style.fill.get_or_insert_with(Fill::default);
        fill_style.pattern = FillPattern::Solid;
        fill_style.fg_color = Some(fill);
        fill_style.bg_color = None;
    }

    if overlay.font_color.is_some() || overlay.bold.is_some() || overlay.italic.is_some() {
        let font = style.font.get_or_insert_with(Font::default);
        if let Some(color) = overlay.font_color {
            font.color = Some(color);
        }
        if let Some(bold) = overlay.bold {
            font.bold = bold;
        }
        if let Some(italic) = overlay.italic {
            font.italic = italic;
        }
    }
}

/// Resolve the final per-cell formatting to render by merging:
/// - the base cell style (via [`StyleTable`])
/// - the conditional formatting result (`cf`)
///
/// This helper is pure (no UI dependencies) and is suitable for building a
/// renderer-facing payload.
pub fn resolve_cell_formatting(
    styles: &StyleTable,
    base_style_id: u32,
    cf: &CellFormatResult,
) -> ResolvedCellFormat {
    let mut style = styles.get(base_style_id).cloned().unwrap_or_default();
    apply_cf_style_override(&mut style, &cf.style);
    ResolvedCellFormat {
        style,
        data_bar: cf.data_bar.clone(),
        icon: cf.icon.clone(),
    }
}

#[derive(Clone, Debug, Default)]
pub struct ConditionalFormattingEngine {
    cache: HashMap<Range, CachedEvaluation>,
}

#[derive(Clone, Debug)]
struct CachedEvaluation {
    dependencies: Vec<Range>,
    result: CfEvaluationResult,
}

impl ConditionalFormattingEngine {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn clear_cache(&mut self) {
        self.cache.clear();
    }

    pub fn invalidate_cells<I: IntoIterator<Item = CellRef>>(&mut self, changed: I) {
        let iter = changed.into_iter();
        let (lower, upper) = iter.size_hint();
        let reserve = upper.unwrap_or(lower);
        let mut changed_cells: Vec<CellRef> = Vec::new();
        if changed_cells.try_reserve(reserve).is_err() {
            // Best-effort: if we can't materialize the change set, invalidate everything.
            debug_assert!(
                false,
                "allocation failed (conditional formatting invalidate_cells, hint={reserve})"
            );
            self.cache.clear();
            return;
        }
        for cell in iter {
            changed_cells.push(cell);
        }
        self.cache.retain(|_, entry| {
            !changed_cells
                .iter()
                .any(|cell| entry.dependencies.iter().any(|r| r.contains(*cell)))
        });
    }

    pub fn evaluate_visible_range(
        &mut self,
        rules: &[CfRule],
        visible: Range,
        values: &dyn CellValueProvider,
        formula_evaluator: Option<&dyn FormulaEvaluator>,
        dxfs: Option<&dyn DifferentialFormatProvider>,
    ) -> &CfEvaluationResult {
        use std::collections::hash_map::Entry;

        match self.cache.entry(visible) {
            Entry::Occupied(entry) => &entry.into_mut().result,
            Entry::Vacant(entry) => {
                let visible = *entry.key();
                let (deps, result) =
                    evaluate_rules(rules, visible, values, formula_evaluator, dxfs);
                &entry
                    .insert(CachedEvaluation {
                        dependencies: deps,
                        result,
                    })
                    .result
            }
        }
    }
}

fn evaluate_rules(
    rules: &[CfRule],
    visible: Range,
    values: &dyn CellValueProvider,
    formula_evaluator: Option<&dyn FormulaEvaluator>,
    dxfs: Option<&dyn DifferentialFormatProvider>,
) -> (Vec<Range>, CfEvaluationResult) {
    let mut result = CfEvaluationResult::new(visible);

    let mut applicable: Vec<&CfRule> = Vec::new();
    if applicable.try_reserve(rules.len()).is_err() {
        debug_assert!(
            false,
            "allocation failed (conditional formatting applicable rules, total={})",
            rules.len()
        );
        return (Vec::new(), result);
    }
    for rule in rules {
        if rule
            .applies_to
            .iter()
            .any(|ap| ranges_intersect(*ap, visible))
        {
            applicable.push(rule);
        }
    }
    applicable.sort_by_key(|r| r.priority);

    let mut deps: Vec<Range> = Vec::new();
    for rule in &applicable {
        deps.extend(rule.dependencies.iter().copied());
    }
    deps = normalize_ranges(deps);

    for rule in applicable {
        match &rule.kind {
            CfRuleKind::CellIs { operator, formulas } => apply_cell_is(
                &mut result,
                rule,
                *operator,
                formulas,
                visible,
                values,
                formula_evaluator,
                dxfs,
            ),
            CfRuleKind::Expression { formula } => {
                apply_expression(&mut result, rule, formula, visible, formula_evaluator, dxfs)
            }
            CfRuleKind::DataBar(db) => {
                apply_data_bar(&mut result, rule, db, visible, values, formula_evaluator)
            }
            CfRuleKind::ColorScale(cs) => {
                apply_color_scale(&mut result, rule, cs, visible, values, formula_evaluator)
            }
            CfRuleKind::IconSet(is) => {
                apply_icon_set(&mut result, rule, is, visible, values, formula_evaluator)
            }
            CfRuleKind::TopBottom(tb) => {
                apply_top_bottom(&mut result, rule, tb, visible, values, dxfs)
            }
            CfRuleKind::UniqueDuplicate(ud) => {
                apply_unique_duplicate(&mut result, rule, ud, visible, values, dxfs)
            }
            CfRuleKind::Unsupported { .. } => {}
        }
    }

    (deps, result)
}

fn apply_cell_is(
    result: &mut CfEvaluationResult,
    rule: &CfRule,
    operator: CellIsOperator,
    formulas: &[String],
    visible: Range,
    values: &dyn CellValueProvider,
    formula_evaluator: Option<&dyn FormulaEvaluator>,
    dxfs: Option<&dyn DifferentialFormatProvider>,
) {
    let style = resolve_dxf(rule.dxf_id, dxfs);

    for cell in iter_rule_cells(rule, visible) {
        let threshold1 = eval_threshold(
            formulas.get(0).map(String::as_str).unwrap_or(""),
            cell,
            values,
            formula_evaluator,
        );
        let threshold2 = formulas
            .get(1)
            .and_then(|f| eval_threshold(f, cell, values, formula_evaluator));

        let Some(cell_value) = values.get_value(cell).and_then(cell_value_as_number) else {
            continue;
        };

        let matches = match operator {
            CellIsOperator::GreaterThan => threshold1.map_or(false, |t| cell_value > t),
            CellIsOperator::GreaterThanOrEqual => threshold1.map_or(false, |t| cell_value >= t),
            CellIsOperator::LessThan => threshold1.map_or(false, |t| cell_value < t),
            CellIsOperator::LessThanOrEqual => threshold1.map_or(false, |t| cell_value <= t),
            CellIsOperator::Equal => threshold1.map_or(false, |t| cell_value == t),
            CellIsOperator::NotEqual => threshold1.map_or(false, |t| cell_value != t),
            CellIsOperator::Between => match (threshold1, threshold2) {
                (Some(a), Some(b)) => cell_value >= a.min(b) && cell_value <= a.max(b),
                _ => false,
            },
            CellIsOperator::NotBetween => match (threshold1, threshold2) {
                (Some(a), Some(b)) => cell_value < a.min(b) || cell_value > a.max(b),
                _ => false,
            },
        };

        if matches {
            let Some(entry) = result.get_mut(cell) else {
                debug_assert!(false, "expected cell in visible range: {cell:?}");
                continue;
            };
            entry.style = merge_style(entry.style.clone(), style.clone());
        }
    }
}

fn apply_expression(
    result: &mut CfEvaluationResult,
    rule: &CfRule,
    formula: &str,
    visible: Range,
    formula_evaluator: Option<&dyn FormulaEvaluator>,
    dxfs: Option<&dyn DifferentialFormatProvider>,
) {
    let Some(fe) = formula_evaluator else {
        return;
    };
    let style = resolve_dxf(rule.dxf_id, dxfs);

    for cell in iter_rule_cells(rule, visible) {
        if let Some(v) = fe.eval(formula, cell) {
            if cell_value_truthy(&v) {
                let Some(entry) = result.get_mut(cell) else {
                    debug_assert!(false, "expected cell in visible range: {cell:?}");
                    continue;
                };
                entry.style = merge_style(entry.style.clone(), style.clone());
            }
        }
    }
}

fn apply_data_bar(
    result: &mut CfEvaluationResult,
    rule: &CfRule,
    db: &DataBarRule,
    visible: Range,
    values: &dyn CellValueProvider,
    formula_evaluator: Option<&dyn FormulaEvaluator>,
) {
    let Some(ctx) = rule_context_cell(rule) else {
        return;
    };
    let Some(mut stats) = NumericRangeStats::collect(&rule.applies_to, values) else {
        return;
    };
    let Some(min) = cfvo_threshold(
        &db.min,
        &mut stats,
        ctx,
        values,
        formula_evaluator,
        CfvoContextKind::DataBar,
    ) else {
        return;
    };
    let Some(max) = cfvo_threshold(
        &db.max,
        &mut stats,
        ctx,
        values,
        formula_evaluator,
        CfvoContextKind::DataBar,
    ) else {
        return;
    };
    let denom = max - min;
    let color = db.color.unwrap_or_else(|| Color::new_argb(0xFF638EC6));
    let min_length = db.min_length.unwrap_or(0).min(100);
    let max_length = db.max_length.unwrap_or(100).min(100);
    let gradient = db.gradient.unwrap_or(true);

    for cell in iter_rule_cells(rule, visible) {
        let Some(v) = values.get_value(cell).and_then(cell_value_as_number) else {
            continue;
        };
        if v.is_nan() {
            continue;
        }
        let ratio = if denom == 0.0 {
            0.0
        } else {
            ((v - min) / denom) as f32
        };
        let Some(entry) = result.get_mut(cell) else {
            debug_assert!(false, "expected cell in visible range: {cell:?}");
            continue;
        };
        entry.data_bar = Some(DataBarRender {
            color,
            fill_ratio: ratio.clamp(0.0, 1.0),
            min_length,
            max_length,
            gradient,
        });
    }
}

fn apply_color_scale(
    result: &mut CfEvaluationResult,
    rule: &CfRule,
    cs: &ColorScaleRule,
    visible: Range,
    values: &dyn CellValueProvider,
    formula_evaluator: Option<&dyn FormulaEvaluator>,
) {
    if cs.colors.len() < 2 || cs.cfvos.len() < 2 {
        return;
    }
    let Some(ctx) = rule_context_cell(rule) else {
        return;
    };
    let Some(mut stats) = NumericRangeStats::collect(&rule.applies_to, values) else {
        return;
    };

    let has_mid = cs.colors.len() >= 3 && cs.cfvos.len() >= 3;
    let min_cfvo = &cs.cfvos[0];
    let max_cfvo = &cs.cfvos[cs.cfvos.len() - 1];
    let mid_cfvo = has_mid.then(|| &cs.cfvos[1]);

    let Some(min) = cfvo_threshold(
        min_cfvo,
        &mut stats,
        ctx,
        values,
        formula_evaluator,
        CfvoContextKind::Other,
    ) else {
        return;
    };
    let Some(max) = cfvo_threshold(
        max_cfvo,
        &mut stats,
        ctx,
        values,
        formula_evaluator,
        CfvoContextKind::Other,
    ) else {
        return;
    };
    let mid = mid_cfvo.and_then(|m| {
        cfvo_threshold(
            m,
            &mut stats,
            ctx,
            values,
            formula_evaluator,
            CfvoContextKind::Other,
        )
    });
    let min_color = cs.colors[0];
    let max_color = cs.colors[cs.colors.len() - 1];

    for cell in iter_rule_cells(rule, visible) {
        let Some(v) = values.get_value(cell).and_then(cell_value_as_number) else {
            continue;
        };
        if v.is_nan() {
            continue;
        }

        let fill = if has_mid && mid.is_some() {
            let Some(mid) = mid else {
                debug_assert!(false, "expected mid threshold when has_mid is set");
                continue;
            };
            let mid_color = cs.colors[1];
            let high_color = cs.colors[2];
            if v <= mid {
                let denom = mid - min;
                let t = if denom == 0.0 {
                    0.0
                } else {
                    ((v - min) / denom) as f32
                };
                lerp_color(min_color, mid_color, t)
            } else {
                let denom = max - mid;
                let t = if denom == 0.0 {
                    0.0
                } else {
                    ((v - mid) / denom) as f32
                };
                lerp_color(mid_color, high_color, t)
            }
        } else {
            let denom = max - min;
            let t = if denom == 0.0 {
                0.0
            } else {
                ((v - min) / denom) as f32
            };
            lerp_color(min_color, max_color, t)
        };

        let Some(entry) = result.get_mut(cell) else {
            debug_assert!(false, "expected cell in visible range: {cell:?}");
            continue;
        };
        entry.style.fill = Some(fill);
    }
}

fn apply_icon_set(
    result: &mut CfEvaluationResult,
    rule: &CfRule,
    is: &IconSetRule,
    visible: Range,
    values: &dyn CellValueProvider,
    formula_evaluator: Option<&dyn FormulaEvaluator>,
) {
    let Some(ctx) = rule_context_cell(rule) else {
        return;
    };
    let Some(mut stats) = NumericRangeStats::collect(&rule.applies_to, values) else {
        return;
    };
    let Some(thresholds) = icon_set_thresholds(is, &mut stats, ctx, values, formula_evaluator)
    else {
        return;
    };
    let icon_count = is.set.icon_count();

    for cell in iter_rule_cells(rule, visible) {
        let Some(v) = values.get_value(cell).and_then(cell_value_as_number) else {
            continue;
        };
        if v.is_nan() {
            continue;
        }
        let mut idx = thresholds.iter().take_while(|t| v >= **t).count();
        if is.reverse && icon_count > 0 {
            idx = (icon_count - 1).saturating_sub(idx);
        }
        let Some(entry) = result.get_mut(cell) else {
            debug_assert!(false, "expected cell in visible range: {cell:?}");
            continue;
        };
        entry.icon = Some(IconRender {
            set: is.set,
            index: idx,
            show_value: is.show_value,
            reverse: is.reverse,
        });
    }
}

fn apply_top_bottom(
    result: &mut CfEvaluationResult,
    rule: &CfRule,
    tb: &TopBottomRule,
    visible: Range,
    values: &dyn CellValueProvider,
    dxfs: Option<&dyn DifferentialFormatProvider>,
) {
    let style = resolve_dxf(rule.dxf_id, dxfs);
    let mut all: Vec<(CellRef, f64)> = Vec::new();
    for ap in &rule.applies_to {
        for cell in iter_cells(*ap) {
            if let Some(v) = values.get_value(cell).and_then(cell_value_as_number) {
                all.push((cell, v));
            }
        }
    }
    if all.is_empty() {
        return;
    }
    all.sort_by(|a, b| a.1.partial_cmp(&b.1).unwrap_or(std::cmp::Ordering::Equal));
    let count = all.len() as u32;
    let rank = if tb.percent {
        let pct = tb.rank.min(100);
        ((count as f64) * (pct as f64 / 100.0)).ceil().max(1.0) as usize
    } else {
        tb.rank.max(1) as usize
    };
    let rank = rank.min(all.len());
    let threshold = match tb.kind {
        TopBottomKind::Top => {
            let idx = all.len() - rank;
            all[idx].1
        }
        TopBottomKind::Bottom => all[rank - 1].1,
    };

    for cell in iter_rule_cells(rule, visible) {
        let Some(v) = values.get_value(cell).and_then(cell_value_as_number) else {
            continue;
        };
        let matches = match tb.kind {
            TopBottomKind::Top => v >= threshold,
            TopBottomKind::Bottom => v <= threshold,
        };
        if matches {
            let Some(entry) = result.get_mut(cell) else {
                debug_assert!(false, "expected cell in visible range: {cell:?}");
                continue;
            };
            entry.style = merge_style(entry.style.clone(), style.clone());
        }
    }
}

fn apply_unique_duplicate(
    result: &mut CfEvaluationResult,
    rule: &CfRule,
    ud: &UniqueDuplicateRule,
    visible: Range,
    values: &dyn CellValueProvider,
    dxfs: Option<&dyn DifferentialFormatProvider>,
) {
    let style = resolve_dxf(rule.dxf_id, dxfs);
    let mut counts: HashMap<ValueKey, u32> = HashMap::new();
    for ap in &rule.applies_to {
        for cell in iter_cells(*ap) {
            if let Some(v) = values.get_value(cell).and_then(ValueKey::from_cell_value) {
                *counts.entry(v).or_default() += 1;
            }
        }
    }
    if counts.is_empty() {
        return;
    }

    for cell in iter_rule_cells(rule, visible) {
        let Some(key) = values.get_value(cell).and_then(ValueKey::from_cell_value) else {
            continue;
        };
        let c = *counts.get(&key).unwrap_or(&0);
        let matches = if ud.unique { c == 1 } else { c > 1 };
        if matches {
            let Some(entry) = result.get_mut(cell) else {
                debug_assert!(false, "expected cell in visible range: {cell:?}");
                continue;
            };
            entry.style = merge_style(entry.style.clone(), style.clone());
        }
    }
}

fn iter_rule_cells<'a>(rule: &'a CfRule, visible: Range) -> impl Iterator<Item = CellRef> + 'a {
    rule.applies_to
        .iter()
        .filter_map(move |ap| range_intersection(*ap, visible))
        .flat_map(iter_cells)
}

fn resolve_dxf(
    dxf_id: Option<u32>,
    dxfs: Option<&dyn DifferentialFormatProvider>,
) -> CfStyleOverride {
    match (dxf_id, dxfs) {
        (Some(id), Some(p)) => p.get_dxf(id).unwrap_or_default(),
        _ => CfStyleOverride::default(),
    }
}

fn merge_style(mut base: CfStyleOverride, overlay: CfStyleOverride) -> CfStyleOverride {
    if overlay.fill.is_some() {
        base.fill = overlay.fill;
    }
    if overlay.font_color.is_some() {
        base.font_color = overlay.font_color;
    }
    if overlay.bold.is_some() {
        base.bold = overlay.bold;
    }
    if overlay.italic.is_some() {
        base.italic = overlay.italic;
    }
    base
}

fn cell_value_as_number(value: CellValue) -> Option<f64> {
    match value {
        CellValue::Number(n) => Some(n),
        _ => None,
    }
}

fn cell_value_truthy(value: &CellValue) -> bool {
    match value {
        CellValue::Boolean(b) => *b,
        CellValue::Number(n) => *n != 0.0 && !n.is_nan(),
        CellValue::String(s) => !s.is_empty(),
        CellValue::Empty => false,
        CellValue::Error(_) => false,
        CellValue::RichText(t) => !t.text.is_empty(),
        CellValue::Entity(e) => !e.display_value.is_empty(),
        // Record values are treated like text for conditional formatting truthiness:
        // truthy iff their derived display string is non-empty.
        CellValue::Record(r) => {
            if let Some(field) = r.display_field.as_deref() {
                if let Some(value) = r.get_field_case_insensitive(field) {
                    return match value {
                        CellValue::Empty => false,
                        CellValue::String(s) => !s.is_empty(),
                        CellValue::RichText(rt) => !rt.text.is_empty(),
                        // Scalar display strings for these variants are always non-empty.
                        CellValue::Number(_) | CellValue::Boolean(_) | CellValue::Error(_) => true,
                        CellValue::Entity(entity) => !entity.display_value.is_empty(),
                        CellValue::Record(record) => !record.to_string().is_empty(),
                        CellValue::Image(_) => true,
                        // Non-scalar displayField values fall back to `display_value`.
                        _ => !r.display_value.is_empty(),
                    };
                }
            }
            !r.display_value.is_empty()
        }
        CellValue::Image(image) => !image.image_id.as_str().is_empty(),
        CellValue::Array(a) => !a.data.is_empty(),
        CellValue::Spill(_) => true,
    }
}

fn rule_context_cell(rule: &CfRule) -> Option<CellRef> {
    rule.applies_to
        .iter()
        .map(|r| r.start)
        .min_by_key(|c| (c.row, c.col))
}

#[derive(Clone, Debug)]
struct NumericRangeStats {
    values: Vec<f64>,
    sorted: bool,
    min: f64,
    max: f64,
}

impl NumericRangeStats {
    fn collect(ranges: &[Range], values: &dyn CellValueProvider) -> Option<Self> {
        let mut out = Self {
            values: Vec::new(),
            sorted: false,
            min: f64::INFINITY,
            max: f64::NEG_INFINITY,
        };

        for r in ranges {
            for cell in iter_cells(*r) {
                let Some(v) = values.get_value(cell).and_then(cell_value_as_number) else {
                    continue;
                };
                if v.is_nan() {
                    continue;
                }
                out.values.push(v);
                if v < out.min {
                    out.min = v;
                }
                if v > out.max {
                    out.max = v;
                }
            }
        }

        if out.values.is_empty() {
            return None;
        }

        Some(out)
    }

    fn ensure_sorted(&mut self) {
        if self.sorted {
            return;
        }
        self.values
            .sort_by(|a, b| a.partial_cmp(b).unwrap_or(Ordering::Equal));
        self.sorted = true;
    }

    /// "Percent" threshold semantics for conditional formatting cfvos.
    ///
    /// This behaves like Excel's percentile interpolation (similar to `PERCENTILE.INC`),
    /// where `percent=0` is the minimum and `percent=100` is the maximum.
    fn percentile_rank(&mut self, percent: f64) -> Option<f64> {
        if percent.is_nan() {
            return None;
        }
        let n = self.values.len();
        if n == 0 {
            return None;
        }
        self.ensure_sorted();
        if n == 1 {
            return Some(self.values[0]);
        }

        let p = (percent / 100.0).clamp(0.0, 1.0);
        let pos = p * (n as f64 - 1.0);
        let lo = pos.floor() as usize;
        let hi = pos.ceil() as usize;
        if lo == hi {
            return Some(self.values[lo]);
        }
        let t = pos - lo as f64;
        Some(self.values[lo] + (self.values[hi] - self.values[lo]) * t)
    }

    /// "Percentile" threshold semantics for conditional formatting cfvos.
    ///
    /// This uses a nearest-rank selection: `ceil(p/100 * N)` (clamped).
    fn percentile_nearest_rank(&mut self, percent: f64) -> Option<f64> {
        if percent.is_nan() {
            return None;
        }
        let n = self.values.len();
        if n == 0 {
            return None;
        }
        self.ensure_sorted();
        if n == 1 {
            return Some(self.values[0]);
        }

        let p = (percent / 100.0).clamp(0.0, 1.0);
        let rank = (p * (n as f64)).ceil() as usize;
        let idx = rank.saturating_sub(1).min(n - 1);
        Some(self.values[idx])
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum CfvoContextKind {
    DataBar,
    Other,
}

fn cfvo_threshold(
    cfvo: &Cfvo,
    stats: &mut NumericRangeStats,
    ctx: CellRef,
    values: &dyn CellValueProvider,
    formula_evaluator: Option<&dyn FormulaEvaluator>,
    kind: CfvoContextKind,
) -> Option<f64> {
    let v = match cfvo.type_ {
        CfvoType::Min => Some(stats.min),
        CfvoType::Max => Some(stats.max),
        CfvoType::AutoMin => match kind {
            CfvoContextKind::DataBar => Some(if stats.min >= 0.0 { 0.0 } else { stats.min }),
            CfvoContextKind::Other => Some(stats.min),
        },
        CfvoType::AutoMax => match kind {
            CfvoContextKind::DataBar => Some(if stats.max <= 0.0 { 0.0 } else { stats.max }),
            CfvoContextKind::Other => Some(stats.max),
        },
        CfvoType::Number => cfvo
            .value
            .as_deref()
            .and_then(|s| s.trim().parse::<f64>().ok())
            .filter(|n| !n.is_nan()),
        CfvoType::Percent => cfvo
            .value
            .as_deref()
            .and_then(|s| s.trim().parse::<f64>().ok())
            .and_then(|p| stats.percentile_rank(p)),
        CfvoType::Percentile => cfvo
            .value
            .as_deref()
            .and_then(|s| s.trim().parse::<f64>().ok())
            .and_then(|p| stats.percentile_nearest_rank(p)),
        CfvoType::Formula => eval_threshold(
            cfvo.value.as_deref().unwrap_or(""),
            ctx,
            values,
            formula_evaluator,
        ),
    };

    v.filter(|n| !n.is_nan())
}

fn icon_set_thresholds(
    rule: &IconSetRule,
    stats: &mut NumericRangeStats,
    ctx: CellRef,
    values: &dyn CellValueProvider,
    formula_evaluator: Option<&dyn FormulaEvaluator>,
) -> Option<Vec<f64>> {
    let icon_count = rule.set.icon_count();
    if icon_count < 2 {
        return None;
    }
    let mut thresholds: Vec<f64> = Vec::new();
    if thresholds.try_reserve_exact(icon_count.saturating_sub(1)).is_err() {
        debug_assert!(
            false,
            "allocation failed (icon set thresholds, count={})",
            icon_count.saturating_sub(1)
        );
        return None;
    }
    for boundary in 1..icon_count {
        let default_percent = 100.0 * (boundary as f64) / (icon_count as f64);
        let threshold = rule
            .cfvos
            .get(boundary)
            .and_then(|cfvo| {
                cfvo_threshold(
                    cfvo,
                    stats,
                    ctx,
                    values,
                    formula_evaluator,
                    CfvoContextKind::Other,
                )
            })
            .or_else(|| stats.percentile_rank(default_percent));

        thresholds.push(threshold?);
    }

    Some(thresholds)
}

fn eval_threshold(
    formula: &str,
    ctx: CellRef,
    values: &dyn CellValueProvider,
    formula_evaluator: Option<&dyn FormulaEvaluator>,
) -> Option<f64> {
    let trimmed = formula.trim();
    let expr = trimmed.strip_prefix('=').unwrap_or(trimmed).trim();
    if expr.is_empty() {
        return None;
    }
    if let Ok(num) = expr.parse::<f64>() {
        if !num.is_nan() {
            return Some(num);
        }
    }
    if let Some(fe) = formula_evaluator {
        if let Some(v) = fe
            .eval(expr, ctx)
            .and_then(cell_value_as_number)
            .filter(|n| !n.is_nan())
        {
            return Some(v);
        }
        if expr != trimmed {
            return fe
                .eval(trimmed, ctx)
                .and_then(cell_value_as_number)
                .filter(|n| !n.is_nan());
        }
    }
    // Support a bare A1 reference even without a formula evaluator.
    if let Ok(cell) = CellRef::from_a1(expr) {
        return values
            .get_value(cell)
            .and_then(cell_value_as_number)
            .filter(|n| !n.is_nan());
    }
    if let Some((_, a1)) = expr.rsplit_once('!') {
        if let Ok(cell) = CellRef::from_a1(a1) {
            return values
                .get_value(cell)
                .and_then(cell_value_as_number)
                .filter(|n| !n.is_nan());
        }
    }
    None
}

/// Parse a SpreadsheetML `sqref` attribute (`A1`, `A1:B2`, `A1 A3:B7`).
pub fn parse_sqref(sqref: &str) -> Result<Vec<Range>, A1ParseError> {
    let mut out = Vec::new();
    for token in sqref.split_whitespace() {
        if token.is_empty() {
            continue;
        }
        out.push(parse_range_a1(token)?);
    }
    Ok(out)
}

/// Parse an A1 range (`A1` or `A1:B2`) into a normalized [`Range`].
pub fn parse_range_a1(a1: &str) -> Result<Range, A1ParseError> {
    let a1 = a1.trim();
    if let Some((lhs, rhs)) = a1.split_once(':') {
        let start = CellRef::from_a1(lhs)?;
        let end = CellRef::from_a1(rhs)?;
        Ok(Range::new(start, end))
    } else {
        let cell = CellRef::from_a1(a1)?;
        Ok(Range::new(cell, cell))
    }
}

/// Best-effort extractor of A1-style references from a formula string.
///
/// This intentionally does not support structured references or R1C1.
pub fn extract_a1_references(formula: &str) -> Vec<Range> {
    static CELL_REF_RE: std::sync::OnceLock<Option<Regex>> = std::sync::OnceLock::new();
    let Some(re) = CELL_REF_RE
        .get_or_init(|| {
            Regex::new(
                r"(?i)(?:^|[^A-Z0-9_])(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)",
            )
            .ok()
        })
        .as_ref()
    else {
        debug_assert!(false, "failed to compile A1 reference regex");
        return Vec::new();
    };

    let mut refs: Vec<Range> = Vec::new();
    for cap in re.captures_iter(formula) {
        if let Some(m) = cap.get(1) {
            // Avoid matching function names that look like cell references (e.g. LOG10()).
            if formula[m.end()..].starts_with('(') {
                continue;
            }
            // Avoid matching sheet names that look like cell references (e.g. `ABC1!A1`).
            if formula[m.end()..].starts_with('!') || formula[m.end()..].starts_with("'!") {
                continue;
            }
            if let Ok(r) = parse_range_a1(m.as_str()) {
                refs.push(r);
            }
        }
    }

    normalize_ranges(refs)
}

fn normalize_ranges(mut ranges: Vec<Range>) -> Vec<Range> {
    ranges.sort_by_key(|r| (r.start.row, r.start.col, r.end.row, r.end.col));
    ranges.dedup();
    ranges
}

#[derive(Clone, Debug, PartialEq, Eq, Hash)]
enum ValueKey {
    Number(u64),
    String(String),
    Boolean(bool),
    Error(ErrorValue),
}

impl ValueKey {
    fn from_cell_value(value: CellValue) -> Option<Self> {
        match value {
            CellValue::Number(n) => Some(ValueKey::Number(n.to_bits())),
            CellValue::String(s) => Some(ValueKey::String(s)),
            CellValue::Boolean(b) => Some(ValueKey::Boolean(b)),
            CellValue::Error(e) => Some(ValueKey::Error(e)),
            CellValue::Empty => None,
            CellValue::RichText(t) => Some(ValueKey::String(t.text)),
            CellValue::Entity(e) => Some(ValueKey::String(e.display_value)),
            CellValue::Record(r) => Some(ValueKey::String(r.to_string())),
            CellValue::Image(image) => Some(ValueKey::String(
                image
                    .alt_text
                    .as_deref()
                    .filter(|s| !s.is_empty())
                    .unwrap_or("[Image]")
                    .to_string(),
            )),
            CellValue::Array(_) => None,
            CellValue::Spill(_) => None,
        }
    }
}

pub fn format_render_plan(visible: Range, eval: &CfEvaluationResult) -> String {
    let mut lines = Vec::new();
    for cell in iter_cells(visible) {
        let mut label = String::new();
        crate::push_a1_cell_ref(cell.row, cell.col, false, false, &mut label);
        let Some(res) = eval.get(cell) else {
            debug_assert!(false, "expected cell in visible range: {cell:?}");
            lines.push(format!("{label}: missing"));
            continue;
        };
        let mut parts = Vec::new();

        if let Some(fill) = res.style.fill {
            parts.push(format!("fill={}", color_to_argb_hex(fill)));
        }
        if let Some(font) = res.style.font_color {
            parts.push(format!("font={}", color_to_argb_hex(font)));
        }
        if let Some(db) = &res.data_bar {
            parts.push(format!(
                "dataBar color={} ratio={:.2}",
                color_to_argb_hex(db.color),
                db.fill_ratio
            ));
        }
        if let Some(icon) = &res.icon {
            parts.push(format!("icon={:?}:{}", icon.set, icon.index));
        }

        if parts.is_empty() {
            parts.push("none".to_string());
        }

        lines.push(format!("{label}: {}", parts.join(" ")));
    }

    lines.join("\n")
}

fn ranges_intersect(a: Range, b: Range) -> bool {
    a.start.row <= b.end.row
        && a.end.row >= b.start.row
        && a.start.col <= b.end.col
        && a.end.col >= b.start.col
}

fn range_intersection(a: Range, b: Range) -> Option<Range> {
    if !ranges_intersect(a, b) {
        return None;
    }
    Some(Range {
        start: CellRef {
            row: a.start.row.max(b.start.row),
            col: a.start.col.max(b.start.col),
        },
        end: CellRef {
            row: a.end.row.min(b.end.row),
            col: a.end.col.min(b.end.col),
        },
    })
}

fn iter_cells(range: Range) -> RangeIter {
    RangeIter {
        range,
        cur: Some(range.start),
    }
}

struct RangeIter {
    range: Range,
    cur: Option<CellRef>,
}

impl Iterator for RangeIter {
    type Item = CellRef;

    fn next(&mut self) -> Option<Self::Item> {
        let current = self.cur?;
        if current.row > self.range.end.row {
            self.cur = None;
            return None;
        }

        let out = current;
        if current.row == self.range.end.row && current.col == self.range.end.col {
            self.cur = None;
            return Some(out);
        }

        if current.col == self.range.end.col {
            // Safe: if we're not at `range.end`, then `current.row < range.end.row`,
            // so incrementing cannot overflow `u32`.
            self.cur = Some(CellRef {
                row: current.row + 1,
                col: self.range.start.col,
            });
        } else {
            // Safe: if we're not at `range.end`, then `current.col < range.end.col`.
            self.cur = Some(CellRef {
                row: current.row,
                col: current.col + 1,
            });
        }
        Some(out)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::collections::HashMap;

    struct CellMapFormulaEvaluator {
        formula: String,
        results: HashMap<CellRef, CellValue>,
    }

    impl FormulaEvaluator for CellMapFormulaEvaluator {
        fn eval(&self, formula: &str, ctx: CellRef) -> Option<CellValue> {
            assert_eq!(formula.trim(), self.formula);
            self.results.get(&ctx).cloned()
        }
    }

    #[derive(Default)]
    struct TestValues {
        values: HashMap<CellRef, CellValue>,
    }

    impl TestValues {
        fn with_numbers<const N: usize>(items: [(&'static str, f64); N]) -> Self {
            let mut values = HashMap::new();
            for (a1, n) in items {
                values.insert(CellRef::from_a1(a1).unwrap(), CellValue::Number(n));
            }
            Self { values }
        }
    }

    impl CellValueProvider for TestValues {
        fn get_value(&self, cell: CellRef) -> Option<CellValue> {
            self.values.get(&cell).cloned()
        }
    }

    struct AssertCtxFormulaEvaluator {
        expected_ctx: CellRef,
        results: HashMap<String, CellValue>,
    }

    impl FormulaEvaluator for AssertCtxFormulaEvaluator {
        fn eval(&self, formula: &str, ctx: CellRef) -> Option<CellValue> {
            assert_eq!(
                ctx, self.expected_ctx,
                "formula evaluated with unexpected ctx"
            );
            self.results.get(formula.trim()).cloned()
        }
    }

    #[test]
    fn extract_refs_avoids_function_name() {
        let refs = extract_a1_references("LOG10(A1)+Sheet1!$B$2");
        assert_eq!(
            refs,
            vec![
                parse_range_a1("A1").unwrap(),
                parse_range_a1("$B$2").unwrap()
            ]
        );
    }

    #[test]
    fn extract_refs_avoids_sheet_names_that_look_like_cells() {
        let refs = extract_a1_references("SUM('ABC1'!A1, ABC1!B2)");
        assert_eq!(
            refs,
            vec![parse_range_a1("A1").unwrap(), parse_range_a1("B2").unwrap()]
        );
    }

    #[test]
    fn percent_and_percentile_thresholds_differ() {
        let values =
            TestValues::with_numbers([("A1", 0.0), ("A2", 0.0), ("A3", 100.0), ("A4", 100.0)]);

        let visible = parse_range_a1("A1:A4").unwrap();
        let percent_rule = CfRule {
            schema: CfRuleSchema::Office2007,
            id: None,
            priority: 1,
            applies_to: vec![visible],
            dxf_id: None,
            stop_if_true: false,
            kind: CfRuleKind::IconSet(IconSetRule {
                set: IconSet::ThreeArrows,
                cfvos: vec![
                    Cfvo {
                        type_: CfvoType::Min,
                        value: None,
                    },
                    Cfvo {
                        type_: CfvoType::Percent,
                        value: Some("50".to_string()),
                    },
                    Cfvo {
                        type_: CfvoType::Max,
                        value: None,
                    },
                ],
                show_value: true,
                reverse: false,
            }),
            dependencies: vec![],
        };

        let percentile_rule = CfRule {
            kind: CfRuleKind::IconSet(IconSetRule {
                set: IconSet::ThreeArrows,
                cfvos: vec![
                    Cfvo {
                        type_: CfvoType::Min,
                        value: None,
                    },
                    Cfvo {
                        type_: CfvoType::Percentile,
                        value: Some("50".to_string()),
                    },
                    Cfvo {
                        type_: CfvoType::Max,
                        value: None,
                    },
                ],
                show_value: true,
                reverse: false,
            }),
            ..percent_rule.clone()
        };

        let rules = vec![percent_rule];
        let mut engine = ConditionalFormattingEngine::new();
        let eval = engine.evaluate_visible_range(&rules, visible, &values, None, None);
        assert_eq!(
            eval.get(CellRef::from_a1("A1").unwrap())
                .unwrap()
                .icon
                .as_ref()
                .unwrap()
                .index,
            0
        );
        assert_eq!(
            eval.get(CellRef::from_a1("A2").unwrap())
                .unwrap()
                .icon
                .as_ref()
                .unwrap()
                .index,
            0
        );

        let rules = vec![percentile_rule];
        let mut engine = ConditionalFormattingEngine::new();
        let eval = engine.evaluate_visible_range(&rules, visible, &values, None, None);
        assert_eq!(
            eval.get(CellRef::from_a1("A1").unwrap())
                .unwrap()
                .icon
                .as_ref()
                .unwrap()
                .index,
            1
        );
        assert_eq!(
            eval.get(CellRef::from_a1("A2").unwrap())
                .unwrap()
                .icon
                .as_ref()
                .unwrap()
                .index,
            1
        );
    }

    #[test]
    fn color_scale_midpoint_is_driven_by_cfvo() {
        let values =
            TestValues::with_numbers([("A1", 0.0), ("A2", 1.0), ("A3", 2.0), ("A4", 100.0)]);
        let visible = parse_range_a1("A1:A4").unwrap();

        let rule = CfRule {
            schema: CfRuleSchema::Office2007,
            id: None,
            priority: 1,
            applies_to: vec![visible],
            dxf_id: None,
            stop_if_true: false,
            kind: CfRuleKind::ColorScale(ColorScaleRule {
                cfvos: vec![
                    Cfvo {
                        type_: CfvoType::Min,
                        value: None,
                    },
                    Cfvo {
                        type_: CfvoType::Percentile,
                        value: Some("50".to_string()),
                    },
                    Cfvo {
                        type_: CfvoType::Max,
                        value: None,
                    },
                ],
                colors: vec![
                    Color::new_argb(0xFFFF0000),
                    Color::new_argb(0xFFFFFF00),
                    Color::new_argb(0xFF00FF00),
                ],
            }),
            dependencies: vec![],
        };

        let rules = vec![rule];
        let mut engine = ConditionalFormattingEngine::new();
        let eval = engine.evaluate_visible_range(&rules, visible, &values, None, None);
        assert_eq!(
            eval.get(CellRef::from_a1("A2").unwrap())
                .unwrap()
                .style
                .fill,
            Some(Color::new_argb(0xFFFFFF00)),
            "value at cfvo-driven midpoint should receive the midpoint color"
        );
    }

    #[test]
    fn formula_cfvo_thresholds_and_databar_options_are_applied() {
        let values = TestValues::with_numbers([("B1", 100.0), ("B2", 200.0)]);
        let visible = parse_range_a1("B1:B2").unwrap();

        let evaluator = AssertCtxFormulaEvaluator {
            expected_ctx: CellRef::from_a1("B1").unwrap(),
            results: HashMap::from([("BAR_MAX".to_string(), CellValue::Number(200.0))]),
        };

        let rule = CfRule {
            schema: CfRuleSchema::Office2007,
            id: None,
            priority: 1,
            applies_to: vec![visible],
            dxf_id: None,
            stop_if_true: false,
            kind: CfRuleKind::DataBar(DataBarRule {
                min: Cfvo {
                    type_: CfvoType::Number,
                    value: Some("0".to_string()),
                },
                max: Cfvo {
                    type_: CfvoType::Formula,
                    value: Some("BAR_MAX".to_string()),
                },
                color: None,
                min_length: Some(10),
                max_length: Some(90),
                gradient: Some(false),
                negative_fill_color: None,
                axis_color: None,
                direction: None,
            }),
            dependencies: vec![],
        };

        let rules = vec![rule];
        let mut engine = ConditionalFormattingEngine::new();
        let eval = engine.evaluate_visible_range(&rules, visible, &values, Some(&evaluator), None);

        let bar1 = eval
            .get(CellRef::from_a1("B1").unwrap())
            .unwrap()
            .data_bar
            .as_ref()
            .unwrap();
        assert!((bar1.fill_ratio - 0.5).abs() < 1e-6);
        assert_eq!(bar1.min_length, 10);
        assert_eq!(bar1.max_length, 90);
        assert!(!bar1.gradient);
    }

    #[test]
    fn icon_set_reverse_and_show_value_are_recorded() {
        let values = TestValues::with_numbers([("C1", 0.0), ("C2", 100.0), ("C3", 200.0)]);
        let visible = parse_range_a1("C1:C3").unwrap();

        let rule = CfRule {
            schema: CfRuleSchema::Office2007,
            id: None,
            priority: 1,
            applies_to: vec![visible],
            dxf_id: None,
            stop_if_true: false,
            kind: CfRuleKind::IconSet(IconSetRule {
                set: IconSet::ThreeArrows,
                cfvos: vec![
                    Cfvo {
                        type_: CfvoType::Min,
                        value: None,
                    },
                    Cfvo {
                        type_: CfvoType::Number,
                        value: Some("100".to_string()),
                    },
                    Cfvo {
                        type_: CfvoType::Number,
                        value: Some("200".to_string()),
                    },
                ],
                show_value: false,
                reverse: true,
            }),
            dependencies: vec![],
        };

        let rules = vec![rule];
        let mut engine = ConditionalFormattingEngine::new();
        let eval = engine.evaluate_visible_range(&rules, visible, &values, None, None);
        let c1 = eval
            .get(CellRef::from_a1("C1").unwrap())
            .unwrap()
            .icon
            .as_ref()
            .unwrap();
        let c3 = eval
            .get(CellRef::from_a1("C3").unwrap())
            .unwrap()
            .icon
            .as_ref()
            .unwrap();
        assert_eq!(c1.index, 2);
        assert_eq!(c3.index, 0);
        assert!(!c1.show_value);
        assert!(c1.reverse);
    }

    #[test]
    fn databar_automatic_min_includes_zero_baseline_for_all_positive_values() {
        let values = TestValues::with_numbers([("D1", 10.0), ("D2", 20.0)]);
        let visible = parse_range_a1("D1:D2").unwrap();

        let rule = CfRule {
            schema: CfRuleSchema::Office2007,
            id: None,
            priority: 1,
            applies_to: vec![visible],
            dxf_id: None,
            stop_if_true: false,
            kind: CfRuleKind::DataBar(DataBarRule {
                min: Cfvo {
                    type_: CfvoType::AutoMin,
                    value: None,
                },
                max: Cfvo {
                    type_: CfvoType::AutoMax,
                    value: None,
                },
                color: None,
                min_length: None,
                max_length: None,
                gradient: None,
                negative_fill_color: None,
                axis_color: None,
                direction: None,
            }),
            dependencies: vec![],
        };

        let rules = vec![rule];
        let mut engine = ConditionalFormattingEngine::new();
        let eval = engine.evaluate_visible_range(&rules, visible, &values, None, None);
        let d1 = eval
            .get(CellRef::from_a1("D1").unwrap())
            .unwrap()
            .data_bar
            .as_ref()
            .unwrap();
        assert!((d1.fill_ratio - 0.5).abs() < 1e-6);
    }

    #[test]
    fn expression_truthiness_treats_entities_like_text_display_string() {
        let visible = parse_range_a1("A1:A2").unwrap();

        let rule = CfRule {
            schema: CfRuleSchema::Office2007,
            id: None,
            priority: 1,
            applies_to: vec![visible],
            dxf_id: Some(0),
            stop_if_true: false,
            kind: CfRuleKind::Expression {
                formula: "ENTITY".to_string(),
            },
            dependencies: vec![],
        };

        let dxfs = vec![CfStyleOverride {
            fill: Some(Color::new_argb(0xFFFF0000)),
            font_color: None,
            bold: None,
            italic: None,
        }];

        let evaluator = CellMapFormulaEvaluator {
            formula: "ENTITY".to_string(),
            results: HashMap::from([
                (
                    CellRef::from_a1("A1").unwrap(),
                    CellValue::Entity(crate::EntityValue::new("Seattle")),
                ),
                (
                    CellRef::from_a1("A2").unwrap(),
                    CellValue::Entity(crate::EntityValue::new("")),
                ),
            ]),
        };

        let values = TestValues::default();
        let mut engine = ConditionalFormattingEngine::new();
        let eval =
            engine.evaluate_visible_range(&[rule], visible, &values, Some(&evaluator), Some(&dxfs));

        assert_eq!(
            eval.get(CellRef::from_a1("A1").unwrap())
                .unwrap()
                .style
                .fill,
            Some(Color::new_argb(0xFFFF0000))
        );
        assert_eq!(
            eval.get(CellRef::from_a1("A2").unwrap())
                .unwrap()
                .style
                .fill,
            None
        );
    }

    #[test]
    fn expression_truthiness_treats_records_like_text_display_string() {
        let visible = parse_range_a1("A1:A2").unwrap();

        let rule = CfRule {
            schema: CfRuleSchema::Office2007,
            id: None,
            priority: 1,
            applies_to: vec![visible],
            dxf_id: Some(0),
            stop_if_true: false,
            kind: CfRuleKind::Expression {
                formula: "RECORD".to_string(),
            },
            dependencies: vec![],
        };

        let dxfs = vec![CfStyleOverride {
            fill: Some(Color::new_argb(0xFFFF0000)),
            font_color: None,
            bold: None,
            italic: None,
        }];

        let evaluator = CellMapFormulaEvaluator {
            formula: "RECORD".to_string(),
            results: HashMap::from([
                (
                    CellRef::from_a1("A1").unwrap(),
                    CellValue::Record(crate::RecordValue {
                        display_field: Some("name".to_string()),
                        fields: std::collections::BTreeMap::from([(
                            "name".to_string(),
                            CellValue::String("Ada".to_string()),
                        )]),
                        ..Default::default()
                    }),
                ),
                (
                    CellRef::from_a1("A2").unwrap(),
                    CellValue::Record(crate::RecordValue::default()),
                ),
            ]),
        };

        let values = TestValues::default();
        let mut engine = ConditionalFormattingEngine::new();
        let eval =
            engine.evaluate_visible_range(&[rule], visible, &values, Some(&evaluator), Some(&dxfs));

        assert_eq!(
            eval.get(CellRef::from_a1("A1").unwrap())
                .unwrap()
                .style
                .fill,
            Some(Color::new_argb(0xFFFF0000))
        );
        assert_eq!(
            eval.get(CellRef::from_a1("A2").unwrap())
                .unwrap()
                .style
                .fill,
            None
        );
    }

    #[test]
    fn expression_truthiness_treats_record_display_field_entities_like_text_display_string() {
        let visible = parse_range_a1("A1:A2").unwrap();

        let rule = CfRule {
            schema: CfRuleSchema::Office2007,
            id: None,
            priority: 1,
            applies_to: vec![visible],
            dxf_id: Some(0),
            stop_if_true: false,
            kind: CfRuleKind::Expression {
                formula: "RECORD_ENTITY".to_string(),
            },
            dependencies: vec![],
        };

        let dxfs = vec![CfStyleOverride {
            fill: Some(Color::new_argb(0xFFFF0000)),
            font_color: None,
            bold: None,
            italic: None,
        }];

        let evaluator = CellMapFormulaEvaluator {
            formula: "RECORD_ENTITY".to_string(),
            results: HashMap::from([
                (
                    CellRef::from_a1("A1").unwrap(),
                    CellValue::Record(crate::RecordValue {
                        display_field: Some("company".to_string()),
                        fields: std::collections::BTreeMap::from([(
                            "company".to_string(),
                            CellValue::Entity(crate::EntityValue::new("Apple")),
                        )]),
                        ..Default::default()
                    }),
                ),
                (
                    CellRef::from_a1("A2").unwrap(),
                    CellValue::Record(crate::RecordValue {
                        display_field: Some("company".to_string()),
                        fields: std::collections::BTreeMap::from([(
                            "company".to_string(),
                            CellValue::Entity(crate::EntityValue::new("")),
                        )]),
                        ..Default::default()
                    }),
                ),
            ]),
        };

        let values = TestValues::default();
        let mut engine = ConditionalFormattingEngine::new();
        let eval =
            engine.evaluate_visible_range(&[rule], visible, &values, Some(&evaluator), Some(&dxfs));

        assert_eq!(
            eval.get(CellRef::from_a1("A1").unwrap())
                .unwrap()
                .style
                .fill,
            Some(Color::new_argb(0xFFFF0000))
        );
        assert_eq!(
            eval.get(CellRef::from_a1("A2").unwrap())
                .unwrap()
                .style
                .fill,
            None
        );
    }

    #[test]
    fn value_key_maps_rich_values_to_display_string() {
        let entity = crate::EntityValue::new("Seattle");
        assert_eq!(
            ValueKey::from_cell_value(CellValue::Entity(entity)),
            Some(ValueKey::String("Seattle".to_string()))
        );

        let record = crate::RecordValue {
            display_field: Some("name".to_string()),
            fields: std::collections::BTreeMap::from([(
                "name".to_string(),
                CellValue::String("Ada".to_string()),
            )]),
            ..Default::default()
        };
        let record_display = record.to_string();
        assert_eq!(
            ValueKey::from_cell_value(CellValue::Record(record)),
            Some(ValueKey::String(record_display))
        );

        let image = crate::ImageValue {
            image_id: crate::drawings::ImageId::new("image1.png"),
            alt_text: Some("Logo".to_string()),
            width: None,
            height: None,
        };
        assert_eq!(
            ValueKey::from_cell_value(CellValue::Image(image)),
            Some(ValueKey::String("Logo".to_string()))
        );

        let record_entity = crate::RecordValue {
            display_field: Some("company".to_string()),
            fields: std::collections::BTreeMap::from([(
                "company".to_string(),
                CellValue::Entity(crate::EntityValue::new("Apple")),
            )]),
            ..Default::default()
        };
        let record_entity_display = record_entity.to_string();
        assert_eq!(
            ValueKey::from_cell_value(CellValue::Record(record_entity)),
            Some(ValueKey::String(record_entity_display))
        );

        let record_nested = crate::RecordValue {
            display_field: Some("person".to_string()),
            fields: std::collections::BTreeMap::from([(
                "person".to_string(),
                CellValue::Record(crate::RecordValue {
                    display_field: Some("name".to_string()),
                    fields: std::collections::BTreeMap::from([(
                        "name".to_string(),
                        CellValue::String("Ada".to_string()),
                    )]),
                    ..Default::default()
                }),
            )]),
            ..Default::default()
        };
        let record_nested_display = record_nested.to_string();
        assert_eq!(
            ValueKey::from_cell_value(CellValue::Record(record_nested)),
            Some(ValueKey::String(record_nested_display))
        );
    }

    #[test]
    fn numeric_coercion_does_not_parse_rich_value_display_strings() {
        assert_eq!(
            cell_value_as_number(CellValue::Entity(crate::EntityValue::new("123"))),
            None
        );
        assert_eq!(
            cell_value_as_number(CellValue::Record(crate::RecordValue {
                display_field: Some("n".to_string()),
                fields: std::collections::BTreeMap::from([(
                    "n".to_string(),
                    CellValue::String("123".to_string()),
                )]),
                ..Default::default()
            })),
            None
        );
    }

    #[test]
    fn databar_rule_serde_round_trip_preserves_x14_fields() {
        let rule = DataBarRule {
            min: Cfvo {
                type_: CfvoType::AutoMin,
                value: None,
            },
            max: Cfvo {
                type_: CfvoType::AutoMax,
                value: None,
            },
            color: Some(Color::new_argb(0xFF638EC6)),
            min_length: Some(0),
            max_length: Some(100),
            gradient: Some(false),
            negative_fill_color: Some(Color::new_argb(0xFFFF0000)),
            axis_color: Some(Color::new_argb(0xFF000000)),
            direction: Some(DataBarDirection::LeftToRight),
        };

        let json = serde_json::to_string(&rule).expect("serialize");
        let round_tripped: DataBarRule = serde_json::from_str(&json).expect("deserialize");
        assert_eq!(round_tripped, rule);
    }

    #[test]
    fn databar_rule_serde_backwards_compatible_when_x14_fields_absent() {
        let rule = DataBarRule {
            min: Cfvo {
                type_: CfvoType::AutoMin,
                value: None,
            },
            max: Cfvo {
                type_: CfvoType::AutoMax,
                value: None,
            },
            color: None,
            min_length: None,
            max_length: None,
            gradient: None,
            negative_fill_color: Some(Color::new_argb(0xFFFF0000)),
            axis_color: Some(Color::new_argb(0xFF000000)),
            direction: Some(DataBarDirection::RightToLeft),
        };

        let mut value = serde_json::to_value(&rule).expect("serialize to value");
        let obj = value.as_object_mut().expect("object");
        obj.remove("negative_fill_color");
        obj.remove("axis_color");
        obj.remove("direction");

        let deserialized: DataBarRule =
            serde_json::from_value(value).expect("deserialize without x14 fields");
        assert_eq!(deserialized.negative_fill_color, None);
        assert_eq!(deserialized.axis_color, None);
        assert_eq!(deserialized.direction, None);
    }
}
