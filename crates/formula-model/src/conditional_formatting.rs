use regex::Regex;
use serde::{Deserialize, Serialize};
use std::collections::{HashMap, HashSet};

use crate::{A1ParseError, CellRef, CellValue, Color, ErrorValue, Range};

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

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct DataBarRule {
    pub min: Cfvo,
    pub max: Cfvo,
    pub color: Option<Color>,
    // x14 extensions (Excel 2010+)
    pub min_length: Option<u8>,
    pub max_length: Option<u8>,
    pub gradient: Option<bool>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct ColorScaleRule {
    pub cfvos: Vec<Cfvo>,
    pub colors: Vec<Color>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub enum IconSet {
    ThreeArrows,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct IconSetRule {
    pub set: IconSet,
    pub cfvos: Vec<Cfvo>,
    pub show_value: bool,
    pub reverse: bool,
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
}

pub trait CellValueProvider {
    fn get_value(&self, cell: CellRef) -> Option<CellValue>;
}

pub trait FormulaEvaluator {
    fn eval(&self, formula: &str, ctx: CellRef) -> Option<CellValue>;
}

pub trait DifferentialFormatProvider {
    fn get_dxf(&self, dxf_id: u32) -> Option<CfStyleOverride>;
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct DataBarRender {
    pub color: Color,
    pub fill_ratio: f32,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct IconRender {
    pub set: IconSet,
    pub index: usize,
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

#[derive(Default)]
pub struct ConditionalFormattingEngine {
    rules: Vec<CfRule>,
    cache: HashMap<Range, CachedEvaluation>,
}

struct CachedEvaluation {
    dependencies: Vec<Range>,
    result: CfEvaluationResult,
}

impl ConditionalFormattingEngine {
    pub fn new(rules: Vec<CfRule>) -> Self {
        Self {
            rules,
            cache: HashMap::new(),
        }
    }

    pub fn set_rules(&mut self, rules: Vec<CfRule>) {
        self.rules = rules;
        self.cache.clear();
    }

    pub fn invalidate_cells<I: IntoIterator<Item = CellRef>>(&mut self, changed: I) {
        let changed: HashSet<CellRef> = changed.into_iter().collect();
        self.cache.retain(|_, entry| {
            !changed
                .iter()
                .any(|cell| entry.dependencies.iter().any(|r| r.contains(*cell)))
        });
    }

    pub fn evaluate_visible_range(
        &mut self,
        visible: Range,
        values: &dyn CellValueProvider,
        formula_evaluator: Option<&dyn FormulaEvaluator>,
        dxfs: Option<&dyn DifferentialFormatProvider>,
    ) -> &CfEvaluationResult {
        if !self.cache.contains_key(&visible) {
            let (deps, result) = evaluate_rules(&self.rules, visible, values, formula_evaluator, dxfs);
            self.cache.insert(
                visible,
                CachedEvaluation {
                    dependencies: deps,
                    result,
                },
            );
        }
        &self.cache.get(&visible).expect("cached").result
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

    let mut applicable: Vec<&CfRule> = rules
        .iter()
        .filter(|r| r.applies_to.iter().any(|ap| ranges_intersect(*ap, visible)))
        .collect();
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
            CfRuleKind::DataBar(db) => apply_data_bar(&mut result, rule, db, visible, values),
            CfRuleKind::ColorScale(cs) => apply_color_scale(&mut result, rule, cs, visible, values),
            CfRuleKind::IconSet(is) => apply_icon_set(&mut result, rule, is, visible, values),
            CfRuleKind::TopBottom(tb) => apply_top_bottom(&mut result, rule, tb, visible, values, dxfs),
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
            let entry = result.get_mut(cell).expect("visible intersection");
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
                let entry = result.get_mut(cell).expect("visible intersection");
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
) {
    let Some((min, max)) = min_max_for_ranges(&rule.applies_to, values) else {
        return;
    };
    let denom = (max - min).abs();
    let color = db.color.unwrap_or_else(|| Color::new_argb(0xFF638EC6));

    for cell in iter_rule_cells(rule, visible) {
        let Some(v) = values.get_value(cell).and_then(cell_value_as_number) else {
            continue;
        };
        let ratio = if denom == 0.0 {
            0.0
        } else {
            ((v - min) / (max - min)) as f32
        };
        let entry = result.get_mut(cell).expect("visible intersection");
        entry.data_bar = Some(DataBarRender {
            color,
            fill_ratio: ratio.clamp(0.0, 1.0),
        });
    }
}

fn apply_color_scale(
    result: &mut CfEvaluationResult,
    rule: &CfRule,
    cs: &ColorScaleRule,
    visible: Range,
    values: &dyn CellValueProvider,
) {
    let Some((min, max)) = min_max_for_ranges(&rule.applies_to, values) else {
        return;
    };
    let denom = (max - min).abs();
    if cs.colors.len() < 2 {
        return;
    }

    // TODO: Implement cfvo-driven midpoints and percentile thresholds.
    let has_mid = cs.colors.len() >= 3 && cs.cfvos.len() >= 3;
    let mid_value = (min + max) / 2.0;

    for cell in iter_rule_cells(rule, visible) {
        let Some(v) = values.get_value(cell).and_then(cell_value_as_number) else {
            continue;
        };

        let fill = if denom == 0.0 {
            cs.colors[0]
        } else if has_mid {
            if v <= mid_value {
                let t = if (mid_value - min).abs() == 0.0 {
                    0.0
                } else {
                    ((v - min) / (mid_value - min)) as f32
                };
                lerp_color(cs.colors[0], cs.colors[1], t)
            } else {
                let t = if (max - mid_value).abs() == 0.0 {
                    0.0
                } else {
                    ((v - mid_value) / (max - mid_value)) as f32
                };
                lerp_color(cs.colors[1], cs.colors[2], t)
            }
        } else {
            let t = ((v - min) / (max - min)) as f32;
            lerp_color(cs.colors[0], cs.colors[cs.colors.len() - 1], t)
        };

        let entry = result.get_mut(cell).expect("visible intersection");
        entry.style.fill = Some(fill);
    }
}

fn apply_icon_set(
    result: &mut CfEvaluationResult,
    rule: &CfRule,
    is: &IconSetRule,
    visible: Range,
    values: &dyn CellValueProvider,
) {
    let Some((min, max)) = min_max_for_ranges(&rule.applies_to, values) else {
        return;
    };
    let denom = (max - min).abs();
    if denom == 0.0 {
        return;
    }
    let thresholds = icon_set_thresholds(is, min, max);

    for cell in iter_rule_cells(rule, visible) {
        let Some(v) = values.get_value(cell).and_then(cell_value_as_number) else {
            continue;
        };
        let mut idx = 0;
        if v >= thresholds[1] {
            idx = 2;
        } else if v >= thresholds[0] {
            idx = 1;
        }
        let entry = result.get_mut(cell).expect("visible intersection");
        entry.icon = Some(IconRender {
            set: is.set,
            index: idx,
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
            let entry = result.get_mut(cell).expect("visible intersection");
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
            let entry = result.get_mut(cell).expect("visible intersection");
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

fn resolve_dxf(dxf_id: Option<u32>, dxfs: Option<&dyn DifferentialFormatProvider>) -> CfStyleOverride {
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
        CellValue::Array(a) => !a.data.is_empty(),
        CellValue::Spill(_) => true,
    }
}

fn min_max_for_ranges(ranges: &[Range], values: &dyn CellValueProvider) -> Option<(f64, f64)> {
    let mut min = f64::INFINITY;
    let mut max = f64::NEG_INFINITY;
    let mut found = false;

    for r in ranges {
        for cell in iter_cells(*r) {
            if let Some(v) = values.get_value(cell).and_then(cell_value_as_number) {
                if v.is_nan() {
                    continue;
                }
                found = true;
                if v < min {
                    min = v;
                }
                if v > max {
                    max = v;
                }
            }
        }
    }

    if found {
        Some((min, max))
    } else {
        None
    }
}

fn eval_threshold(
    formula: &str,
    ctx: CellRef,
    values: &dyn CellValueProvider,
    formula_evaluator: Option<&dyn FormulaEvaluator>,
) -> Option<f64> {
    let trimmed = formula.trim();
    if trimmed.is_empty() {
        return None;
    }
    if let Ok(num) = trimmed.parse::<f64>() {
        return Some(num);
    }
    if let Some(fe) = formula_evaluator {
        return fe.eval(trimmed, ctx).and_then(cell_value_as_number);
    }
    // Support a bare A1 reference even without a formula evaluator.
    if let Ok(cell) = CellRef::from_a1(trimmed) {
        return values.get_value(cell).and_then(cell_value_as_number);
    }
    None
}

fn icon_set_thresholds(rule: &IconSetRule, min: f64, max: f64) -> [f64; 2] {
    let mut t1 = 33.0;
    let mut t2 = 67.0;
    if rule.cfvos.len() >= 3 {
        if let Some(v) = rule.cfvos.get(1).and_then(|c| c.value.as_deref()) {
            if let Ok(n) = v.parse::<f64>() {
                t1 = n;
            }
        }
        if let Some(v) = rule.cfvos.get(2).and_then(|c| c.value.as_deref()) {
            if let Ok(n) = v.parse::<f64>() {
                t2 = n;
            }
        }
    }
    let span = max - min;
    [min + span * (t1 / 100.0), min + span * (t2 / 100.0)]
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
    static CELL_REF_RE: std::sync::OnceLock<Regex> = std::sync::OnceLock::new();
    let re = CELL_REF_RE.get_or_init(|| {
        Regex::new(r"(?i)(?:^|[^A-Z0-9_])(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)")
            .expect("valid regex")
    });

    let mut refs: Vec<Range> = Vec::new();
    for cap in re.captures_iter(formula) {
        if let Some(m) = cap.get(1) {
            // Avoid matching function names that look like cell references (e.g. LOG10()).
            if formula[m.end()..].starts_with('(') {
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
            CellValue::Array(_) => None,
            CellValue::Spill(_) => None,
        }
    }
}

pub fn format_render_plan(visible: Range, eval: &CfEvaluationResult) -> String {
    let mut lines = Vec::new();
    for cell in iter_cells(visible) {
        let label = cell.to_a1();
        let res = eval.get(cell).expect("cell in visible range");
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
    a.start.row <= b.end.row && a.end.row >= b.start.row && a.start.col <= b.end.col && a.end.col >= b.start.col
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
        cur: range.start,
    }
}

struct RangeIter {
    range: Range,
    cur: CellRef,
}

impl Iterator for RangeIter {
    type Item = CellRef;

    fn next(&mut self) -> Option<Self::Item> {
        if self.cur.row > self.range.end.row {
            return None;
        }
        let out = self.cur;
        if self.cur.col == self.range.end.col {
            self.cur = CellRef {
                row: self.cur.row + 1,
                col: self.range.start.col,
            };
        } else {
            self.cur = CellRef {
                row: self.cur.row,
                col: self.cur.col + 1,
            };
        }
        Some(out)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn extract_refs_avoids_function_name() {
        let refs = extract_a1_references("LOG10(A1)+Sheet1!$B$2");
        assert_eq!(refs, vec![parse_range_a1("A1").unwrap(), parse_range_a1("$B$2").unwrap()]);
    }
}
