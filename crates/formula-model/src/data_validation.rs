use serde::{Deserialize, Serialize};

use crate::value::text_eq_case_insensitive;
use crate::{CellValue, Range};

pub type DataValidationId = u32;

fn is_false(b: &bool) -> bool {
    !*b
}

/// Excel-style data validation rule kind.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum DataValidationKind {
    Whole,
    Decimal,
    List,
    Date,
    Time,
    TextLength,
    Custom,
}

/// Excel-style comparison operator for validation rules.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum DataValidationOperator {
    Between,
    NotBetween,
    Equal,
    NotEqual,
    GreaterThan,
    GreaterThanOrEqual,
    LessThan,
    LessThanOrEqual,
}

/// Error alert style shown by Excel when validation fails.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum DataValidationErrorStyle {
    Stop,
    Warning,
    Information,
}

impl Default for DataValidationErrorStyle {
    fn default() -> Self {
        DataValidationErrorStyle::Stop
    }
}

/// Input message shown when a validated cell is selected.
#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
pub struct DataValidationInputMessage {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub body: Option<String>,
}

/// Error alert configuration shown when a validated cell rejects a value.
#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
pub struct DataValidationErrorAlert {
    #[serde(default)]
    pub style: DataValidationErrorStyle,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub body: Option<String>,
}

/// An Excel-style data validation rule.
///
/// `formula1` / `formula2` are stored without a leading `=`.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct DataValidation {
    pub kind: DataValidationKind,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub operator: Option<DataValidationOperator>,
    #[serde(default, skip_serializing_if = "String::is_empty")]
    pub formula1: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub formula2: Option<String>,

    #[serde(default, skip_serializing_if = "is_false")]
    pub allow_blank: bool,
    #[serde(default, skip_serializing_if = "is_false")]
    pub show_input_message: bool,
    #[serde(default, skip_serializing_if = "is_false")]
    pub show_error_message: bool,
    /// For list validations, whether to show the in-cell dropdown arrow (Excel UI: "In-cell
    /// dropdown").
    ///
    /// Note: SpreadsheetML (XLSX) uses the inverted `showDropDown` attribute:
    /// `showDropDown="1"` means "suppress (hide) the in-cell dropdown arrow". This field stores
    /// the UI-facing behavior (`true` = show).
    #[serde(default, skip_serializing_if = "is_false")]
    pub show_drop_down: bool,

    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub input_message: Option<DataValidationInputMessage>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub error_alert: Option<DataValidationErrorAlert>,
}

impl DataValidation {
    pub(crate) fn rewrite_sheet_references(&mut self, old_name: &str, new_name: &str) {
        // For list validations, `formula1` can contain a literal list (e.g. `"A,B,C"`). Preserve
        // literal lists unchanged while still rewriting formula-like list sources.
        let formula1_is_literal_list =
            self.kind == DataValidationKind::List && parse_list_constant(&self.formula1).is_some();

        if !formula1_is_literal_list && !self.formula1.is_empty() {
            self.formula1 =
                crate::rewrite_sheet_names_in_formula(&self.formula1, old_name, new_name);
        }

        if let Some(formula2) = self.formula2.as_mut() {
            *formula2 = crate::rewrite_sheet_names_in_formula(formula2, old_name, new_name);
        }
    }

    pub(crate) fn rewrite_sheet_references_internal_refs_only(
        &mut self,
        old_name: &str,
        new_name: &str,
    ) {
        // For list validations, `formula1` can contain a literal list (e.g. `"A,B,C"`). Preserve
        // literal lists unchanged while still rewriting formula-like list sources.
        let formula1_is_literal_list =
            self.kind == DataValidationKind::List && parse_list_constant(&self.formula1).is_some();

        if !formula1_is_literal_list && !self.formula1.is_empty() {
            self.formula1 =
                crate::formula_rewrite::rewrite_sheet_names_in_formula_internal_refs_only(
                    &self.formula1,
                    old_name,
                    new_name,
                );
        }

        if let Some(formula2) = self.formula2.as_mut() {
            *formula2 = crate::formula_rewrite::rewrite_sheet_names_in_formula_internal_refs_only(
                formula2, old_name, new_name,
            );
        }
    }

    pub(crate) fn rewrite_table_references(&mut self, renames: &[(String, String)]) {
        let formula1_is_literal_list =
            self.kind == DataValidationKind::List && parse_list_constant(&self.formula1).is_some();

        if !formula1_is_literal_list && !self.formula1.is_empty() {
            self.formula1 = crate::rewrite_table_names_in_formula(&self.formula1, renames);
        }

        if let Some(formula2) = self.formula2.as_mut() {
            *formula2 = crate::rewrite_table_names_in_formula(formula2, renames);
        }
    }

    pub(crate) fn invalidate_deleted_sheet_references(
        &mut self,
        deleted_sheet: &str,
        sheet_order: &[String],
    ) {
        // For list validations, `formula1` can contain a literal list (e.g. `"A,B,C"`). Preserve
        // literal lists unchanged while still rewriting formula-like list sources.
        let formula1_is_literal_list =
            self.kind == DataValidationKind::List && parse_list_constant(&self.formula1).is_some();

        if !formula1_is_literal_list && !self.formula1.is_empty() {
            self.formula1 = crate::rewrite_deleted_sheet_references_in_formula(
                &self.formula1,
                deleted_sheet,
                sheet_order,
            );
        }

        if let Some(formula2) = self.formula2.as_mut() {
            *formula2 = crate::rewrite_deleted_sheet_references_in_formula(
                formula2,
                deleted_sheet,
                sheet_order,
            );
        }
    }
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct DataValidationAssignment {
    pub id: DataValidationId,
    #[serde(default)]
    pub ranges: Vec<Range>,
    pub validation: DataValidation,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum DataValidationErrorKind {
    BlankNotAllowed,
    TypeMismatch,
    NotWholeNumber,
    ComparisonFailed,
    NotInList,
    UnresolvedListSource,
    UnsupportedFormula,
    CustomFormulaFalse,
    CustomFormulaUnresolved,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DataValidationResult {
    pub ok: bool,
    pub error_kind: Option<DataValidationErrorKind>,
    pub error_style: Option<DataValidationErrorStyle>,
    pub error_title: Option<String>,
    pub error_message: Option<String>,
}

impl DataValidationResult {
    pub fn ok() -> Self {
        Self {
            ok: true,
            error_kind: None,
            error_style: None,
            error_title: None,
            error_message: None,
        }
    }

    fn fail(validation: &DataValidation, kind: DataValidationErrorKind) -> Self {
        if !validation.show_error_message {
            return Self {
                ok: false,
                error_kind: Some(kind),
                error_style: None,
                error_title: None,
                error_message: None,
            };
        }

        let style = validation
            .error_alert
            .as_ref()
            .map(|a| a.style)
            .unwrap_or_default();
        let title = validation
            .error_alert
            .as_ref()
            .and_then(|a| a.title.clone())
            .or_else(|| Some("Invalid value".to_string()));
        let message = validation
            .error_alert
            .as_ref()
            .and_then(|a| a.body.clone())
            .or_else(|| Some(default_error_message(kind).to_string()));

        Self {
            ok: false,
            error_kind: Some(kind),
            error_style: Some(style),
            error_title: title,
            error_message: message,
        }
    }
}

fn default_error_message(kind: DataValidationErrorKind) -> &'static str {
    match kind {
        DataValidationErrorKind::BlankNotAllowed => "This cell cannot be left blank.",
        DataValidationErrorKind::TypeMismatch => "The value does not match the required type.",
        DataValidationErrorKind::NotWholeNumber => "The value must be a whole number.",
        DataValidationErrorKind::ComparisonFailed => {
            "The value does not satisfy the validation rule."
        }
        DataValidationErrorKind::NotInList => "The value is not in the list of allowed values.",
        DataValidationErrorKind::UnresolvedListSource => {
            "The validation list source could not be resolved."
        }
        DataValidationErrorKind::UnsupportedFormula => {
            "The validation rule uses an unsupported formula."
        }
        DataValidationErrorKind::CustomFormulaFalse => {
            "The custom validation formula evaluated to FALSE."
        }
        DataValidationErrorKind::CustomFormulaUnresolved => {
            "The custom validation formula could not be evaluated."
        }
    }
}

/// Hooks used by the data validation evaluator.
pub trait DataValidationContext {
    /// Resolve a list source formula (e.g. `A1:A5` or `MyNamedRange`) to allowed values.
    fn resolve_list_source(&self, _formula: &str) -> Option<Vec<String>> {
        None
    }

    /// Evaluate a custom validation formula, returning whether the candidate passes.
    fn eval_custom_formula(&self, _formula: &str, _candidate: &CellValue) -> Option<bool> {
        None
    }
}

impl DataValidationContext for () {}

pub fn validate_value(
    validation: &DataValidation,
    candidate: &CellValue,
    ctx: &impl DataValidationContext,
) -> DataValidationResult {
    if is_blank(candidate) {
        return if validation.allow_blank {
            DataValidationResult::ok()
        } else {
            DataValidationResult::fail(validation, DataValidationErrorKind::BlankNotAllowed)
        };
    }

    match validation.kind {
        DataValidationKind::Whole => {
            let Some(n) = coerce_number(candidate) else {
                return DataValidationResult::fail(
                    validation,
                    DataValidationErrorKind::TypeMismatch,
                );
            };
            if !is_effectively_integer(n) {
                return DataValidationResult::fail(
                    validation,
                    DataValidationErrorKind::NotWholeNumber,
                );
            }
            validate_with_operator(validation, n)
        }
        DataValidationKind::Decimal => {
            let Some(n) = coerce_number(candidate) else {
                return DataValidationResult::fail(
                    validation,
                    DataValidationErrorKind::TypeMismatch,
                );
            };
            validate_with_operator(validation, n)
        }
        DataValidationKind::List => {
            let Some(text) = coerce_text(candidate) else {
                return DataValidationResult::fail(
                    validation,
                    DataValidationErrorKind::TypeMismatch,
                );
            };
            let allowed = if let Some(list) = parse_list_constant(&validation.formula1) {
                list
            } else {
                ctx.resolve_list_source(normalize_formula(&validation.formula1))
                    .ok_or(())
                    .unwrap_or_else(|_| Vec::new())
            };

            if allowed.is_empty() {
                return DataValidationResult::fail(
                    validation,
                    if parse_list_constant(&validation.formula1).is_some() {
                        DataValidationErrorKind::UnsupportedFormula
                    } else {
                        DataValidationErrorKind::UnresolvedListSource
                    },
                );
            }

            if allowed
                .iter()
                .any(|v| text_eq_case_insensitive(v, text.trim()))
            {
                DataValidationResult::ok()
            } else {
                DataValidationResult::fail(validation, DataValidationErrorKind::NotInList)
            }
        }
        DataValidationKind::Date => {
            let Some(serial) = coerce_date_serial(candidate) else {
                return DataValidationResult::fail(
                    validation,
                    DataValidationErrorKind::TypeMismatch,
                );
            };
            validate_with_operator(validation, serial)
        }
        DataValidationKind::Time => {
            let Some(serial) = coerce_time_fraction(candidate) else {
                return DataValidationResult::fail(
                    validation,
                    DataValidationErrorKind::TypeMismatch,
                );
            };
            validate_with_operator(validation, serial)
        }
        DataValidationKind::TextLength => {
            let Some(text) = coerce_text(candidate) else {
                return DataValidationResult::fail(
                    validation,
                    DataValidationErrorKind::TypeMismatch,
                );
            };
            let len = excel_text_len(&text) as f64;
            validate_with_operator(validation, len)
        }
        DataValidationKind::Custom => match ctx
            .eval_custom_formula(normalize_formula(&validation.formula1), candidate)
        {
            Some(true) => DataValidationResult::ok(),
            Some(false) => {
                DataValidationResult::fail(validation, DataValidationErrorKind::CustomFormulaFalse)
            }
            None => DataValidationResult::fail(
                validation,
                DataValidationErrorKind::CustomFormulaUnresolved,
            ),
        },
    }
}

fn validate_with_operator(validation: &DataValidation, candidate: f64) -> DataValidationResult {
    let Some(op) = validation.operator else {
        return DataValidationResult::fail(validation, DataValidationErrorKind::UnsupportedFormula);
    };

    let Some(a) = parse_operand(validation, &validation.formula1) else {
        return DataValidationResult::fail(validation, DataValidationErrorKind::UnsupportedFormula);
    };

    let ok = match op {
        DataValidationOperator::Between | DataValidationOperator::NotBetween => {
            let Some(b_str) = validation.formula2.as_deref() else {
                return DataValidationResult::fail(
                    validation,
                    DataValidationErrorKind::UnsupportedFormula,
                );
            };
            let Some(b) = parse_operand(validation, b_str) else {
                return DataValidationResult::fail(
                    validation,
                    DataValidationErrorKind::UnsupportedFormula,
                );
            };
            let between = candidate >= a && candidate <= b;
            if matches!(op, DataValidationOperator::Between) {
                between
            } else {
                !between
            }
        }
        DataValidationOperator::Equal => approx_eq(candidate, a),
        DataValidationOperator::NotEqual => !approx_eq(candidate, a),
        DataValidationOperator::GreaterThan => candidate > a,
        DataValidationOperator::GreaterThanOrEqual => candidate >= a || approx_eq(candidate, a),
        DataValidationOperator::LessThan => candidate < a,
        DataValidationOperator::LessThanOrEqual => candidate <= a || approx_eq(candidate, a),
    };

    if ok {
        DataValidationResult::ok()
    } else {
        DataValidationResult::fail(validation, DataValidationErrorKind::ComparisonFailed)
    }
}

fn approx_eq(a: f64, b: f64) -> bool {
    (a - b).abs() <= 1e-9
}

fn is_blank(candidate: &CellValue) -> bool {
    match candidate {
        CellValue::Empty => true,
        CellValue::String(s) => s.is_empty(),
        CellValue::RichText(rt) => rt.plain_text().is_empty(),
        CellValue::Entity(e) => e.display_value.is_empty(),
        CellValue::Record(r) => {
            if let Some(field) = r.display_field.as_deref() {
                if let Some(value) = r.get_field_case_insensitive(field) {
                    return match value {
                        CellValue::Empty => true,
                        CellValue::String(s) => s.is_empty(),
                        CellValue::RichText(rt) => rt.text.is_empty(),
                        // Scalar display strings for these variants are always non-empty.
                        CellValue::Number(_) | CellValue::Boolean(_) | CellValue::Error(_) => false,
                        CellValue::Entity(entity) => entity.display_value.is_empty(),
                        CellValue::Record(record) => record.to_string().is_empty(),
                        CellValue::Image(_) => false,
                        // Non-scalar displayField values fall back to `display_value`.
                        _ => r.display_value.is_empty(),
                    };
                }
            }
            r.display_value.is_empty()
        }
        _ => false,
    }
}

fn normalize_formula(formula: &str) -> &str {
    let s = formula.trim();
    s.strip_prefix('=').unwrap_or(s).trim()
}

fn parse_excel_string_literal(formula: &str) -> Option<String> {
    let s = normalize_formula(formula);
    if s.len() < 2 {
        return None;
    }
    let bytes = s.as_bytes();
    if bytes.first() != Some(&b'"') || bytes.last() != Some(&b'"') {
        return None;
    }
    let inner = &s[1..s.len() - 1];
    Some(inner.replace("\"\"", "\""))
}

fn parse_number_like_literal(formula: &str) -> Option<f64> {
    let s = normalize_formula(formula);
    if let Some(lit) = parse_excel_string_literal(s) {
        lit.trim().parse::<f64>().ok()
    } else {
        s.parse::<f64>().ok()
    }
}

fn parse_operand(validation: &DataValidation, formula: &str) -> Option<f64> {
    match validation.kind {
        DataValidationKind::Date => parse_date_operand(formula),
        DataValidationKind::Time => parse_time_operand(formula),
        _ => parse_number_like_literal(formula),
    }
}

fn parse_date_operand(formula: &str) -> Option<f64> {
    let s = normalize_formula(formula);
    if let Ok(n) = s.parse::<f64>() {
        return Some(n);
    }
    if let Some(lit) = parse_excel_string_literal(s) {
        return parse_date_time_serial(lit.trim());
    }
    parse_date_time_serial(s)
}

fn parse_time_operand(formula: &str) -> Option<f64> {
    let s = normalize_formula(formula);
    if let Ok(n) = s.parse::<f64>() {
        return Some(n.rem_euclid(1.0));
    }
    if let Some(lit) = parse_excel_string_literal(s) {
        return parse_time_fraction(lit.trim());
    }
    parse_time_fraction(s)
}

fn parse_list_constant(formula: &str) -> Option<Vec<String>> {
    let s = normalize_formula(formula);
    let list_str = if let Some(lit) = parse_excel_string_literal(s) {
        lit
    } else if s.contains(',') || s.contains(';') {
        s.to_string()
    } else {
        return None;
    };

    let delimiter = if list_str.contains(';') && !list_str.contains(',') {
        ';'
    } else {
        ','
    };

    let items: Vec<String> = list_str
        .split(delimiter)
        .map(|item| item.trim().to_string())
        .filter(|s| !s.is_empty())
        .collect();
    Some(items)
}

fn coerce_number(value: &CellValue) -> Option<f64> {
    match value {
        CellValue::Number(n) => Some(*n),
        CellValue::String(s) => s.trim().parse::<f64>().ok(),
        CellValue::Boolean(b) => Some(if *b { 1.0 } else { 0.0 }),
        CellValue::RichText(rt) => rt.plain_text().trim().parse::<f64>().ok(),
        // Rich values are treated as text-like for data validation; attempt to parse the display
        // string using the same coercion rules as plain strings.
        CellValue::Entity(e) => e.display_value.trim().parse::<f64>().ok(),
        CellValue::Record(r) => r.to_string().trim().parse::<f64>().ok(),
        CellValue::Image(image) => image
            .alt_text
            .as_deref()
            .filter(|s| !s.is_empty())
            .unwrap_or("[Image]")
            .trim()
            .parse::<f64>()
            .ok(),
        _ => None,
    }
}

fn is_effectively_integer(value: f64) -> bool {
    let rounded = value.round();
    approx_eq(value, rounded)
}

fn coerce_text(value: &CellValue) -> Option<String> {
    match value {
        CellValue::String(s) => Some(s.clone()),
        CellValue::RichText(rt) => Some(rt.plain_text().to_string()),
        CellValue::Number(n) => Some(number_to_string(*n)),
        CellValue::Boolean(b) => Some(if *b {
            "TRUE".to_string()
        } else {
            "FALSE".to_string()
        }),
        CellValue::Entity(e) => Some(e.display_value.clone()),
        CellValue::Record(r) => Some(r.to_string()),
        CellValue::Image(image) => Some(
            image
                .alt_text
                .as_deref()
                .filter(|s| !s.is_empty())
                .unwrap_or("[Image]")
                .to_string(),
        ),
        _ => None,
    }
}

fn number_to_string(n: f64) -> String {
    if is_effectively_integer(n) {
        format!("{}", n.round() as i64)
    } else {
        n.to_string()
    }
}

fn coerce_date_serial(value: &CellValue) -> Option<f64> {
    match value {
        CellValue::Number(n) => Some(*n),
        CellValue::String(s) => parse_date_time_serial(s.trim()),
        CellValue::RichText(rt) => parse_date_time_serial(rt.plain_text().trim()),
        // Rich values are treated as text-like for data validation; attempt to parse the display
        // string using the same coercion rules as plain strings.
        CellValue::Entity(e) => parse_date_time_serial(e.display_value.trim()),
        CellValue::Record(r) => parse_date_time_serial(r.to_string().trim()),
        CellValue::Image(image) => parse_date_time_serial(
            image
                .alt_text
                .as_deref()
                .filter(|s| !s.is_empty())
                .unwrap_or("[Image]")
                .trim(),
        ),
        _ => None,
    }
}

fn coerce_time_fraction(value: &CellValue) -> Option<f64> {
    match value {
        CellValue::Number(n) => Some(n.rem_euclid(1.0)),
        CellValue::String(s) => parse_time_fraction(s.trim()),
        CellValue::RichText(rt) => parse_time_fraction(rt.plain_text().trim()),
        // Rich values are treated as text-like for data validation; attempt to parse the display
        // string using the same coercion rules as plain strings.
        CellValue::Entity(e) => parse_time_fraction(e.display_value.trim()),
        CellValue::Record(r) => parse_time_fraction(r.to_string().trim()),
        CellValue::Image(image) => parse_time_fraction(
            image
                .alt_text
                .as_deref()
                .filter(|s| !s.is_empty())
                .unwrap_or("[Image]")
                .trim(),
        ),
        _ => None,
    }
}

fn parse_date_time_serial(s: &str) -> Option<f64> {
    use chrono::{NaiveDate, NaiveDateTime};

    if s.is_empty() {
        return None;
    }

    if let Ok(date) = NaiveDate::parse_from_str(s, "%Y-%m-%d") {
        return Some(date_to_excel_serial(date));
    }
    if let Ok(date) = NaiveDate::parse_from_str(s, "%Y/%m/%d") {
        return Some(date_to_excel_serial(date));
    }
    if let Ok(dt) = NaiveDateTime::parse_from_str(s, "%Y-%m-%dT%H:%M:%S") {
        return Some(date_to_excel_serial(dt.date()) + time_to_excel_fraction(dt.time()));
    }
    if let Ok(dt) = NaiveDateTime::parse_from_str(s, "%Y-%m-%d %H:%M:%S") {
        return Some(date_to_excel_serial(dt.date()) + time_to_excel_fraction(dt.time()));
    }

    None
}

fn parse_time_fraction(s: &str) -> Option<f64> {
    use chrono::NaiveTime;

    if s.is_empty() {
        return None;
    }

    if let Ok(time) = NaiveTime::parse_from_str(s, "%H:%M:%S") {
        return Some(time_to_excel_fraction(time));
    }
    if let Ok(time) = NaiveTime::parse_from_str(s, "%H:%M") {
        return Some(time_to_excel_fraction(time));
    }

    None
}

fn time_to_excel_fraction(time: chrono::NaiveTime) -> f64 {
    use chrono::Timelike as _;
    let seconds = time.num_seconds_from_midnight() as f64;
    let nanos = time.nanosecond() as f64;
    (seconds + nanos / 1_000_000_000.0) / 86_400.0
}

fn date_to_excel_serial(date: chrono::NaiveDate) -> f64 {
    let base = chrono::NaiveDate::from_ymd_opt(1899, 12, 31).expect("valid base date");
    let days = (date - base).num_days() as f64;
    let leap_bug_cutover =
        chrono::NaiveDate::from_ymd_opt(1900, 3, 1).expect("valid leap bug cutover");
    if date >= leap_bug_cutover {
        days + 1.0
    } else {
        days
    }
}

/// Excel's text length semantics for data validation rules.
///
/// Excel's `LEN` / text-length validation counts UTF-16 code units, which means
/// characters outside the BMP (e.g. 🙂) count as 2.
fn excel_text_len(s: &str) -> usize {
    s.encode_utf16().count()
}
