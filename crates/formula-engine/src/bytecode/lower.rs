use super::ast::{BinaryOp, Expr as BytecodeExpr, Function, UnaryOp};
use super::value::{
    Array, ErrorKind as BytecodeErrorKind, MultiRangeRef, RangeRef, Ref, SheetId, SheetRangeRef,
    Value,
};
use crate::value::{casefolded_key_arc, casefolded_key_arc_if};
use formula_model::{EXCEL_MAX_COLS, EXCEL_MAX_ROWS};
use std::collections::HashSet;
use std::sync::Arc;

#[derive(thiserror::Error, Debug, Clone, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub enum LowerError {
    #[error("unsupported expression")]
    Unsupported,
    /// External workbook reference that the bytecode lowering layer cannot represent.
    ///
    /// This indicates an external workbook reference shape that the bytecode backend does not
    /// currently implement.
    ///
    /// Notably, some external workbook reference shapes require runtime expansion using external
    /// workbook tab order.
    #[error("unsupported external workbook reference")]
    ExternalReference,
    #[error("cross-sheet references are not supported")]
    CrossSheetReference,
    #[error("unknown sheet reference")]
    UnknownSheet,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct RefPrefix {
    workbook: Option<String>,
    sheet: Option<crate::SheetRef>,
}

impl RefPrefix {
    fn is_unprefixed(&self) -> bool {
        self.workbook.is_none() && self.sheet.is_none()
    }

    fn from_parts(workbook: &Option<String>, sheet: &Option<crate::SheetRef>) -> Self {
        Self {
            workbook: workbook.clone(),
            sheet: sheet.clone(),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum RectKind {
    Cell,
    /// Whole-column reference (`A:A`) which spans all rows.
    Col,
    /// Whole-row reference (`1:1`) which spans all columns.
    Row,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct RectRef {
    kind: RectKind,
    row: i32,
    col: i32,
    row_abs: bool,
    col_abs: bool,
}

impl RectRef {
    fn spans_full_rows(&self) -> bool {
        matches!(self.kind, RectKind::Col)
    }

    fn spans_full_cols(&self) -> bool {
        matches!(self.kind, RectKind::Row)
    }

    fn start_on_sheet(&self, _max_row: i32, _max_col: i32) -> Ref {
        match self.kind {
            RectKind::Cell => Ref::new(self.row, self.col, self.row_abs, self.col_abs),
            RectKind::Col => Ref::new(0, self.col, true, self.col_abs),
            RectKind::Row => Ref::new(self.row, 0, self.row_abs, true),
        }
    }

    fn end_on_sheet(&self, max_row: i32, max_col: i32) -> Ref {
        match self.kind {
            RectKind::Cell => Ref::new(self.row, self.col, self.row_abs, self.col_abs),
            RectKind::Col => Ref::new(max_row, self.col, true, self.col_abs),
            RectKind::Row => Ref::new(self.row, max_col, self.row_abs, true),
        }
    }
}

fn sheet_max_indices(
    sheet: &SheetId,
    sheet_dimensions: &mut impl FnMut(usize) -> Option<(u32, u32)>,
) -> Result<(i32, i32), LowerError> {
    let (rows, cols) = match sheet {
        SheetId::Local(id) => sheet_dimensions(*id).unwrap_or((EXCEL_MAX_ROWS, EXCEL_MAX_COLS)),
        SheetId::External(_) | SheetId::ExternalSpan { .. } => (EXCEL_MAX_ROWS, EXCEL_MAX_COLS),
    };
    let rows = i32::try_from(rows).map_err(|_| LowerError::Unsupported)?;
    let cols = i32::try_from(cols).map_err(|_| LowerError::Unsupported)?;
    let max_row = rows.checked_sub(1).ok_or(LowerError::Unsupported)?;
    let max_col = cols.checked_sub(1).ok_or(LowerError::Unsupported)?;
    Ok((max_row, max_col))
}

fn validate_prefix(
    prefix: &RefPrefix,
    _current_sheet: usize,
    resolve_sheet_id: &mut impl FnMut(&str) -> Option<usize>,
) -> Result<(), LowerError> {
    if prefix.workbook.is_some() {
        // External workbook reference.
        // Validate that the external sheet key/span is representable.
        let _ = external_sheet_id(prefix)?;
        return Ok(());
    }
    if let Some(sheet) = prefix.sheet.as_ref() {
        match sheet {
            crate::SheetRef::Sheet(name) => {
                if resolve_sheet_id(name).is_none() {
                    return Err(LowerError::UnknownSheet);
                };
            }
            crate::SheetRef::SheetRange { start, end } => {
                // Sheet-span references are allowed in the bytecode backend (lowered as a
                // multi-area reference). We still validate that both sheet names resolve.
                if resolve_sheet_id(start).is_none() || resolve_sheet_id(end).is_none() {
                    return Err(LowerError::UnknownSheet);
                }
            }
        }
    }
    Ok(())
}

fn external_sheet_id(prefix: &RefPrefix) -> Result<SheetId, LowerError> {
    let Some(book) = prefix.workbook.as_ref() else {
        return Err(LowerError::ExternalReference);
    };

    // Mirror `eval::compiler::lower_sheet_reference` so we build the same canonical key string
    // that `ValueResolver::get_external_value` expects.
    match prefix.sheet.as_ref() {
        Some(crate::SheetRef::Sheet(sheet)) => {
            let key = crate::external_refs::format_external_key(book, sheet);
            if !crate::eval::is_valid_external_single_sheet_key(&key) {
                return Err(LowerError::ExternalReference);
            }
            Ok(SheetId::External(Arc::from(key)))
        }
        Some(crate::SheetRef::SheetRange { start, end }) => {
            if formula_model::sheet_name_eq_case_insensitive(start, end) {
                let key = crate::external_refs::format_external_key(book, start);
                if !crate::eval::is_valid_external_single_sheet_key(&key) {
                    return Err(LowerError::ExternalReference);
                }
                Ok(SheetId::External(Arc::from(key)))
            } else {
                if start.is_empty() || end.is_empty() || start.contains(':') || end.contains(':') {
                    return Err(LowerError::ExternalReference);
                }
                Ok(SheetId::ExternalSpan {
                    workbook: Arc::from(book.as_str()),
                    start: Arc::from(start.as_str()),
                    end: Arc::from(end.as_str()),
                })
            }
        }
        None => Err(LowerError::ExternalReference),
    }
}

fn expand_sheet_span(
    start: &str,
    end: &str,
    resolve_sheet_id: &mut impl FnMut(&str) -> Option<usize>,
    expand_sheet_span_ids: &mut impl FnMut(usize, usize) -> Option<Vec<usize>>,
) -> Result<Vec<usize>, LowerError> {
    let Some(a) = resolve_sheet_id(start) else {
        return Err(LowerError::UnknownSheet);
    };
    let Some(b) = resolve_sheet_id(end) else {
        return Err(LowerError::UnknownSheet);
    };
    expand_sheet_span_ids(a, b).ok_or(LowerError::UnknownSheet)
}

fn lower_coord(coord: &crate::Coord, origin: u32) -> Result<(i32, bool), LowerError> {
    match coord {
        crate::Coord::A1 { index, abs } => {
            let idx = i32::try_from(*index).map_err(|_| LowerError::Unsupported)?;
            if *abs {
                Ok((idx, true))
            } else {
                let origin = i32::try_from(origin).map_err(|_| LowerError::Unsupported)?;
                Ok((idx - origin, false))
            }
        }
        crate::Coord::Offset(delta) => Ok((*delta, false)),
    }
}

fn lower_cell_ref(
    r: &crate::CellRef,
    origin: crate::CellAddr,
    current_sheet: usize,
    resolve_sheet_id: &mut impl FnMut(&str) -> Option<usize>,
) -> Result<Ref, LowerError> {
    validate_prefix(
        &RefPrefix::from_parts(&r.workbook, &r.sheet),
        current_sheet,
        resolve_sheet_id,
    )?;
    let (row, row_abs) = lower_coord(&r.row, origin.row)?;
    let (col, col_abs) = lower_coord(&r.col, origin.col)?;
    Ok(Ref::new(row, col, row_abs, col_abs))
}

fn lower_cell_ref_expr(
    r: &crate::CellRef,
    origin: crate::CellAddr,
    current_sheet: usize,
    resolve_sheet_id: &mut impl FnMut(&str) -> Option<usize>,
    expand_sheet_span_ids: &mut impl FnMut(usize, usize) -> Option<Vec<usize>>,
) -> Result<BytecodeExpr, LowerError> {
    let prefix = RefPrefix::from_parts(&r.workbook, &r.sheet);
    let cell = lower_cell_ref(r, origin, current_sheet, resolve_sheet_id)?;

    if prefix.workbook.is_some() {
        let sheet = external_sheet_id(&prefix)?;
        let range = RangeRef::new(cell, cell);
        return Ok(BytecodeExpr::MultiRangeRef(MultiRangeRef::new(
            vec![SheetRangeRef::new(sheet, range)].into(),
        )));
    }

    match prefix.sheet.as_ref() {
        None => Ok(BytecodeExpr::CellRef(cell)),
        Some(crate::SheetRef::Sheet(name)) => {
            let Some(sheet_id) = resolve_sheet_id(name) else {
                return Err(LowerError::UnknownSheet);
            };
            if sheet_id == current_sheet {
                Ok(BytecodeExpr::CellRef(cell))
            } else {
                let range = RangeRef::new(cell, cell);
                let areas = vec![SheetRangeRef::new(SheetId::Local(sheet_id), range)];
                Ok(BytecodeExpr::MultiRangeRef(MultiRangeRef::new(
                    areas.into(),
                )))
            }
        }
        Some(crate::SheetRef::SheetRange { start, end }) => {
            let sheets = expand_sheet_span(start, end, resolve_sheet_id, expand_sheet_span_ids)?;
            let range = RangeRef::new(cell, cell);
            let areas: Vec<SheetRangeRef> = sheets
                .into_iter()
                .map(|sheet| SheetRangeRef::new(SheetId::Local(sheet), range))
                .collect();
            Ok(BytecodeExpr::MultiRangeRef(MultiRangeRef::new(
                areas.into(),
            )))
        }
    }
}

fn lower_rect_ref(
    expr: &crate::Expr,
    origin: crate::CellAddr,
    current_sheet: usize,
    resolve_sheet_id: &mut impl FnMut(&str) -> Option<usize>,
) -> Result<(RefPrefix, RectRef), LowerError> {
    match expr {
        crate::Expr::CellRef(r) => {
            let prefix = RefPrefix::from_parts(&r.workbook, &r.sheet);
            let r = lower_cell_ref(r, origin, current_sheet, resolve_sheet_id)?;
            Ok((
                prefix,
                RectRef {
                    kind: RectKind::Cell,
                    row: r.row,
                    col: r.col,
                    row_abs: r.row_abs,
                    col_abs: r.col_abs,
                },
            ))
        }
        crate::Expr::ColRef(r) => {
            let prefix = RefPrefix::from_parts(&r.workbook, &r.sheet);
            validate_prefix(&prefix, current_sheet, resolve_sheet_id)?;

            let (col, col_abs) = lower_coord(&r.col, origin.col)?;
            Ok((
                prefix,
                RectRef {
                    kind: RectKind::Col,
                    row: 0,
                    col,
                    row_abs: true,
                    col_abs,
                },
            ))
        }
        crate::Expr::RowRef(r) => {
            let prefix = RefPrefix::from_parts(&r.workbook, &r.sheet);
            validate_prefix(&prefix, current_sheet, resolve_sheet_id)?;

            let (row, row_abs) = lower_coord(&r.row, origin.row)?;
            Ok((
                prefix,
                RectRef {
                    kind: RectKind::Row,
                    row,
                    col: 0,
                    row_abs,
                    col_abs: true,
                },
            ))
        }
        _ => Err(LowerError::Unsupported),
    }
}

fn merge_range_prefix(left: &RefPrefix, right: &RefPrefix) -> Result<RefPrefix, LowerError> {
    if left == right {
        return Ok(left.clone());
    }
    if left.is_unprefixed() && !right.is_unprefixed() {
        return Ok(right.clone());
    }
    if right.is_unprefixed() && !left.is_unprefixed() {
        return Ok(left.clone());
    }
    Err(LowerError::Unsupported)
}

fn lower_range_ref(
    left: &crate::Expr,
    right: &crate::Expr,
    origin: crate::CellAddr,
    current_sheet: usize,
    resolve_sheet_id: &mut impl FnMut(&str) -> Option<usize>,
    expand_sheet_span_ids: &mut impl FnMut(usize, usize) -> Option<Vec<usize>>,
    sheet_dimensions: &mut impl FnMut(usize) -> Option<(u32, u32)>,
) -> Result<BytecodeExpr, LowerError> {
    let (left_prefix, left_rect) = lower_rect_ref(left, origin, current_sheet, resolve_sheet_id)?;
    let (right_prefix, right_rect) =
        lower_rect_ref(right, origin, current_sheet, resolve_sheet_id)?;
    let merged_prefix = merge_range_prefix(&left_prefix, &right_prefix)?;
    validate_prefix(&merged_prefix, current_sheet, resolve_sheet_id)?;

    let full_rows = left_rect.spans_full_rows() || right_rect.spans_full_rows();
    let full_cols = left_rect.spans_full_cols() || right_rect.spans_full_cols();

    let mut build_range = |sheet: SheetId| -> Result<RangeRef, LowerError> {
        let (max_row, max_col) = sheet_max_indices(&sheet, sheet_dimensions)?;
        let start_ref = left_rect.start_on_sheet(max_row, max_col);
        let end_ref = right_rect.end_on_sheet(max_row, max_col);

        let (start_row, start_row_abs) = if full_rows {
            (0, true)
        } else {
            (start_ref.row, start_ref.row_abs)
        };
        let (end_row, end_row_abs) = if full_rows {
            (max_row, true)
        } else {
            (end_ref.row, end_ref.row_abs)
        };
        let (start_col, start_col_abs) = if full_cols {
            (0, true)
        } else {
            (start_ref.col, start_ref.col_abs)
        };
        let (end_col, end_col_abs) = if full_cols {
            (max_col, true)
        } else {
            (end_ref.col, end_ref.col_abs)
        };

        Ok(RangeRef::new(
            Ref::new(start_row, start_col, start_row_abs, start_col_abs),
            Ref::new(end_row, end_col, end_row_abs, end_col_abs),
        ))
    };

    if merged_prefix.workbook.is_some() {
        let sheet = external_sheet_id(&merged_prefix)?;
        let range = build_range(sheet.clone())?;
        return Ok(BytecodeExpr::MultiRangeRef(MultiRangeRef::new(
            vec![SheetRangeRef::new(sheet, range)].into(),
        )));
    }

    match merged_prefix.sheet.as_ref() {
        Some(crate::SheetRef::Sheet(name)) => {
            let Some(sheet_id) = resolve_sheet_id(name) else {
                return Err(LowerError::UnknownSheet);
            };
            let sheet = SheetId::Local(sheet_id);
            let range = build_range(sheet.clone())?;
            if sheet_id == current_sheet {
                Ok(BytecodeExpr::RangeRef(range))
            } else {
                let areas = vec![SheetRangeRef::new(sheet, range)];
                Ok(BytecodeExpr::MultiRangeRef(MultiRangeRef::new(
                    areas.into(),
                )))
            }
        }
        Some(crate::SheetRef::SheetRange { start, end }) => {
            let sheets = expand_sheet_span(start, end, resolve_sheet_id, expand_sheet_span_ids)?;
            let mut areas: Vec<SheetRangeRef> = Vec::with_capacity(sheets.len());
            for sheet_id in sheets {
                let sheet = SheetId::Local(sheet_id);
                let range = build_range(sheet.clone())?;
                areas.push(SheetRangeRef::new(sheet, range));
            }
            Ok(BytecodeExpr::MultiRangeRef(MultiRangeRef::new(
                areas.into(),
            )))
        }
        _ => {
            let range = build_range(SheetId::Local(current_sheet))?;
            Ok(BytecodeExpr::RangeRef(range))
        }
    }
}

fn parse_number(raw: &str) -> Result<f64, LowerError> {
    match raw.parse::<f64>() {
        Ok(n) if n.is_finite() => Ok(n),
        _ => Err(LowerError::Unsupported),
    }
}

fn parse_error_kind(raw: &str) -> BytecodeErrorKind {
    // Keep this in sync with `eval::compiler::parse_error_kind` so AST and bytecode evaluation
    // agree on the canonical set of supported error literals.
    BytecodeErrorKind::from_code(raw).unwrap_or(BytecodeErrorKind::Value)
}

fn lower_array_literal_element(expr: &crate::Expr) -> Result<Value, LowerError> {
    match expr {
        crate::Expr::Number(raw) => Ok(Value::Number(parse_number(raw)?)),
        crate::Expr::String(s) => Ok(Value::Text(Arc::from(s.as_str()))),
        crate::Expr::Boolean(b) => Ok(Value::Bool(*b)),
        crate::Expr::Error(raw) => Ok(Value::Error(BytecodeErrorKind::from(parse_error_kind(raw)))),
        crate::Expr::Missing => Ok(Value::Empty),
        crate::Expr::Unary(u) => match u.op {
            crate::UnaryOp::Plus => match lower_array_literal_element(&u.expr)? {
                Value::Number(n) => Ok(Value::Number(n)),
                _ => Err(LowerError::Unsupported),
            },
            crate::UnaryOp::Minus => match lower_array_literal_element(&u.expr)? {
                Value::Number(n) => Ok(Value::Number(-n)),
                _ => Err(LowerError::Unsupported),
            },
            crate::UnaryOp::ImplicitIntersection => Err(LowerError::Unsupported),
        },
        // Reject any non-literal element (e.g. references or function calls).
        _ => Err(LowerError::Unsupported),
    }
}

fn lower_array_literal(arr: &crate::ArrayLiteral) -> Result<Value, LowerError> {
    if arr.rows.is_empty() {
        return Err(LowerError::Unsupported);
    }
    let rows = arr.rows.len();
    let cols = arr.rows[0].len();
    if cols == 0 {
        return Err(LowerError::Unsupported);
    }
    if arr.rows.iter().any(|row| row.len() != cols) {
        return Err(LowerError::Unsupported);
    }

    let mut values = Vec::with_capacity(rows.saturating_mul(cols));
    for row in &arr.rows {
        for el in row {
            values.push(lower_array_literal_element(el)?);
        }
    }

    Ok(Value::Array(Array::new(rows, cols, values)))
}

fn collect_concat_operands<'a>(expr: &'a crate::Expr, out: &mut Vec<&'a crate::Expr>) {
    match expr {
        crate::Expr::Binary(b) if b.op == crate::BinaryOp::Concat => {
            collect_concat_operands(&b.left, out);
            collect_concat_operands(&b.right, out);
        }
        other => out.push(other),
    }
}

#[derive(Default)]
struct LexicalScopes {
    scopes: Vec<HashSet<Arc<str>>>,
}

impl LexicalScopes {
    fn push_scope(&mut self) {
        self.scopes.push(HashSet::new());
    }

    fn pop_scope(&mut self) {
        self.scopes.pop();
    }

    fn define(&mut self, key: Arc<str>) {
        if self.scopes.is_empty() {
            self.push_scope();
        }
        if let Some(scope) = self.scopes.last_mut() {
            scope.insert(key);
        }
    }

    fn is_defined(&self, key: &str) -> bool {
        self.scopes.iter().rev().any(|scope| scope.contains(key))
    }
}

fn value_error_literal() -> BytecodeExpr {
    BytecodeExpr::Literal(Value::Error(super::value::ErrorKind::Value))
}

fn bare_identifier(expr: &crate::Expr) -> Option<&str> {
    match expr {
        crate::Expr::NameRef(nref) if nref.workbook.is_none() && nref.sheet.is_none() => {
            Some(nref.name.as_str())
        }
        _ => None,
    }
}

fn lower_canonical_reference_expr(
    expr: &crate::Expr,
    origin: crate::CellAddr,
    current_sheet: usize,
    resolve_sheet_id: &mut impl FnMut(&str) -> Option<usize>,
    expand_sheet_span_ids: &mut impl FnMut(usize, usize) -> Option<Vec<usize>>,
    sheet_dimensions: &mut impl FnMut(usize) -> Option<(u32, u32)>,
    scopes: &mut LexicalScopes,
    lambda_self_name: Option<&str>,
) -> Result<BytecodeExpr, LowerError> {
    match expr {
        crate::Expr::CellRef(r) => {
            // In reference contexts (spill/union/intersect), cell references must preserve
            // reference semantics (as a single-cell range) while still respecting explicit sheet
            // prefixes (`Sheet2!A1`).
            match lower_cell_ref_expr(
                r,
                origin,
                current_sheet,
                resolve_sheet_id,
                expand_sheet_span_ids,
            )? {
                BytecodeExpr::CellRef(cell) => {
                    Ok(BytecodeExpr::RangeRef(RangeRef::new(cell, cell)))
                }
                BytecodeExpr::MultiRangeRef(r) => Ok(BytecodeExpr::MultiRangeRef(r)),
                other => unreachable!(
                    "lower_cell_ref_expr only lowers to CellRef/MultiRangeRef, got {other:?}"
                ),
            }
        }
        crate::Expr::Binary(b) if b.op == crate::BinaryOp::Range => lower_range_ref(
            &b.left,
            &b.right,
            origin,
            current_sheet,
            resolve_sheet_id,
            expand_sheet_span_ids,
            sheet_dimensions,
        ),
        crate::Expr::Binary(b)
            if matches!(b.op, crate::BinaryOp::Union | crate::BinaryOp::Intersect) =>
        {
            let op = match b.op {
                crate::BinaryOp::Union => BinaryOp::Union,
                crate::BinaryOp::Intersect => BinaryOp::Intersect,
                _ => unreachable!("guarded above"),
            };
            Ok(BytecodeExpr::Binary {
                op,
                left: Box::new(lower_canonical_reference_expr(
                    &b.left,
                    origin,
                    current_sheet,
                    resolve_sheet_id,
                    expand_sheet_span_ids,
                    sheet_dimensions,
                    scopes,
                    lambda_self_name,
                )?),
                right: Box::new(lower_canonical_reference_expr(
                    &b.right,
                    origin,
                    current_sheet,
                    resolve_sheet_id,
                    expand_sheet_span_ids,
                    sheet_dimensions,
                    scopes,
                    lambda_self_name,
                )?),
            })
        }
        crate::Expr::Postfix(p) if p.op == crate::PostfixOp::SpillRange => Ok(
            BytecodeExpr::SpillRange(Box::new(lower_canonical_reference_expr(
                &p.expr,
                origin,
                current_sheet,
                resolve_sheet_id,
                expand_sheet_span_ids,
                sheet_dimensions,
                scopes,
                lambda_self_name,
            )?)),
        ),
        // Fall back to normal lowering for non-reference expressions; runtime will surface #VALUE!.
        _ => lower_canonical_expr_inner(
            expr,
            origin,
            current_sheet,
            resolve_sheet_id,
            expand_sheet_span_ids,
            sheet_dimensions,
            scopes,
            lambda_self_name,
        ),
    }
}

pub fn lower_canonical_expr_with_sheet_span(
    expr: &crate::Expr,
    origin: crate::CellAddr,
    current_sheet: usize,
    resolve_sheet_id: &mut impl FnMut(&str) -> Option<usize>,
    expand_sheet_span_ids: &mut impl FnMut(usize, usize) -> Option<Vec<usize>>,
    sheet_dimensions: &mut impl FnMut(usize) -> Option<(u32, u32)>,
) -> Result<BytecodeExpr, LowerError> {
    let mut scopes = LexicalScopes::default();
    lower_canonical_expr_inner(
        expr,
        origin,
        current_sheet,
        resolve_sheet_id,
        expand_sheet_span_ids,
        sheet_dimensions,
        &mut scopes,
        None,
    )
}

pub fn lower_canonical_expr(
    expr: &crate::Expr,
    origin: crate::CellAddr,
    current_sheet: usize,
    resolve_sheet_id: &mut impl FnMut(&str) -> Option<usize>,
    sheet_dimensions: &mut impl FnMut(usize) -> Option<(u32, u32)>,
) -> Result<BytecodeExpr, LowerError> {
    let mut expand_numeric = |a: usize, b: usize| {
        let (start, end) = if a <= b { (a, b) } else { (b, a) };
        Some((start..=end).collect())
    };
    lower_canonical_expr_with_sheet_span(
        expr,
        origin,
        current_sheet,
        resolve_sheet_id,
        &mut expand_numeric,
        sheet_dimensions,
    )
}

fn lower_canonical_expr_inner(
    expr: &crate::Expr,
    origin: crate::CellAddr,
    current_sheet: usize,
    resolve_sheet_id: &mut impl FnMut(&str) -> Option<usize>,
    expand_sheet_span_ids: &mut impl FnMut(usize, usize) -> Option<Vec<usize>>,
    sheet_dimensions: &mut impl FnMut(usize) -> Option<(u32, u32)>,
    scopes: &mut LexicalScopes,
    lambda_self_name: Option<&str>,
) -> Result<BytecodeExpr, LowerError> {
    match expr {
        crate::Expr::Number(raw) => Ok(BytecodeExpr::Literal(Value::Number(parse_number(raw)?))),
        crate::Expr::String(s) => Ok(BytecodeExpr::Literal(Value::Text(Arc::from(s.as_str())))),
        crate::Expr::Boolean(b) => Ok(BytecodeExpr::Literal(Value::Bool(*b))),
        crate::Expr::Error(raw) => Ok(BytecodeExpr::Literal(Value::Error(parse_error_kind(raw)))),
        crate::Expr::Missing => Ok(BytecodeExpr::Literal(Value::Missing)),
        crate::Expr::Array(arr) => Ok(BytecodeExpr::Literal(lower_array_literal(arr)?)),
        crate::Expr::CellRef(r) => lower_cell_ref_expr(
            r,
            origin,
            current_sheet,
            resolve_sheet_id,
            expand_sheet_span_ids,
        ),
        crate::Expr::ColRef(_) | crate::Expr::RowRef(_) => {
            let (prefix, rect) = lower_rect_ref(expr, origin, current_sheet, resolve_sheet_id)?;
            let mut build_range = |sheet: SheetId| -> Result<RangeRef, LowerError> {
                let (max_row, max_col) = sheet_max_indices(&sheet, sheet_dimensions)?;
                let start = rect.start_on_sheet(max_row, max_col);
                let end = rect.end_on_sheet(max_row, max_col);
                Ok(RangeRef::new(start, end))
            };

            if prefix.workbook.is_some() {
                let sheet = external_sheet_id(&prefix)?;
                let range = build_range(sheet.clone())?;
                return Ok(BytecodeExpr::MultiRangeRef(MultiRangeRef::new(
                    vec![SheetRangeRef::new(sheet, range)].into(),
                )));
            }
            match prefix.sheet.as_ref() {
                Some(crate::SheetRef::Sheet(name)) => {
                    let Some(sheet_id) = resolve_sheet_id(name) else {
                        return Err(LowerError::UnknownSheet);
                    };
                    let sheet = SheetId::Local(sheet_id);
                    let range = build_range(sheet.clone())?;
                    if sheet_id == current_sheet {
                        Ok(BytecodeExpr::RangeRef(range))
                    } else {
                        let areas = vec![SheetRangeRef::new(sheet, range)];
                        Ok(BytecodeExpr::MultiRangeRef(MultiRangeRef::new(
                            areas.into(),
                        )))
                    }
                }
                Some(crate::SheetRef::SheetRange { start, end }) => {
                    let sheets =
                        expand_sheet_span(start, end, resolve_sheet_id, expand_sheet_span_ids)?;
                    let mut areas: Vec<SheetRangeRef> = Vec::with_capacity(sheets.len());
                    for sheet_id in sheets {
                        let sheet = SheetId::Local(sheet_id);
                        let range = build_range(sheet.clone())?;
                        areas.push(SheetRangeRef::new(sheet, range));
                    }
                    Ok(BytecodeExpr::MultiRangeRef(MultiRangeRef::new(
                        areas.into(),
                    )))
                }
                _ => {
                    let range = build_range(SheetId::Local(current_sheet))?;
                    Ok(BytecodeExpr::RangeRef(range))
                }
            }
        }
        crate::Expr::Binary(b) => match b.op {
            crate::BinaryOp::Range => lower_range_ref(
                &b.left,
                &b.right,
                origin,
                current_sheet,
                resolve_sheet_id,
                expand_sheet_span_ids,
                sheet_dimensions,
            ),
            crate::BinaryOp::Concat => {
                // Flatten `a&b&c` into a single CONCAT_OP call so we avoid intermediate allocations
                // during evaluation and maximize cache sharing between equivalent concat chains.
                let mut operands = Vec::new();
                collect_concat_operands(&b.left, &mut operands);
                collect_concat_operands(&b.right, &mut operands);
                let mut args = Vec::with_capacity(operands.len());
                for expr in operands {
                    args.push(lower_canonical_expr_inner(
                        expr,
                        origin,
                        current_sheet,
                        resolve_sheet_id,
                        expand_sheet_span_ids,
                        sheet_dimensions,
                        scopes,
                        lambda_self_name,
                    )?);
                }
                Ok(BytecodeExpr::FuncCall {
                    func: Function::ConcatOp,
                    args,
                })
            }
            crate::BinaryOp::Add
            | crate::BinaryOp::Sub
            | crate::BinaryOp::Mul
            | crate::BinaryOp::Div
            | crate::BinaryOp::Pow
            | crate::BinaryOp::Eq
            | crate::BinaryOp::Ne
            | crate::BinaryOp::Lt
            | crate::BinaryOp::Le
            | crate::BinaryOp::Gt
            | crate::BinaryOp::Ge => {
                let op = match b.op {
                    crate::BinaryOp::Add => BinaryOp::Add,
                    crate::BinaryOp::Sub => BinaryOp::Sub,
                    crate::BinaryOp::Mul => BinaryOp::Mul,
                    crate::BinaryOp::Div => BinaryOp::Div,
                    crate::BinaryOp::Pow => BinaryOp::Pow,
                    crate::BinaryOp::Eq => BinaryOp::Eq,
                    crate::BinaryOp::Ne => BinaryOp::Ne,
                    crate::BinaryOp::Lt => BinaryOp::Lt,
                    crate::BinaryOp::Le => BinaryOp::Le,
                    crate::BinaryOp::Gt => BinaryOp::Gt,
                    crate::BinaryOp::Ge => BinaryOp::Ge,
                    _ => unreachable!("guarded above"),
                };
                Ok(BytecodeExpr::Binary {
                    op,
                    left: Box::new(lower_canonical_expr_inner(
                        &b.left,
                        origin,
                        current_sheet,
                        resolve_sheet_id,
                        expand_sheet_span_ids,
                        sheet_dimensions,
                        scopes,
                        lambda_self_name,
                    )?),
                    right: Box::new(lower_canonical_expr_inner(
                        &b.right,
                        origin,
                        current_sheet,
                        resolve_sheet_id,
                        expand_sheet_span_ids,
                        sheet_dimensions,
                        scopes,
                        lambda_self_name,
                    )?),
                })
            }
            crate::BinaryOp::Union | crate::BinaryOp::Intersect => {
                // Reference algebra operators evaluate operands in "reference context" (e.g. `A1`
                // behaves like a single-cell range).
                lower_canonical_reference_expr(
                    expr,
                    origin,
                    current_sheet,
                    resolve_sheet_id,
                    expand_sheet_span_ids,
                    sheet_dimensions,
                    scopes,
                    lambda_self_name,
                )
            }
        },
        crate::Expr::Unary(u) => match u.op {
            crate::UnaryOp::Plus => Ok(BytecodeExpr::Unary {
                op: UnaryOp::Plus,
                expr: Box::new(lower_canonical_expr_inner(
                    &u.expr,
                    origin,
                    current_sheet,
                    resolve_sheet_id,
                    expand_sheet_span_ids,
                    sheet_dimensions,
                    scopes,
                    lambda_self_name,
                )?),
            }),
            crate::UnaryOp::Minus => Ok(BytecodeExpr::Unary {
                op: UnaryOp::Neg,
                expr: Box::new(lower_canonical_expr_inner(
                    &u.expr,
                    origin,
                    current_sheet,
                    resolve_sheet_id,
                    expand_sheet_span_ids,
                    sheet_dimensions,
                    scopes,
                    lambda_self_name,
                )?),
            }),
            crate::UnaryOp::ImplicitIntersection => Ok(BytecodeExpr::Unary {
                op: UnaryOp::ImplicitIntersection,
                expr: Box::new(lower_canonical_expr_inner(
                    &u.expr,
                    origin,
                    current_sheet,
                    resolve_sheet_id,
                    expand_sheet_span_ids,
                    sheet_dimensions,
                    scopes,
                    lambda_self_name,
                )?),
            }),
        },
        crate::Expr::Call(call) => {
            let callee = lower_canonical_expr_inner(
                &call.callee,
                origin,
                current_sheet,
                resolve_sheet_id,
                expand_sheet_span_ids,
                sheet_dimensions,
                scopes,
                lambda_self_name,
            )?;
            let args = call
                .args
                .iter()
                .map(|a| {
                    lower_canonical_expr_inner(
                        a,
                        origin,
                        current_sheet,
                        resolve_sheet_id,
                        expand_sheet_span_ids,
                        sheet_dimensions,
                        scopes,
                        lambda_self_name,
                    )
                })
                .collect::<Result<Vec<_>, _>>()?;
            Ok(BytecodeExpr::Call {
                callee: Box::new(callee),
                args,
            })
        }
        crate::Expr::FunctionCall(call) => match call.name.name_upper.as_str() {
            "LET" => lower_let(
                call,
                origin,
                current_sheet,
                resolve_sheet_id,
                expand_sheet_span_ids,
                sheet_dimensions,
                scopes,
            ),
            "LAMBDA" => lower_lambda(
                call,
                origin,
                current_sheet,
                resolve_sheet_id,
                expand_sheet_span_ids,
                sheet_dimensions,
                scopes,
                lambda_self_name,
            ),
            "ISOMITTED" => lower_isomitted(call),
            name_upper => {
                let func = Function::from_name(name_upper);
                match func {
                    Function::Unknown(name) => {
                        let args = call
                            .args
                            .iter()
                            .map(|a| {
                                lower_canonical_expr_inner(
                                    a,
                                    origin,
                                    current_sheet,
                                    resolve_sheet_id,
                                    expand_sheet_span_ids,
                                    sheet_dimensions,
                                    scopes,
                                    lambda_self_name,
                                )
                            })
                            .collect::<Result<Vec<_>, _>>()?;

                        // If this name is a known builtin function (e.g. RAND), do not treat it as
                        // a lambda invocation. Lower it as an unknown function call so bytecode
                        // eligibility can reject it (yielding `IneligibleExpr` rather than a lower
                        // error).
                        if crate::functions::lookup_function_upper(name_upper).is_some() {
                            return Ok(BytecodeExpr::FuncCall {
                                func: Function::Unknown(name),
                                args,
                            });
                        }

                        // Non-builtin function call. Treat this as a lambda invocation only when the
                        // name is in lexical scope (LET/LAMBDA parameters).
                        let Some(key) = casefolded_key_arc_if(name_upper.trim(), |folded| {
                            scopes.is_defined(folded)
                        }) else {
                            return Err(LowerError::Unsupported);
                        };

                        Ok(BytecodeExpr::Call {
                            callee: Box::new(BytecodeExpr::NameRef(key)),
                            args,
                        })
                    }
                    other => {
                        let args = call
                            .args
                            .iter()
                            .map(|a| {
                                lower_canonical_expr_inner(
                                    a,
                                    origin,
                                    current_sheet,
                                    resolve_sheet_id,
                                    expand_sheet_span_ids,
                                    sheet_dimensions,
                                    scopes,
                                    lambda_self_name,
                                )
                            })
                            .collect::<Result<Vec<_>, _>>()?;
                        Ok(BytecodeExpr::FuncCall { func: other, args })
                    }
                }
            }
        },
        crate::Expr::NameRef(nref) => {
            let prefix = RefPrefix::from_parts(&nref.workbook, &nref.sheet);
            validate_prefix(&prefix, current_sheet, resolve_sheet_id)?;

            // Bytecode currently supports only lexical names (LET/LAMBDA bindings), not workbook
            // defined names. Reject non-local name refs so the engine falls back to AST evaluation.
            if !prefix.is_unprefixed() {
                return Err(LowerError::Unsupported);
            }
            let name = nref.name.trim();
            let Some(key) =
                casefolded_key_arc_if(name, |folded| scopes.is_defined(folded))
            else {
                return Err(LowerError::Unsupported);
            };
            Ok(BytecodeExpr::NameRef(key))
        }
        crate::Expr::Postfix(p) => match p.op {
            crate::PostfixOp::Percent => Ok(BytecodeExpr::Binary {
                op: BinaryOp::Div,
                left: Box::new(lower_canonical_expr_inner(
                    &p.expr,
                    origin,
                    current_sheet,
                    resolve_sheet_id,
                    expand_sheet_span_ids,
                    sheet_dimensions,
                    scopes,
                    lambda_self_name,
                )?),
                right: Box::new(BytecodeExpr::Literal(Value::Number(100.0))),
            }),
            crate::PostfixOp::SpillRange => Ok(BytecodeExpr::SpillRange(Box::new(
                lower_canonical_reference_expr(
                    &p.expr,
                    origin,
                    current_sheet,
                    resolve_sheet_id,
                    expand_sheet_span_ids,
                    sheet_dimensions,
                    scopes,
                    lambda_self_name,
                )?,
            ))),
        },
        crate::Expr::FieldAccess(access) => Ok(BytecodeExpr::FuncCall {
            func: Function::FieldAccess,
            args: vec![
                lower_canonical_expr_inner(
                    &access.base,
                    origin,
                    current_sheet,
                    resolve_sheet_id,
                    expand_sheet_span_ids,
                    sheet_dimensions,
                    scopes,
                    lambda_self_name,
                )?,
                BytecodeExpr::Literal(Value::Text(Arc::from(access.field.as_str()))),
            ],
        }),
        crate::Expr::StructuredRef(_) => Err(LowerError::Unsupported),
    }
}

fn lower_let(
    call: &crate::FunctionCall,
    origin: crate::CellAddr,
    current_sheet: usize,
    resolve_sheet_id: &mut impl FnMut(&str) -> Option<usize>,
    expand_sheet_span_ids: &mut impl FnMut(usize, usize) -> Option<Vec<usize>>,
    sheet_dimensions: &mut impl FnMut(usize) -> Option<(u32, u32)>,
    scopes: &mut LexicalScopes,
) -> Result<BytecodeExpr, LowerError> {
    scopes.push_scope();
    let result = (|| {
        let mut args_out: Vec<BytecodeExpr> = Vec::with_capacity(call.args.len());
        if call.args.len() < 3 || call.args.len() % 2 == 0 {
            // Invalid LET arity: still lower into a LET call so bytecode eligibility can reject it
            // (ensuring we fall back to the AST evaluator for validation + error semantics).
            for (idx, arg) in call.args.iter().enumerate() {
                if idx % 2 == 0 {
                    if let Some(name) = bare_identifier(arg) {
                        let key: Arc<str> = casefolded_key_arc(name.trim());
                        args_out.push(BytecodeExpr::NameRef(Arc::clone(&key)));
                        scopes.define(key);
                        continue;
                    }
                }

                args_out.push(lower_canonical_expr_inner(
                    arg,
                    origin,
                    current_sheet,
                    resolve_sheet_id,
                    expand_sheet_span_ids,
                    sheet_dimensions,
                    scopes,
                    None,
                )?);
            }

            return Ok(BytecodeExpr::FuncCall {
                func: Function::Let,
                args: args_out,
            });
        }

        let last = call.args.len() - 1;
        for pair in call.args[..last].chunks_exact(2) {
            let Some(name) = bare_identifier(&pair[0]) else {
                // LET binding identifiers must be bare unqualified names. For invalid name args,
                // fall back to the AST evaluator so it can surface Excel's exact validation and
                // error semantics.
                return Err(LowerError::Unsupported);
            };
            let key: Arc<str> = casefolded_key_arc(name.trim());
            args_out.push(BytecodeExpr::NameRef(Arc::clone(&key)));

            // Allow the LET binding name to be referenced inside any LAMBDA bodies produced by the
            // value expression (for recursion via `f(x)`).
            let value_expr = lower_canonical_expr_inner(
                &pair[1],
                origin,
                current_sheet,
                resolve_sheet_id,
                expand_sheet_span_ids,
                sheet_dimensions,
                scopes,
                Some(key.as_ref()),
            )?;
            args_out.push(value_expr);
            scopes.define(key);
        }

        let body = lower_canonical_expr_inner(
            &call.args[last],
            origin,
            current_sheet,
            resolve_sheet_id,
            expand_sheet_span_ids,
            sheet_dimensions,
            scopes,
            None,
        )?;
        args_out.push(body);

        Ok(BytecodeExpr::FuncCall {
            func: Function::Let,
            args: args_out,
        })
    })();
    scopes.pop_scope();
    result
}

fn lower_lambda(
    call: &crate::FunctionCall,
    origin: crate::CellAddr,
    current_sheet: usize,
    resolve_sheet_id: &mut impl FnMut(&str) -> Option<usize>,
    expand_sheet_span_ids: &mut impl FnMut(usize, usize) -> Option<Vec<usize>>,
    sheet_dimensions: &mut impl FnMut(usize) -> Option<(u32, u32)>,
    scopes: &mut LexicalScopes,
    lambda_self_name: Option<&str>,
) -> Result<BytecodeExpr, LowerError> {
    if call.args.is_empty() {
        return Ok(value_error_literal());
    }

    let mut params: Vec<Arc<str>> = Vec::with_capacity(call.args.len().saturating_sub(1));
    let mut seen: HashSet<Arc<str>> = HashSet::new();

    for param_expr in &call.args[..call.args.len() - 1] {
        let Some(name) = bare_identifier(param_expr) else {
            return Ok(value_error_literal());
        };
        let key: Arc<str> = casefolded_key_arc(name.trim());
        if !seen.insert(Arc::clone(&key)) {
            return Ok(value_error_literal());
        }
        params.push(key);
    }

    let Some(body_expr) = call.args.last() else {
        return Ok(value_error_literal());
    };

    scopes.push_scope();
    let result = (|| {
        if let Some(self_name) = lambda_self_name {
            scopes.define(Arc::from(self_name));
        }
        for p in &params {
            scopes.define(Arc::clone(p));
        }

        let body = lower_canonical_expr_inner(
            body_expr,
            origin,
            current_sheet,
            resolve_sheet_id,
            expand_sheet_span_ids,
            sheet_dimensions,
            scopes,
            None,
        )?;

        Ok(BytecodeExpr::Lambda {
            params: Arc::from(params.into_boxed_slice()),
            body: Box::new(body),
        })
    })();
    scopes.pop_scope();
    result
}

fn lower_isomitted(call: &crate::FunctionCall) -> Result<BytecodeExpr, LowerError> {
    // ISOMITTED is a special form: it expects a bare identifier and does not evaluate the
    // argument expression. Outside of a lambda invocation, it returns FALSE. For non-identifier
    // arguments (or invalid arity), it returns #VALUE!.
    if call.args.len() != 1 {
        return Ok(value_error_literal());
    }

    let Some(name) = bare_identifier(&call.args[0]) else {
        return Ok(value_error_literal());
    };
    let key: Arc<str> = casefolded_key_arc(name.trim());
    Ok(BytecodeExpr::FuncCall {
        func: Function::IsOmitted,
        args: vec![BytecodeExpr::NameRef(key)],
    })
}
