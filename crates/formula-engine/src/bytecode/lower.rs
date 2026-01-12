use super::ast::{BinaryOp, Expr as BytecodeExpr, Function, UnaryOp};
use super::value::{
    Array, ErrorKind as BytecodeErrorKind, MultiRangeRef, RangeRef, Ref, SheetRangeRef, Value,
};
use std::sync::Arc;

#[derive(thiserror::Error, Debug, Clone, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub enum LowerError {
    #[error("unsupported expression")]
    Unsupported,
    #[error("external workbook references are not supported")]
    ExternalReference,
    #[error("cross-sheet references are not supported")]
    CrossSheetReference,
    #[error("unknown sheet reference")]
    UnknownSheet,
}

const EXCEL_MAX_ROWS_I32: i32 = formula_model::EXCEL_MAX_ROWS as i32;
const EXCEL_MAX_COLS_I32: i32 = formula_model::EXCEL_MAX_COLS as i32;

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

fn validate_prefix(
    prefix: &RefPrefix,
    current_sheet: usize,
    resolve_sheet: &mut impl FnMut(&str) -> Option<usize>,
) -> Result<(), LowerError> {
    if prefix.workbook.is_some() {
        return Err(LowerError::ExternalReference);
    }
    if let Some(sheet) = prefix.sheet.as_ref() {
        match sheet {
            crate::SheetRef::Sheet(name) => {
                let Some(sheet_id) = resolve_sheet(name) else {
                    return Err(LowerError::UnknownSheet);
                };
                if sheet_id != current_sheet {
                    return Err(LowerError::CrossSheetReference);
                }
            }
            crate::SheetRef::SheetRange { start, end } => {
                // Sheet-span references are allowed in the bytecode backend (lowered as a
                // multi-area reference). We still validate that both sheet names resolve.
                if resolve_sheet(start).is_none() || resolve_sheet(end).is_none() {
                    return Err(LowerError::UnknownSheet);
                }
            }
        }
    }
    Ok(())
}

fn expand_sheet_span(
    start: &str,
    end: &str,
    resolve_sheet: &mut impl FnMut(&str) -> Option<usize>,
) -> Result<Vec<usize>, LowerError> {
    let Some(a) = resolve_sheet(start) else {
        return Err(LowerError::UnknownSheet);
    };
    let Some(b) = resolve_sheet(end) else {
        return Err(LowerError::UnknownSheet);
    };
    let (start, end) = if a <= b { (a, b) } else { (b, a) };
    Ok((start..=end).collect())
}

fn lower_coord(coord: &crate::Coord, origin: u32) -> (i32, bool) {
    match coord {
        crate::Coord::A1 { index, abs } => {
            let idx = *index as i32;
            if *abs {
                (idx, true)
            } else {
                (idx - origin as i32, false)
            }
        }
        crate::Coord::Offset(delta) => (*delta, false),
    }
}

fn lower_cell_ref(
    r: &crate::CellRef,
    origin: crate::CellAddr,
    current_sheet: usize,
    resolve_sheet: &mut impl FnMut(&str) -> Option<usize>,
) -> Result<Ref, LowerError> {
    validate_prefix(
        &RefPrefix::from_parts(&r.workbook, &r.sheet),
        current_sheet,
        resolve_sheet,
    )?;
    let (row, row_abs) = lower_coord(&r.row, origin.row);
    let (col, col_abs) = lower_coord(&r.col, origin.col);
    Ok(Ref::new(row, col, row_abs, col_abs))
}

fn lower_cell_ref_expr(
    r: &crate::CellRef,
    origin: crate::CellAddr,
    current_sheet: usize,
    resolve_sheet: &mut impl FnMut(&str) -> Option<usize>,
) -> Result<BytecodeExpr, LowerError> {
    let prefix = RefPrefix::from_parts(&r.workbook, &r.sheet);
    if prefix.workbook.is_some() {
        return Err(LowerError::ExternalReference);
    }

    let cell = lower_cell_ref(r, origin, current_sheet, resolve_sheet)?;

    match prefix.sheet.as_ref() {
        None => Ok(BytecodeExpr::CellRef(cell)),
        Some(crate::SheetRef::Sheet(_)) => {
            // `lower_cell_ref` validated that the sheet matches the current sheet.
            Ok(BytecodeExpr::CellRef(cell))
        }
        Some(crate::SheetRef::SheetRange { start, end }) => {
            let sheets = expand_sheet_span(start, end, resolve_sheet)?;
            let range = RangeRef::new(cell, cell);
            let areas: Vec<SheetRangeRef> = sheets
                .into_iter()
                .map(|sheet| SheetRangeRef::new(sheet, range))
                .collect();
            Ok(BytecodeExpr::MultiRangeRef(MultiRangeRef::new(areas.into())))
        }
    }
}

#[derive(Debug, Clone)]
enum RangeEndpoint {
    Cell(Ref),
    Col { col: (i32, bool) },
    Row { row: (i32, bool) },
}

fn lower_range_endpoint(
    expr: &crate::Expr,
    origin: crate::CellAddr,
    current_sheet: usize,
    resolve_sheet: &mut impl FnMut(&str) -> Option<usize>,
) -> Result<(RefPrefix, RangeEndpoint), LowerError> {
    match expr {
        crate::Expr::CellRef(r) => Ok((
            RefPrefix::from_parts(&r.workbook, &r.sheet),
            RangeEndpoint::Cell(lower_cell_ref(r, origin, current_sheet, resolve_sheet)?),
        )),
        crate::Expr::ColRef(r) => {
            let prefix = RefPrefix::from_parts(&r.workbook, &r.sheet);
            validate_prefix(&prefix, current_sheet, resolve_sheet)?;
            Ok((
                prefix,
                RangeEndpoint::Col {
                    col: lower_coord(&r.col, origin.col),
                },
            ))
        }
        crate::Expr::RowRef(r) => {
            let prefix = RefPrefix::from_parts(&r.workbook, &r.sheet);
            validate_prefix(&prefix, current_sheet, resolve_sheet)?;
            Ok((
                prefix,
                RangeEndpoint::Row {
                    row: lower_coord(&r.row, origin.row),
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
    resolve_sheet: &mut impl FnMut(&str) -> Option<usize>,
) -> Result<BytecodeExpr, LowerError> {
    let (left_prefix, left_ep) = lower_range_endpoint(left, origin, current_sheet, resolve_sheet)?;
    let (right_prefix, right_ep) =
        lower_range_endpoint(right, origin, current_sheet, resolve_sheet)?;
    let merged_prefix = merge_range_prefix(&left_prefix, &right_prefix)?;
    validate_prefix(&merged_prefix, current_sheet, resolve_sheet)?;

    let range = match (left_ep, right_ep) {
        (RangeEndpoint::Cell(a), RangeEndpoint::Cell(b)) => RangeRef::new(a, b),
        (RangeEndpoint::Col { col: a }, RangeEndpoint::Col { col: b }) => {
            let (col_a, col_abs_a) = a;
            let (col_b, col_abs_b) = b;
            let start = Ref::new(0, col_a, true, col_abs_a);
            let end = Ref::new(EXCEL_MAX_ROWS_I32 - 1, col_b, true, col_abs_b);
            RangeRef::new(start, end)
        }
        (RangeEndpoint::Row { row: a }, RangeEndpoint::Row { row: b }) => {
            let (row_a, row_abs_a) = a;
            let (row_b, row_abs_b) = b;
            let start = Ref::new(row_a, 0, row_abs_a, true);
            let end = Ref::new(row_b, EXCEL_MAX_COLS_I32 - 1, row_abs_b, true);
            RangeRef::new(start, end)
        }
        _ => return Err(LowerError::Unsupported),
    };

    match merged_prefix.sheet.as_ref() {
        Some(crate::SheetRef::SheetRange { start, end }) => {
            let sheets = expand_sheet_span(start, end, resolve_sheet)?;
            let areas: Vec<SheetRangeRef> = sheets
                .into_iter()
                .map(|sheet| SheetRangeRef::new(sheet, range))
                .collect();
            Ok(BytecodeExpr::MultiRangeRef(MultiRangeRef::new(areas.into())))
        }
        _ => Ok(BytecodeExpr::RangeRef(range)),
    }
}

fn parse_number(raw: &str) -> Result<f64, LowerError> {
    match raw.parse::<f64>() {
        Ok(n) if n.is_finite() => Ok(n),
        _ => Err(LowerError::Unsupported),
    }
}

fn parse_error_kind(raw: &str) -> crate::value::ErrorKind {
    crate::value::ErrorKind::from_code(raw).unwrap_or(crate::value::ErrorKind::Value)
}

fn lower_array_literal_element(expr: &crate::Expr) -> Result<Option<f64>, LowerError> {
    match expr {
        crate::Expr::Number(raw) => Ok(Some(parse_number(raw)?)),
        crate::Expr::Missing | crate::Expr::Boolean(_) | crate::Expr::String(_) => Ok(None),
        crate::Expr::Unary(u) => match u.op {
            crate::UnaryOp::Plus => match lower_array_literal_element(&u.expr)? {
                Some(n) => Ok(Some(n)),
                None => Err(LowerError::Unsupported),
            },
            crate::UnaryOp::Minus => match lower_array_literal_element(&u.expr)? {
                Some(n) => Ok(Some(-n)),
                None => Err(LowerError::Unsupported),
            },
            crate::UnaryOp::ImplicitIntersection => Err(LowerError::Unsupported),
        },
        // Array literals can contain error constants, but the bytecode backend's numeric-only
        // arrays cannot represent them yet (they must be preserved for correct propagation).
        crate::Expr::Error(_) => Err(LowerError::Unsupported),
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
            let n = lower_array_literal_element(el)?;
            values.push(n.unwrap_or(f64::NAN));
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

pub fn lower_canonical_expr(
    expr: &crate::Expr,
    origin: crate::CellAddr,
    current_sheet: usize,
    resolve_sheet: &mut impl FnMut(&str) -> Option<usize>,
) -> Result<BytecodeExpr, LowerError> {
    match expr {
        crate::Expr::Number(raw) => Ok(BytecodeExpr::Literal(Value::Number(parse_number(raw)?))),
        crate::Expr::String(s) => Ok(BytecodeExpr::Literal(Value::Text(Arc::from(s.as_str())))),
        crate::Expr::Boolean(b) => Ok(BytecodeExpr::Literal(Value::Bool(*b))),
        crate::Expr::Error(raw) => Ok(BytecodeExpr::Literal(Value::Error(
            BytecodeErrorKind::from(parse_error_kind(raw)),
        ))),
        crate::Expr::Missing => Ok(BytecodeExpr::Literal(Value::Empty)),
        crate::Expr::Array(arr) => Ok(BytecodeExpr::Literal(lower_array_literal(arr)?)),
        crate::Expr::CellRef(r) => lower_cell_ref_expr(r, origin, current_sheet, resolve_sheet),
        crate::Expr::Binary(b) => match b.op {
            crate::BinaryOp::Range => {
                lower_range_ref(&b.left, &b.right, origin, current_sheet, resolve_sheet)
            }
            crate::BinaryOp::Concat => {
                // Flatten `a&b&c` into a single CONCAT call so we avoid intermediate allocations
                // during evaluation and maximize cache sharing between equivalent concat chains.
                let mut operands = Vec::new();
                collect_concat_operands(&b.left, &mut operands);
                collect_concat_operands(&b.right, &mut operands);
                let args = operands
                    .into_iter()
                    .map(|expr| lower_canonical_expr(expr, origin, current_sheet, resolve_sheet))
                    .collect::<Result<Vec<_>, _>>()?;
                Ok(BytecodeExpr::FuncCall {
                    func: Function::Concat,
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
                    left: Box::new(lower_canonical_expr(
                        &b.left,
                        origin,
                        current_sheet,
                        resolve_sheet,
                    )?),
                    right: Box::new(lower_canonical_expr(
                        &b.right,
                        origin,
                        current_sheet,
                        resolve_sheet,
                    )?),
                })
            }
            crate::BinaryOp::Union | crate::BinaryOp::Intersect => Err(LowerError::Unsupported),
        },
        crate::Expr::Unary(u) => match u.op {
            crate::UnaryOp::Plus => Ok(BytecodeExpr::Unary {
                op: UnaryOp::Plus,
                expr: Box::new(lower_canonical_expr(
                    &u.expr,
                    origin,
                    current_sheet,
                    resolve_sheet,
                )?),
            }),
            crate::UnaryOp::Minus => Ok(BytecodeExpr::Unary {
                op: UnaryOp::Neg,
                expr: Box::new(lower_canonical_expr(
                    &u.expr,
                    origin,
                    current_sheet,
                    resolve_sheet,
                )?),
            }),
            crate::UnaryOp::ImplicitIntersection => Ok(BytecodeExpr::Unary {
                op: UnaryOp::ImplicitIntersection,
                expr: Box::new(lower_canonical_expr(
                    &u.expr,
                    origin,
                    current_sheet,
                    resolve_sheet,
                )?),
            }),
        },
        crate::Expr::FunctionCall(call) => {
            let func = Function::from_name(&call.name.name_upper);
            let args = call
                .args
                .iter()
                .map(|a| lower_canonical_expr(a, origin, current_sheet, resolve_sheet))
                .collect::<Result<Vec<_>, _>>()?;
            Ok(BytecodeExpr::FuncCall { func, args })
        }
        crate::Expr::Call(_) => Err(LowerError::Unsupported),
        crate::Expr::Postfix(p) => match p.op {
            crate::PostfixOp::Percent => Ok(BytecodeExpr::Binary {
                op: BinaryOp::Div,
                left: Box::new(lower_canonical_expr(
                    &p.expr,
                    origin,
                    current_sheet,
                    resolve_sheet,
                )?),
                right: Box::new(BytecodeExpr::Literal(Value::Number(100.0))),
            }),
            crate::PostfixOp::SpillRange => Err(LowerError::Unsupported),
        },
        crate::Expr::NameRef(nref) => {
            // Bytecode locals (LET) only support unqualified identifiers.
            // Defined names / sheet-qualified names are currently handled by the AST evaluator.
            if nref.workbook.is_some() {
                return Err(LowerError::ExternalReference);
            }
            if nref.sheet.is_some() {
                return Err(LowerError::Unsupported);
            }
            let key = crate::value::casefold(nref.name.trim());
            Ok(BytecodeExpr::NameRef(Arc::from(key)))
        }
        crate::Expr::ColRef(_)
        | crate::Expr::RowRef(_)
        | crate::Expr::StructuredRef(_) => Err(LowerError::Unsupported),
    }
}
