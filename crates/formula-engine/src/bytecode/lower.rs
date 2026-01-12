use super::ast::{BinaryOp, Expr as BytecodeExpr, Function, UnaryOp};
use super::value::{ErrorKind as BytecodeErrorKind, RangeRef, Ref, Value};
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
        let Some(sheet_name) = sheet.as_single_sheet() else {
            return Err(LowerError::CrossSheetReference);
        };
        let Some(sheet_id) = resolve_sheet(sheet_name) else {
            return Err(LowerError::UnknownSheet);
        };
        if sheet_id != current_sheet {
            return Err(LowerError::CrossSheetReference);
        }
    }
    Ok(())
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

    Ok(BytecodeExpr::RangeRef(range))
}

fn parse_number(raw: &str) -> Result<f64, LowerError> {
    match raw.parse::<f64>() {
        Ok(n) if n.is_finite() => Ok(n),
        _ => Err(LowerError::Unsupported),
    }
}

fn parse_error_kind(raw: &str) -> crate::value::ErrorKind {
    // Keep this in sync with `eval::compiler::parse_error_kind` so AST and bytecode evaluation
    // agree on the canonical set of supported error literals.
    match raw.trim().to_ascii_uppercase().as_str() {
        "#NULL!" => crate::value::ErrorKind::Null,
        "#DIV/0!" => crate::value::ErrorKind::Div0,
        "#VALUE!" => crate::value::ErrorKind::Value,
        "#REF!" => crate::value::ErrorKind::Ref,
        "#NAME?" => crate::value::ErrorKind::Name,
        "#NUM!" => crate::value::ErrorKind::Num,
        "#N/A" => crate::value::ErrorKind::NA,
        "#SPILL!" => crate::value::ErrorKind::Spill,
        "#CALC!" => crate::value::ErrorKind::Calc,
        _ => crate::value::ErrorKind::Value,
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
        crate::Expr::CellRef(r) => Ok(BytecodeExpr::CellRef(lower_cell_ref(
            r,
            origin,
            current_sheet,
            resolve_sheet,
        )?)),
        crate::Expr::Binary(b) => match b.op {
            crate::BinaryOp::Range => {
                lower_range_ref(&b.left, &b.right, origin, current_sheet, resolve_sheet)
            }
            crate::BinaryOp::Concat => Ok(BytecodeExpr::FuncCall {
                func: Function::Concat,
                args: vec![
                    lower_canonical_expr(&b.left, origin, current_sheet, resolve_sheet)?,
                    lower_canonical_expr(&b.right, origin, current_sheet, resolve_sheet)?,
                ],
            }),
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
        crate::Expr::NameRef(_)
        | crate::Expr::ColRef(_)
        | crate::Expr::RowRef(_)
        | crate::Expr::StructuredRef(_)
        | crate::Expr::Array(_) => Err(LowerError::Unsupported),
    }
}
