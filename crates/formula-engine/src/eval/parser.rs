use std::str::FromStr;

use crate::eval::address::{AddressParseError, CellAddr};
use crate::eval::ast::{
    BinaryOp, CellRef, CompareOp, Expr, NameRef, ParsedExpr, PostfixOp, RangeRef, Ref,
    SheetReference, StructuredRefExpr, UnaryOp,
};
use crate::eval::sheet_reference::lower_sheet_reference;
use crate::value::ErrorKind;
use thiserror::Error;

#[derive(Debug, Error, Clone, PartialEq, Eq)]
pub enum FormulaParseError {
    #[error("unexpected end of input")]
    UnexpectedEof,
    #[error("unexpected token: {0}")]
    UnexpectedToken(String),
    #[error("invalid address: {0}")]
    InvalidAddress(#[from] AddressParseError),
    #[error("expected {expected}, got {got}")]
    Expected { expected: String, got: String },
}

pub struct Parser;

impl Parser {
    pub fn parse(formula: &str) -> Result<ParsedExpr, FormulaParseError> {
        let ast = crate::parse_formula(formula, crate::ParseOptions::default())
            .map_err(|e| FormulaParseError::UnexpectedToken(e.message))?;
        Ok(lower_expr(&ast.expr))
    }
}

fn lower_expr(expr: &crate::Expr) -> ParsedExpr {
    match expr {
        crate::Expr::Number(raw) => match f64::from_str(raw) {
            Ok(n) => Expr::Number(n),
            Err(_) => Expr::Error(ErrorKind::Value),
        },
        crate::Expr::String(s) => Expr::Text(s.clone()),
        crate::Expr::Boolean(b) => Expr::Bool(*b),
        crate::Expr::Error(code) => Expr::Error(parse_error_kind(code)),
        crate::Expr::Missing => Expr::Blank,

        crate::Expr::NameRef(r) => Expr::NameRef(NameRef {
            sheet: lower_sheet_reference(&r.workbook, &r.sheet),
            name: r.name.clone(),
        }),
        crate::Expr::FieldAccess(access) => Expr::FieldAccess {
            base: Box::new(lower_expr(access.base.as_ref())),
            field: access.field.clone(),
        },

        crate::Expr::CellRef(r) => lower_cell_ref(r)
            .map(Expr::CellRef)
            .unwrap_or_else(|| Expr::Error(ErrorKind::Ref)),

        // Standalone row/col refs are uncommon in A1 formulas (Excel normally uses `A:A` / `1:1`),
        // but lower them to their full-row/full-column ranges for completeness.
        crate::Expr::ColRef(r) => rect_from_col_ref(r)
            .and_then(|rect| {
                Some(Expr::RangeRef(RangeRef {
                    sheet: rect.sheet,
                    start: Ref::from_abs_cell_addr(rect.start)?,
                    end: Ref::from_abs_cell_addr(rect.end)?,
                }))
            })
            .unwrap_or_else(|| Expr::Error(ErrorKind::Ref)),
        crate::Expr::RowRef(r) => rect_from_row_ref(r)
            .and_then(|rect| {
                Some(Expr::RangeRef(RangeRef {
                    sheet: rect.sheet,
                    start: Ref::from_abs_cell_addr(rect.start)?,
                    end: Ref::from_abs_cell_addr(rect.end)?,
                }))
            })
            .unwrap_or_else(|| Expr::Error(ErrorKind::Ref)),

        crate::Expr::StructuredRef(r) => lower_structured_ref(r),

        crate::Expr::Array(arr) => lower_array_literal(arr),

        crate::Expr::FunctionCall(call) => Expr::FunctionCall {
            name: call.name.name_upper.clone(),
            original_name: call.name.original.clone(),
            args: {
                let mut args: Vec<ParsedExpr> = Vec::new();
                if args.try_reserve_exact(call.args.len()).is_err() {
                    debug_assert!(
                        false,
                        "allocation failed (lower_expr function args, len={})",
                        call.args.len()
                    );
                    return Expr::Error(ErrorKind::Num);
                }
                for arg in call.args.iter() {
                    args.push(lower_expr(arg));
                }
                args
            },
        },
        crate::Expr::Call(call) => Expr::Call {
            callee: Box::new(lower_expr(call.callee.as_ref())),
            args: {
                let mut args: Vec<ParsedExpr> = Vec::new();
                if args.try_reserve_exact(call.args.len()).is_err() {
                    debug_assert!(
                        false,
                        "allocation failed (lower_expr call args, len={})",
                        call.args.len()
                    );
                    return Expr::Error(ErrorKind::Num);
                }
                for arg in call.args.iter() {
                    args.push(lower_expr(arg));
                }
                args
            },
        },

        crate::Expr::Unary(u) => match u.op {
            crate::UnaryOp::Plus => Expr::Unary {
                op: UnaryOp::Plus,
                expr: Box::new(lower_expr(&u.expr)),
            },
            crate::UnaryOp::Minus => Expr::Unary {
                op: UnaryOp::Minus,
                expr: Box::new(lower_expr(&u.expr)),
            },
            crate::UnaryOp::ImplicitIntersection => {
                Expr::ImplicitIntersection(Box::new(lower_expr(&u.expr)))
            }
        },

        crate::Expr::Postfix(p) => match p.op {
            crate::PostfixOp::Percent => Expr::Postfix {
                op: PostfixOp::Percent,
                expr: Box::new(lower_expr(&p.expr)),
            },
            crate::PostfixOp::SpillRange => Expr::SpillRange(Box::new(lower_expr(&p.expr))),
        },

        crate::Expr::Binary(b) => lower_binary(b),
    }
}

fn lower_array_literal(arr: &crate::ArrayLiteral) -> ParsedExpr {
    let rows = arr.rows.len();
    let cols = arr.rows.first().map(|r| r.len()).unwrap_or(0);

    if rows == 0 || cols == 0 {
        return Expr::Error(ErrorKind::Value);
    }
    if arr.rows.iter().any(|r| r.len() != cols) {
        return Expr::Error(ErrorKind::Value);
    }

    let len = match rows.checked_mul(cols) {
        Some(v) => v,
        None => return Expr::Error(ErrorKind::Num),
    };
    let mut values: Vec<ParsedExpr> = Vec::new();
    if values.try_reserve_exact(len).is_err() {
        debug_assert!(false, "allocation failed (lower_array_literal, len={len})");
        return Expr::Error(ErrorKind::Num);
    }
    for row in &arr.rows {
        for el in row {
            values.push(lower_expr(el));
        }
    }

    Expr::ArrayLiteral {
        rows,
        cols,
        values: values.into(),
    }
}

fn lower_binary(expr: &crate::BinaryExpr) -> ParsedExpr {
    use crate::BinaryOp as Op;

    match expr.op {
        Op::Range => lower_range_ref(&expr.left, &expr.right),
        Op::Union => Expr::Binary {
            op: BinaryOp::Union,
            left: Box::new(lower_expr(&expr.left)),
            right: Box::new(lower_expr(&expr.right)),
        },
        Op::Intersect => Expr::Binary {
            op: BinaryOp::Intersect,
            left: Box::new(lower_expr(&expr.left)),
            right: Box::new(lower_expr(&expr.right)),
        },

        Op::Pow => Expr::Binary {
            op: BinaryOp::Pow,
            left: Box::new(lower_expr(&expr.left)),
            right: Box::new(lower_expr(&expr.right)),
        },
        Op::Add => Expr::Binary {
            op: BinaryOp::Add,
            left: Box::new(lower_expr(&expr.left)),
            right: Box::new(lower_expr(&expr.right)),
        },
        Op::Sub => Expr::Binary {
            op: BinaryOp::Sub,
            left: Box::new(lower_expr(&expr.left)),
            right: Box::new(lower_expr(&expr.right)),
        },
        Op::Mul => Expr::Binary {
            op: BinaryOp::Mul,
            left: Box::new(lower_expr(&expr.left)),
            right: Box::new(lower_expr(&expr.right)),
        },
        Op::Div => Expr::Binary {
            op: BinaryOp::Div,
            left: Box::new(lower_expr(&expr.left)),
            right: Box::new(lower_expr(&expr.right)),
        },
        Op::Concat => Expr::Binary {
            op: BinaryOp::Concat,
            left: Box::new(lower_expr(&expr.left)),
            right: Box::new(lower_expr(&expr.right)),
        },

        Op::Eq => Expr::Compare {
            op: CompareOp::Eq,
            left: Box::new(lower_expr(&expr.left)),
            right: Box::new(lower_expr(&expr.right)),
        },
        Op::Ne => Expr::Compare {
            op: CompareOp::Ne,
            left: Box::new(lower_expr(&expr.left)),
            right: Box::new(lower_expr(&expr.right)),
        },
        Op::Lt => Expr::Compare {
            op: CompareOp::Lt,
            left: Box::new(lower_expr(&expr.left)),
            right: Box::new(lower_expr(&expr.right)),
        },
        Op::Le => Expr::Compare {
            op: CompareOp::Le,
            left: Box::new(lower_expr(&expr.left)),
            right: Box::new(lower_expr(&expr.right)),
        },
        Op::Gt => Expr::Compare {
            op: CompareOp::Gt,
            left: Box::new(lower_expr(&expr.left)),
            right: Box::new(lower_expr(&expr.right)),
        },
        Op::Ge => Expr::Compare {
            op: CompareOp::Ge,
            left: Box::new(lower_expr(&expr.left)),
            right: Box::new(lower_expr(&expr.right)),
        },
    }
}

#[derive(Debug, Clone)]
struct RectRef {
    sheet: SheetReference<String>,
    start: CellAddr,
    end: CellAddr,
}

fn lower_range_ref(left: &crate::Expr, right: &crate::Expr) -> ParsedExpr {
    let (Some(l), Some(r)) = (rect_ref(left), rect_ref(right)) else {
        return Expr::Error(ErrorKind::Value);
    };

    let Ok(sheet) = merge_range_sheets(&l.sheet, &r.sheet) else {
        return Expr::Error(ErrorKind::Ref);
    };

    let start = CellAddr {
        row: l.start.row.min(r.start.row),
        col: l.start.col.min(r.start.col),
    };
    let end = CellAddr {
        row: l.end.row.max(r.end.row),
        col: l.end.col.max(r.end.col),
    };

    let Some(start) = Ref::from_abs_cell_addr(start) else {
        return Expr::Error(ErrorKind::Ref);
    };
    let Some(end) = Ref::from_abs_cell_addr(end) else {
        return Expr::Error(ErrorKind::Ref);
    };

    Expr::RangeRef(RangeRef { sheet, start, end })
}

fn rect_ref(expr: &crate::Expr) -> Option<RectRef> {
    match expr {
        crate::Expr::CellRef(r) => {
            let cell = lower_cell_ref(r)?;
            let addr = cell.addr.as_abs_cell_addr()?;
            Some(RectRef {
                sheet: cell.sheet,
                start: addr,
                end: addr,
            })
        }
        crate::Expr::ColRef(r) => rect_from_col_ref(r),
        crate::Expr::RowRef(r) => rect_from_row_ref(r),
        _ => None,
    }
}

fn merge_range_sheets(
    left: &SheetReference<String>,
    right: &SheetReference<String>,
) -> Result<SheetReference<String>, ()> {
    match (left, right) {
        (SheetReference::Current, SheetReference::Current) => Ok(SheetReference::Current),
        (SheetReference::Current, other) => Ok(other.clone()),
        (other, SheetReference::Current) => Ok(other.clone()),
        (a, b) if a == b => Ok(a.clone()),
        _ => Err(()),
    }
}

fn lower_cell_ref(r: &crate::CellRef) -> Option<CellRef<String>> {
    let sheet = lower_sheet_reference(&r.workbook, &r.sheet);
    let row = coord_index(&r.row)?;
    let col = coord_index(&r.col)?;
    let addr = Ref::from_abs_cell_addr(CellAddr { row, col })?;
    Some(CellRef { sheet, addr })
}

fn rect_from_col_ref(r: &crate::ColRef) -> Option<RectRef> {
    let sheet = lower_sheet_reference(&r.workbook, &r.sheet);
    let col = coord_index(&r.col)?;
    Some(RectRef {
        sheet,
        start: CellAddr { row: 0, col },
        // Whole-column references like `A:A` span to the end of the sheet. Use a sentinel that is
        // resolved against the sheet's runtime dimensions during evaluation.
        end: CellAddr {
            row: CellAddr::SHEET_END,
            col,
        },
    })
}

fn rect_from_row_ref(r: &crate::RowRef) -> Option<RectRef> {
    let sheet = lower_sheet_reference(&r.workbook, &r.sheet);
    let row = coord_index(&r.row)?;
    Some(RectRef {
        sheet,
        start: CellAddr { row, col: 0 },
        // Whole-row references like `1:1` span to the end of the sheet. Use a sentinel that is
        // resolved against the sheet's runtime dimensions during evaluation.
        end: CellAddr {
            row,
            col: CellAddr::SHEET_END,
        },
    })
}

fn coord_index(coord: &crate::Coord) -> Option<u32> {
    match coord {
        crate::Coord::A1 { index, .. } => Some(*index),
        crate::Coord::Offset(_) => None,
    }
}

fn lower_structured_ref(r: &crate::StructuredRef) -> ParsedExpr {
    let sheet = lower_sheet_reference(&r.workbook, &r.sheet);
    let Some(sref) = crate::structured_refs::parse_structured_ref_parts(r.table.as_deref(), &r.spec)
    else {
        return Expr::Error(ErrorKind::Name);
    };
    Expr::StructuredRef(StructuredRefExpr { sheet, sref })
}

fn parse_error_kind(code: &str) -> ErrorKind {
    ErrorKind::from_code(code).unwrap_or(ErrorKind::Value)
}
