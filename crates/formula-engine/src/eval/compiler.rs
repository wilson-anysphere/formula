use crate::eval::address::CellAddr;
use crate::eval::ast::{
    BinaryOp, CellRef, CompareOp, CompiledExpr, Expr, NameRef, PostfixOp, RangeRef, SheetReference,
    UnaryOp,
};
use crate::value::ErrorKind;

/// Excel limits (0-indexed).
///
/// These match the bounds enforced by [`crate::eval::parse_a1`].
const MAX_COL: u32 = 16_383;
const MAX_ROW: u32 = 1_048_575;

/// Compile a canonical parser [`crate::Expr`] into the calc-time [`CompiledExpr`] used by
/// [`crate::eval::Evaluator`].
///
/// `resolve_sheet` is responsible for mapping an internal workbook sheet name (e.g. `"Sheet1"`)
/// to an engine sheet id. Returning `None` indicates that the sheet does not exist and should be
/// treated like an invalid reference (evaluates to `#REF!`).
///
/// External workbook references are preserved syntactically but compile to
/// [`SheetReference::External`], which evaluates to `#REF!`.
pub fn compile_canonical_expr(
    expr: &crate::Expr,
    current_sheet: usize,
    current_cell: CellAddr,
    resolve_sheet: &mut impl FnMut(&str) -> Option<usize>,
) -> CompiledExpr {
    compile_expr_inner(expr, current_sheet, current_cell, resolve_sheet)
}

fn compile_expr_inner(
    expr: &crate::Expr,
    current_sheet: usize,
    current_cell: CellAddr,
    resolve_sheet: &mut impl FnMut(&str) -> Option<usize>,
) -> CompiledExpr {
    match expr {
        crate::Expr::Number(raw) => match raw.parse::<f64>() {
            Ok(n) => Expr::Number(n),
            Err(_) => Expr::Error(ErrorKind::Value),
        },
        crate::Expr::String(s) => Expr::Text(s.clone()),
        crate::Expr::Boolean(b) => Expr::Bool(*b),
        crate::Expr::Error(raw) => Expr::Error(parse_error_kind(raw)),
        crate::Expr::Missing => Expr::Blank,
        crate::Expr::NameRef(r) => Expr::NameRef(NameRef {
            sheet: compile_sheet_reference(&r.workbook, &r.sheet, current_sheet, resolve_sheet),
            name: r.name.clone(),
        }),
        crate::Expr::CellRef(r) => {
            let sheet =
                compile_sheet_reference(&r.workbook, &r.sheet, current_sheet, resolve_sheet);
            let Some(col) = coord_to_index(&r.col, current_cell.col, MAX_COL) else {
                return Expr::Error(ErrorKind::Ref);
            };
            let Some(row) = coord_to_index(&r.row, current_cell.row, MAX_ROW) else {
                return Expr::Error(ErrorKind::Ref);
            };
            Expr::CellRef(CellRef {
                sheet,
                addr: CellAddr { row, col },
            })
        }
        crate::Expr::ColRef(r) => {
            let sheet =
                compile_sheet_reference(&r.workbook, &r.sheet, current_sheet, resolve_sheet);
            let Some(col) = coord_to_index(&r.col, current_cell.col, MAX_COL) else {
                return Expr::Error(ErrorKind::Ref);
            };
            Expr::RangeRef(RangeRef {
                sheet,
                start: CellAddr { row: 0, col },
                end: CellAddr { row: MAX_ROW, col },
            })
        }
        crate::Expr::RowRef(r) => {
            let sheet =
                compile_sheet_reference(&r.workbook, &r.sheet, current_sheet, resolve_sheet);
            let Some(row) = coord_to_index(&r.row, current_cell.row, MAX_ROW) else {
                return Expr::Error(ErrorKind::Ref);
            };
            Expr::RangeRef(RangeRef {
                sheet,
                start: CellAddr { row, col: 0 },
                end: CellAddr { row, col: MAX_COL },
            })
        }
        crate::Expr::StructuredRef(r) => {
            // External workbook structured refs are accepted syntactically but not supported.
            if r.workbook.is_some() {
                return Expr::Error(ErrorKind::Ref);
            }

            // The calc engine's structured-ref resolver is sheet-agnostic when the table name is
            // provided, so we ignore any `sheet` prefix for now.
            let mut text = String::new();
            if let Some(table) = &r.table {
                text.push_str(table);
            }
            text.push('[');
            text.push_str(&r.spec);
            text.push(']');
            match crate::structured_refs::parse_structured_ref(&text, 0) {
                Some((sref, end)) if end == text.len() => Expr::StructuredRef(sref),
                _ => Expr::Error(ErrorKind::Name),
            }
        }
        crate::Expr::Array(_) => Expr::Error(ErrorKind::Value),
        crate::Expr::FunctionCall(call) => {
            let name = call.name.name_upper.clone();
            let original_name = call.name.original.clone();
            let args = call
                .args
                .iter()
                .map(|a| compile_expr_inner(a, current_sheet, current_cell, resolve_sheet))
                .collect();
            Expr::FunctionCall {
                name,
                original_name,
                args,
            }
        }
        crate::Expr::Unary(u) => match u.op {
            crate::UnaryOp::Plus => Expr::Unary {
                op: UnaryOp::Plus,
                expr: Box::new(compile_expr_inner(
                    &u.expr,
                    current_sheet,
                    current_cell,
                    resolve_sheet,
                )),
            },
            crate::UnaryOp::Minus => Expr::Unary {
                op: UnaryOp::Minus,
                expr: Box::new(compile_expr_inner(
                    &u.expr,
                    current_sheet,
                    current_cell,
                    resolve_sheet,
                )),
            },
            crate::UnaryOp::ImplicitIntersection => Expr::ImplicitIntersection(Box::new(
                compile_expr_inner(&u.expr, current_sheet, current_cell, resolve_sheet),
            )),
        },
        crate::Expr::Postfix(p) => match p.op {
            crate::PostfixOp::Percent => Expr::Postfix {
                op: PostfixOp::Percent,
                expr: Box::new(compile_expr_inner(
                    &p.expr,
                    current_sheet,
                    current_cell,
                    resolve_sheet,
                )),
            },
            crate::PostfixOp::SpillRange => Expr::SpillRange(Box::new(compile_expr_inner(
                &p.expr,
                current_sheet,
                current_cell,
                resolve_sheet,
            ))),
        },
        crate::Expr::Binary(b) => compile_binary(b, current_sheet, current_cell, resolve_sheet),
    }
}

fn compile_binary(
    b: &crate::BinaryExpr,
    current_sheet: usize,
    current_cell: CellAddr,
    resolve_sheet: &mut impl FnMut(&str) -> Option<usize>,
) -> CompiledExpr {
    match b.op {
        crate::BinaryOp::Eq
        | crate::BinaryOp::Ne
        | crate::BinaryOp::Lt
        | crate::BinaryOp::Le
        | crate::BinaryOp::Gt
        | crate::BinaryOp::Ge => {
            let op = match b.op {
                crate::BinaryOp::Eq => CompareOp::Eq,
                crate::BinaryOp::Ne => CompareOp::Ne,
                crate::BinaryOp::Lt => CompareOp::Lt,
                crate::BinaryOp::Le => CompareOp::Le,
                crate::BinaryOp::Gt => CompareOp::Gt,
                crate::BinaryOp::Ge => CompareOp::Ge,
                _ => unreachable!("handled by match guard"),
            };
            Expr::Compare {
                op,
                left: Box::new(compile_expr_inner(
                    &b.left,
                    current_sheet,
                    current_cell,
                    resolve_sheet,
                )),
                right: Box::new(compile_expr_inner(
                    &b.right,
                    current_sheet,
                    current_cell,
                    resolve_sheet,
                )),
            }
        }
        crate::BinaryOp::Range => {
            if let (Some(left), Some(right)) = (
                try_compile_static_range_operand(&b.left, current_sheet, current_cell, resolve_sheet),
                try_compile_static_range_operand(&b.right, current_sheet, current_cell, resolve_sheet),
            ) {
                if left.sheet == right.sheet {
                    let (start, end) = bounding_rect(left.start, left.end, right.start, right.end);
                    return Expr::RangeRef(RangeRef {
                        sheet: left.sheet,
                        start,
                        end,
                    });
                }
            }

            Expr::Binary {
                op: BinaryOp::Range,
                left: Box::new(compile_expr_inner(
                    &b.left,
                    current_sheet,
                    current_cell,
                    resolve_sheet,
                )),
                right: Box::new(compile_expr_inner(
                    &b.right,
                    current_sheet,
                    current_cell,
                    resolve_sheet,
                )),
            }
        }
        crate::BinaryOp::Intersect => Expr::Binary {
            op: BinaryOp::Intersect,
            left: Box::new(compile_expr_inner(
                &b.left,
                current_sheet,
                current_cell,
                resolve_sheet,
            )),
            right: Box::new(compile_expr_inner(
                &b.right,
                current_sheet,
                current_cell,
                resolve_sheet,
            )),
        },
        crate::BinaryOp::Union => Expr::Binary {
            op: BinaryOp::Union,
            left: Box::new(compile_expr_inner(
                &b.left,
                current_sheet,
                current_cell,
                resolve_sheet,
            )),
            right: Box::new(compile_expr_inner(
                &b.right,
                current_sheet,
                current_cell,
                resolve_sheet,
            )),
        },
        crate::BinaryOp::Pow => Expr::Binary {
            op: BinaryOp::Pow,
            left: Box::new(compile_expr_inner(
                &b.left,
                current_sheet,
                current_cell,
                resolve_sheet,
            )),
            right: Box::new(compile_expr_inner(
                &b.right,
                current_sheet,
                current_cell,
                resolve_sheet,
            )),
        },
        crate::BinaryOp::Mul => Expr::Binary {
            op: BinaryOp::Mul,
            left: Box::new(compile_expr_inner(
                &b.left,
                current_sheet,
                current_cell,
                resolve_sheet,
            )),
            right: Box::new(compile_expr_inner(
                &b.right,
                current_sheet,
                current_cell,
                resolve_sheet,
            )),
        },
        crate::BinaryOp::Div => Expr::Binary {
            op: BinaryOp::Div,
            left: Box::new(compile_expr_inner(
                &b.left,
                current_sheet,
                current_cell,
                resolve_sheet,
            )),
            right: Box::new(compile_expr_inner(
                &b.right,
                current_sheet,
                current_cell,
                resolve_sheet,
            )),
        },
        crate::BinaryOp::Add => Expr::Binary {
            op: BinaryOp::Add,
            left: Box::new(compile_expr_inner(
                &b.left,
                current_sheet,
                current_cell,
                resolve_sheet,
            )),
            right: Box::new(compile_expr_inner(
                &b.right,
                current_sheet,
                current_cell,
                resolve_sheet,
            )),
        },
        crate::BinaryOp::Sub => Expr::Binary {
            op: BinaryOp::Sub,
            left: Box::new(compile_expr_inner(
                &b.left,
                current_sheet,
                current_cell,
                resolve_sheet,
            )),
            right: Box::new(compile_expr_inner(
                &b.right,
                current_sheet,
                current_cell,
                resolve_sheet,
            )),
        },
        crate::BinaryOp::Concat => Expr::Binary {
            op: BinaryOp::Concat,
            left: Box::new(compile_expr_inner(
                &b.left,
                current_sheet,
                current_cell,
                resolve_sheet,
            )),
            right: Box::new(compile_expr_inner(
                &b.right,
                current_sheet,
                current_cell,
                resolve_sheet,
            )),
        },
    }
}

fn coord_to_index(coord: &crate::Coord, origin: u32, max: u32) -> Option<u32> {
    let idx = match coord {
        crate::Coord::A1 { index, .. } => *index,
        crate::Coord::Offset(delta) => origin.checked_add_signed(*delta)?,
    };
    if idx > max {
        return None;
    }
    Some(idx)
}

fn compile_sheet_reference(
    workbook: &Option<String>,
    sheet: &Option<String>,
    current_sheet: usize,
    resolve_sheet: &mut impl FnMut(&str) -> Option<usize>,
) -> SheetReference<usize> {
    match (workbook.as_ref(), sheet.as_ref()) {
        (Some(book), Some(sheet)) => SheetReference::External(format!("[{book}]{sheet}")),
        (Some(book), None) => SheetReference::External(format!("[{book}]")),
        (None, Some(sheet)) => resolve_sheet(sheet)
            .map(SheetReference::Sheet)
            .unwrap_or_else(|| SheetReference::External(sheet.clone())),
        (None, None) => SheetReference::Sheet(current_sheet),
    }
}

fn parse_error_kind(raw: &str) -> ErrorKind {
    match raw.to_ascii_uppercase().as_str() {
        "#NULL!" => ErrorKind::Null,
        "#DIV/0!" => ErrorKind::Div0,
        "#VALUE!" => ErrorKind::Value,
        "#REF!" => ErrorKind::Ref,
        "#NAME?" => ErrorKind::Name,
        "#NUM!" => ErrorKind::Num,
        "#N/A" => ErrorKind::NA,
        "#SPILL!" => ErrorKind::Spill,
        "#CALC!" => ErrorKind::Calc,
        _ => ErrorKind::Value,
    }
}

#[derive(Debug, Clone)]
struct StaticRangeOperand {
    sheet: SheetReference<usize>,
    start: CellAddr,
    end: CellAddr,
}

fn try_compile_static_range_operand(
    expr: &crate::Expr,
    current_sheet: usize,
    current_cell: CellAddr,
    resolve_sheet: &mut impl FnMut(&str) -> Option<usize>,
) -> Option<StaticRangeOperand> {
    match expr {
        crate::Expr::CellRef(r) => {
            let sheet =
                compile_sheet_reference(&r.workbook, &r.sheet, current_sheet, resolve_sheet);
            let col = coord_to_index(&r.col, current_cell.col, MAX_COL)?;
            let row = coord_to_index(&r.row, current_cell.row, MAX_ROW)?;
            let addr = CellAddr { row, col };
            Some(StaticRangeOperand {
                sheet,
                start: addr,
                end: addr,
            })
        }
        crate::Expr::ColRef(r) => {
            let sheet =
                compile_sheet_reference(&r.workbook, &r.sheet, current_sheet, resolve_sheet);
            let col = coord_to_index(&r.col, current_cell.col, MAX_COL)?;
            Some(StaticRangeOperand {
                sheet,
                start: CellAddr { row: 0, col },
                end: CellAddr { row: MAX_ROW, col },
            })
        }
        crate::Expr::RowRef(r) => {
            let sheet =
                compile_sheet_reference(&r.workbook, &r.sheet, current_sheet, resolve_sheet);
            let row = coord_to_index(&r.row, current_cell.row, MAX_ROW)?;
            Some(StaticRangeOperand {
                sheet,
                start: CellAddr { row, col: 0 },
                end: CellAddr { row, col: MAX_COL },
            })
        }
        _ => None,
    }
}

fn bounding_rect(a_start: CellAddr, a_end: CellAddr, b_start: CellAddr, b_end: CellAddr) -> (CellAddr, CellAddr) {
    let min_row = a_start.row.min(a_end.row).min(b_start.row.min(b_end.row));
    let max_row = a_start.row.max(a_end.row).max(b_start.row.max(b_end.row));
    let min_col = a_start.col.min(a_end.col).min(b_start.col.min(b_end.col));
    let max_col = a_start.col.max(a_end.col).max(b_start.col.max(b_end.col));
    (
        CellAddr {
            row: min_row,
            col: min_col,
        },
        CellAddr {
            row: max_row,
            col: max_col,
        },
    )
}
