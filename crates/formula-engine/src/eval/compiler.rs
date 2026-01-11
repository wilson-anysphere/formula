use crate::eval::address::CellAddr;
use crate::eval::ast::{
    BinaryOp, CellRef, CompareOp, CompiledExpr, Expr, NameRef, PostfixOp, RangeRef, SheetReference,
    UnaryOp,
};
use crate::SheetRef;
use crate::value::ErrorKind;

/// Excel limits (0-indexed).
///
/// These match the bounds enforced by [`crate::eval::parse_a1`].
const MAX_COL: u32 = 16_383;
const MAX_ROW: u32 = 1_048_575;

fn parse_number(raw: &str) -> Option<f64> {
    match raw.parse::<f64>() {
        Ok(n) if n.is_finite() => Some(n),
        _ => None,
    }
}

/// Lower a canonical parser [`crate::Ast`] into the evaluation IR used by the engine.
///
/// This is primarily used for defined-name formulas, which need to preserve
/// [`SheetReference::Current`] so they can be evaluated relative to the sheet where the name is
/// used.
pub fn lower_ast(ast: &crate::Ast, origin: Option<crate::CellAddr>) -> Expr<String> {
    lower_expr(&ast.expr, origin)
}

/// Lower a canonical parser [`crate::Expr`] into the evaluation IR used by the engine.
pub fn lower_expr(expr: &crate::Expr, origin: Option<crate::CellAddr>) -> Expr<String> {
    match expr {
        crate::Expr::Number(raw) => match parse_number(raw) {
            Some(n) => Expr::Number(n),
            None => Expr::Error(ErrorKind::Value),
        },
        crate::Expr::String(s) => Expr::Text(s.clone()),
        crate::Expr::Boolean(b) => Expr::Bool(*b),
        crate::Expr::Error(raw) => Expr::Error(parse_error_kind(raw)),
        crate::Expr::Missing => Expr::Blank,
        crate::Expr::NameRef(r) => Expr::NameRef(NameRef {
            sheet: lower_sheet_reference(&r.workbook, &r.sheet),
            name: r.name.clone(),
        }),
        crate::Expr::CellRef(r) => {
            let sheet = lower_sheet_reference(&r.workbook, &r.sheet);
            let Some(col) = coord_to_index_opt(&r.col, origin.map(|o| o.col), MAX_COL) else {
                return Expr::Error(ErrorKind::Ref);
            };
            let Some(row) = coord_to_index_opt(&r.row, origin.map(|o| o.row), MAX_ROW) else {
                return Expr::Error(ErrorKind::Ref);
            };
            Expr::CellRef(CellRef {
                sheet,
                addr: CellAddr { row, col },
            })
        }
        crate::Expr::ColRef(r) => {
            let sheet = lower_sheet_reference(&r.workbook, &r.sheet);
            let Some(col) = coord_to_index_opt(&r.col, origin.map(|o| o.col), MAX_COL) else {
                return Expr::Error(ErrorKind::Ref);
            };
            Expr::RangeRef(RangeRef {
                sheet,
                start: CellAddr { row: 0, col },
                end: CellAddr { row: MAX_ROW, col },
            })
        }
        crate::Expr::RowRef(r) => {
            let sheet = lower_sheet_reference(&r.workbook, &r.sheet);
            let Some(row) = coord_to_index_opt(&r.row, origin.map(|o| o.row), MAX_ROW) else {
                return Expr::Error(ErrorKind::Ref);
            };
            Expr::RangeRef(RangeRef {
                sheet,
                start: CellAddr { row, col: 0 },
                end: CellAddr { row, col: MAX_COL },
            })
        }
        crate::Expr::StructuredRef(r) => lower_structured_ref(r),
        crate::Expr::Array(arr) => lower_array_literal(arr, origin),
        crate::Expr::FunctionCall(call) => Expr::FunctionCall {
            name: call.name.name_upper.clone(),
            original_name: call.name.original.clone(),
            args: call.args.iter().map(|a| lower_expr(a, origin)).collect(),
        },
        crate::Expr::Unary(u) => match u.op {
            crate::UnaryOp::Plus => Expr::Unary {
                op: UnaryOp::Plus,
                expr: Box::new(lower_expr(&u.expr, origin)),
            },
            crate::UnaryOp::Minus => Expr::Unary {
                op: UnaryOp::Minus,
                expr: Box::new(lower_expr(&u.expr, origin)),
            },
            crate::UnaryOp::ImplicitIntersection => {
                Expr::ImplicitIntersection(Box::new(lower_expr(&u.expr, origin)))
            }
        },
        crate::Expr::Postfix(p) => match p.op {
            crate::PostfixOp::Percent => Expr::Postfix {
                op: PostfixOp::Percent,
                expr: Box::new(lower_expr(&p.expr, origin)),
            },
            crate::PostfixOp::SpillRange => Expr::SpillRange(Box::new(lower_expr(&p.expr, origin))),
        },
        crate::Expr::Binary(b) => lower_binary(b, origin),
    }
}

fn lower_binary(b: &crate::BinaryExpr, origin: Option<crate::CellAddr>) -> Expr<String> {
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
                left: Box::new(lower_expr(&b.left, origin)),
                right: Box::new(lower_expr(&b.right, origin)),
            }
        }
        crate::BinaryOp::Range => {
            if let Some(range) = try_lower_static_range_ref(&b.left, &b.right, origin) {
                return Expr::RangeRef(range);
            }

            Expr::Binary {
                op: BinaryOp::Range,
                left: Box::new(lower_expr(&b.left, origin)),
                right: Box::new(lower_expr(&b.right, origin)),
            }
        }
        crate::BinaryOp::Intersect => Expr::Binary {
            op: BinaryOp::Intersect,
            left: Box::new(lower_expr(&b.left, origin)),
            right: Box::new(lower_expr(&b.right, origin)),
        },
        crate::BinaryOp::Union => Expr::Binary {
            op: BinaryOp::Union,
            left: Box::new(lower_expr(&b.left, origin)),
            right: Box::new(lower_expr(&b.right, origin)),
        },
        crate::BinaryOp::Pow => Expr::Binary {
            op: BinaryOp::Pow,
            left: Box::new(lower_expr(&b.left, origin)),
            right: Box::new(lower_expr(&b.right, origin)),
        },
        crate::BinaryOp::Mul => Expr::Binary {
            op: BinaryOp::Mul,
            left: Box::new(lower_expr(&b.left, origin)),
            right: Box::new(lower_expr(&b.right, origin)),
        },
        crate::BinaryOp::Div => Expr::Binary {
            op: BinaryOp::Div,
            left: Box::new(lower_expr(&b.left, origin)),
            right: Box::new(lower_expr(&b.right, origin)),
        },
        crate::BinaryOp::Add => Expr::Binary {
            op: BinaryOp::Add,
            left: Box::new(lower_expr(&b.left, origin)),
            right: Box::new(lower_expr(&b.right, origin)),
        },
        crate::BinaryOp::Sub => Expr::Binary {
            op: BinaryOp::Sub,
            left: Box::new(lower_expr(&b.left, origin)),
            right: Box::new(lower_expr(&b.right, origin)),
        },
        crate::BinaryOp::Concat => Expr::Binary {
            op: BinaryOp::Concat,
            left: Box::new(lower_expr(&b.left, origin)),
            right: Box::new(lower_expr(&b.right, origin)),
        },
    }
}

fn lower_sheet_reference(
    workbook: &Option<String>,
    sheet: &Option<SheetRef>,
) -> SheetReference<String> {
    match (workbook.as_ref(), sheet.as_ref()) {
        (Some(book), Some(sheet_ref)) => match sheet_ref {
            SheetRef::Sheet(sheet) => SheetReference::External(format!("[{book}]{sheet}")),
            SheetRef::SheetRange { start, end } => {
                SheetReference::External(format!("[{book}]{start}:{end}"))
            }
        },
        (Some(book), None) => SheetReference::External(format!("[{book}]")),
        (None, Some(sheet_ref)) => match sheet_ref {
            SheetRef::Sheet(sheet) => SheetReference::Sheet(sheet.clone()),
            SheetRef::SheetRange { start, end } => SheetReference::SheetRange(start.clone(), end.clone()),
        },
        (None, None) => SheetReference::Current,
    }
}

fn coord_to_index_opt(coord: &crate::Coord, origin: Option<u32>, max: u32) -> Option<u32> {
    let idx = match coord {
        crate::Coord::A1 { index, .. } => *index,
        crate::Coord::Offset(delta) => origin?.checked_add_signed(*delta)?,
    };
    if idx > max {
        return None;
    }
    Some(idx)
}

fn lower_structured_ref(r: &crate::StructuredRef) -> Expr<String> {
    // External workbook structured refs are accepted syntactically but not supported.
    if r.workbook.is_some() {
        return Expr::Error(ErrorKind::Ref);
    }

    // The structured-ref resolver is table-name driven when available, so we ignore any `sheet`
    // prefix for now.
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

fn lower_array_literal(arr: &crate::ArrayLiteral, origin: Option<crate::CellAddr>) -> Expr<String> {
    let rows = arr.rows.len();
    let cols = arr.rows.first().map(|r| r.len()).unwrap_or(0);

    if rows == 0 || cols == 0 {
        return Expr::Error(ErrorKind::Value);
    }

    if arr.rows.iter().any(|r| r.len() != cols) {
        return Expr::Error(ErrorKind::Value);
    }

    let mut values = Vec::with_capacity(rows.saturating_mul(cols));
    for row in &arr.rows {
        for el in row {
            values.push(lower_expr(el, origin));
        }
    }

    Expr::ArrayLiteral {
        rows,
        cols,
        values: values.into(),
    }
}

#[derive(Debug, Clone)]
struct StaticRangeOperandUnresolved {
    workbook: Option<String>,
    sheet: Option<SheetRef>,
    start: CellAddr,
    end: CellAddr,
}

impl StaticRangeOperandUnresolved {
    fn is_unprefixed(&self) -> bool {
        self.workbook.is_none() && self.sheet.is_none()
    }
}

fn try_lower_static_range_ref(
    left: &crate::Expr,
    right: &crate::Expr,
    origin: Option<crate::CellAddr>,
) -> Option<RangeRef<String>> {
    let left_op = try_lower_static_range_operand(left, origin)?;
    let right_op = try_lower_static_range_operand(right, origin)?;

    if left_op.workbook == right_op.workbook && left_op.sheet == right_op.sheet {
        let sheet = lower_sheet_reference(&left_op.workbook, &left_op.sheet);
        let (start, end) = bounding_rect(left_op.start, left_op.end, right_op.start, right_op.end);
        return Some(RangeRef { sheet, start, end });
    }

    let explicit = if left_op.is_unprefixed() && !right_op.is_unprefixed() {
        &right_op
    } else if right_op.is_unprefixed() && !left_op.is_unprefixed() {
        &left_op
    } else {
        return None;
    };

    let sheet = lower_sheet_reference(&explicit.workbook, &explicit.sheet);
    let (start, end) = bounding_rect(left_op.start, left_op.end, right_op.start, right_op.end);
    Some(RangeRef { sheet, start, end })
}

fn try_lower_static_range_operand(
    expr: &crate::Expr,
    origin: Option<crate::CellAddr>,
) -> Option<StaticRangeOperandUnresolved> {
    match expr {
        crate::Expr::CellRef(r) => {
            let col = coord_to_index_opt(&r.col, origin.map(|o| o.col), MAX_COL)?;
            let row = coord_to_index_opt(&r.row, origin.map(|o| o.row), MAX_ROW)?;
            let addr = CellAddr { row, col };
            Some(StaticRangeOperandUnresolved {
                workbook: r.workbook.clone(),
                sheet: r.sheet.clone(),
                start: addr,
                end: addr,
            })
        }
        crate::Expr::ColRef(r) => {
            let col = coord_to_index_opt(&r.col, origin.map(|o| o.col), MAX_COL)?;
            Some(StaticRangeOperandUnresolved {
                workbook: r.workbook.clone(),
                sheet: r.sheet.clone(),
                start: CellAddr { row: 0, col },
                end: CellAddr { row: MAX_ROW, col },
            })
        }
        crate::Expr::RowRef(r) => {
            let row = coord_to_index_opt(&r.row, origin.map(|o| o.row), MAX_ROW)?;
            Some(StaticRangeOperandUnresolved {
                workbook: r.workbook.clone(),
                sheet: r.sheet.clone(),
                start: CellAddr { row, col: 0 },
                end: CellAddr { row, col: MAX_COL },
            })
        }
        _ => None,
    }
}

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
        crate::Expr::Number(raw) => match parse_number(raw) {
            Some(n) => Expr::Number(n),
            None => Expr::Error(ErrorKind::Value),
        },
        crate::Expr::String(s) => Expr::Text(s.clone()),
        crate::Expr::Boolean(b) => Expr::Bool(*b),
        crate::Expr::Error(raw) => Expr::Error(parse_error_kind(raw)),
        crate::Expr::Missing => Expr::Blank,
        crate::Expr::NameRef(r) => {
            let sheet = if r.workbook.is_none() && r.sheet.is_none() {
                SheetReference::Current
            } else {
                compile_sheet_reference(&r.workbook, &r.sheet, current_sheet, resolve_sheet)
            };
            Expr::NameRef(NameRef {
                sheet,
                name: r.name.clone(),
            })
        }
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
        crate::Expr::Array(arr) => {
            compile_array_literal(arr, current_sheet, current_cell, resolve_sheet)
        }
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

fn compile_array_literal(
    arr: &crate::ArrayLiteral,
    current_sheet: usize,
    current_cell: CellAddr,
    resolve_sheet: &mut impl FnMut(&str) -> Option<usize>,
) -> CompiledExpr {
    let rows = arr.rows.len();
    let cols = arr.rows.first().map(|r| r.len()).unwrap_or(0);

    if rows == 0 || cols == 0 {
        return Expr::Error(ErrorKind::Value);
    }

    if arr.rows.iter().any(|r| r.len() != cols) {
        return Expr::Error(ErrorKind::Value);
    }

    let mut values = Vec::with_capacity(rows.saturating_mul(cols));
    for row in &arr.rows {
        for el in row {
            values.push(compile_expr_inner(
                el,
                current_sheet,
                current_cell,
                resolve_sheet,
            ));
        }
    }

    Expr::ArrayLiteral {
        rows,
        cols,
        values: values.into(),
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
            if let Some(range) = try_compile_static_range_ref(
                &b.left,
                &b.right,
                current_sheet,
                current_cell,
                resolve_sheet,
            ) {
                return Expr::RangeRef(range);
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
    sheet: &Option<SheetRef>,
    current_sheet: usize,
    resolve_sheet: &mut impl FnMut(&str) -> Option<usize>,
) -> SheetReference<usize> {
    match (workbook.as_ref(), sheet.as_ref()) {
        (Some(book), Some(sheet_ref)) => match sheet_ref {
            SheetRef::Sheet(sheet) => SheetReference::External(format!("[{book}]{sheet}")),
            SheetRef::SheetRange { start, end } => {
                SheetReference::External(format!("[{book}]{start}:{end}"))
            }
        },
        (Some(book), None) => SheetReference::External(format!("[{book}]")),
        (None, Some(sheet_ref)) => match sheet_ref {
            SheetRef::Sheet(sheet) => resolve_sheet(sheet)
                .map(SheetReference::Sheet)
                .unwrap_or_else(|| SheetReference::External(sheet.clone())),
            SheetRef::SheetRange { start, end } => {
                let start_id = resolve_sheet(start);
                let end_id = resolve_sheet(end);
                match (start_id, end_id) {
                    (Some(a), Some(b)) => SheetReference::SheetRange(a, b),
                    _ => SheetReference::External(format!("{start}:{end}")),
                }
            }
        },
        (None, None) => SheetReference::Sheet(current_sheet),
    }
}

fn parse_error_kind(raw: &str) -> ErrorKind {
    match raw.trim().to_ascii_uppercase().as_str() {
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

fn try_compile_static_range_ref(
    left: &crate::Expr,
    right: &crate::Expr,
    current_sheet: usize,
    current_cell: CellAddr,
    resolve_sheet: &mut impl FnMut(&str) -> Option<usize>,
) -> Option<RangeRef<usize>> {
    let left_op =
        try_compile_static_range_operand(left, current_sheet, current_cell, resolve_sheet)?;
    let right_op =
        try_compile_static_range_operand(right, current_sheet, current_cell, resolve_sheet)?;
    if left_op.sheet == right_op.sheet {
        let (start, end) = bounding_rect(left_op.start, left_op.end, right_op.start, right_op.end);
        return Some(RangeRef {
            sheet: left_op.sheet,
            start,
            end,
        });
    }

    // The canonical parser represents `Sheet1!A1:B2` as a range whose left operand has a sheet
    // prefix and whose right operand is unprefixed. Excel treats the prefix as applying to both
    // endpoints, so treat the prefix as applying to both endpoints.
    //
    // For single-sheet prefixes, we recompile using the explicit endpoint's sheet as the "current"
    // sheet so the unprefixed endpoint compiles to the same sheet.
    //
    // For 3D sheet spans (e.g. `Sheet1:Sheet3!A1:B2`), the range naturally spans multiple sheets;
    // in that case we keep the explicit sheet range and reuse the already-resolved cell addresses.
    let explicit_sheet = if is_unprefixed_static_ref(left) {
        explicit_sheet_reference(right, current_sheet, resolve_sheet)
    } else if is_unprefixed_static_ref(right) {
        explicit_sheet_reference(left, current_sheet, resolve_sheet)
    } else {
        None
    }?;

    match explicit_sheet {
        SheetReference::Sheet(merged_sheet) => {
            let left_op = try_compile_static_range_operand(
                left,
                merged_sheet,
                current_cell,
                resolve_sheet,
            )?;
            let right_op = try_compile_static_range_operand(
                right,
                merged_sheet,
                current_cell,
                resolve_sheet,
            )?;
            if left_op.sheet != right_op.sheet {
                return None;
            }

            let (start, end) =
                bounding_rect(left_op.start, left_op.end, right_op.start, right_op.end);
            Some(RangeRef {
                sheet: left_op.sheet,
                start,
                end,
            })
        }
        SheetReference::SheetRange(start_sheet, end_sheet) => {
            let (start, end) =
                bounding_rect(left_op.start, left_op.end, right_op.start, right_op.end);
            Some(RangeRef {
                sheet: SheetReference::SheetRange(start_sheet, end_sheet),
                start,
                end,
            })
        }
        SheetReference::Current => None,
        SheetReference::External(_) => None,
    }
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

fn is_unprefixed_static_ref(expr: &crate::Expr) -> bool {
    match expr {
        crate::Expr::CellRef(r) => r.workbook.is_none() && r.sheet.is_none(),
        crate::Expr::ColRef(r) => r.workbook.is_none() && r.sheet.is_none(),
        crate::Expr::RowRef(r) => r.workbook.is_none() && r.sheet.is_none(),
        _ => false,
    }
}

fn explicit_sheet_reference(
    expr: &crate::Expr,
    current_sheet: usize,
    resolve_sheet: &mut impl FnMut(&str) -> Option<usize>,
) -> Option<SheetReference<usize>> {
    let (workbook, sheet) = match expr {
        crate::Expr::CellRef(r) => (&r.workbook, &r.sheet),
        crate::Expr::ColRef(r) => (&r.workbook, &r.sheet),
        crate::Expr::RowRef(r) => (&r.workbook, &r.sheet),
        _ => return None,
    };
    if workbook.is_none() && sheet.is_none() {
        return None;
    }
    Some(compile_sheet_reference(
        workbook,
        sheet,
        current_sheet,
        resolve_sheet,
    ))
}

fn bounding_rect(
    a_start: CellAddr,
    a_end: CellAddr,
    b_start: CellAddr,
    b_end: CellAddr,
) -> (CellAddr, CellAddr) {
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
