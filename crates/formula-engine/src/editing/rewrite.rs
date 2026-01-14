use std::cmp::{max, min};

use formula_model::sheet_name_eq_case_insensitive;

use crate::{
    parse_formula, ArrayLiteral, Ast, BinaryExpr, BinaryOp, CallExpr, CellAddr,
    CellRef as AstCellRef, ColRef as AstColRef, Coord, Expr, FieldAccessExpr, FunctionCall,
    ParseOptions, PostfixExpr, RowRef as AstRowRef, SerializeOptions, SheetRef, UnaryExpr,
};

const REF_ERROR: &str = "#REF!";

fn sheet_ref_applies_for_sheet_edit<F>(
    sheet: Option<&SheetRef>,
    ctx_sheet: &str,
    edit_sheet: &str,
    resolve_sheet_order_index: &mut F,
) -> bool
where
    F: FnMut(&str) -> Option<usize>,
{
    match sheet {
        None => match (
            resolve_sheet_order_index(ctx_sheet),
            resolve_sheet_order_index(edit_sheet),
        ) {
            (Some(ctx_id), Some(edit_id)) => ctx_id == edit_id,
            _ => sheet_name_eq_case_insensitive(ctx_sheet, edit_sheet),
        },
        Some(SheetRef::Sheet(name)) => match (
            resolve_sheet_order_index(name),
            resolve_sheet_order_index(edit_sheet),
        ) {
            (Some(name_id), Some(edit_id)) => name_id == edit_id,
            _ => sheet_name_eq_case_insensitive(name, edit_sheet),
        },
        Some(SheetRef::SheetRange { start, end }) => {
            let Some(start_id) = resolve_sheet_order_index(start) else {
                return false;
            };
            let Some(end_id) = resolve_sheet_order_index(end) else {
                return false;
            };
            let Some(edit_id) = resolve_sheet_order_index(edit_sheet) else {
                return false;
            };

            let (lo, hi) = if start_id <= end_id {
                (start_id, end_id)
            } else {
                (end_id, start_id)
            };
            edit_id >= lo && edit_id <= hi
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct GridRange {
    pub start_row: u32,
    pub start_col: u32,
    pub end_row: u32,
    pub end_col: u32,
}

impl GridRange {
    pub fn new(start_row: u32, start_col: u32, end_row: u32, end_col: u32) -> Self {
        let sr = min(start_row, end_row);
        let er = max(start_row, end_row);
        let sc = min(start_col, end_col);
        let ec = max(start_col, end_col);
        Self {
            start_row: sr,
            start_col: sc,
            end_row: er,
            end_col: ec,
        }
    }

    pub fn contains(&self, row: u32, col: u32) -> bool {
        row >= self.start_row && row <= self.end_row && col >= self.start_col && col <= self.end_col
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum StructuralEdit {
    InsertRows { sheet: String, row: u32, count: u32 },
    DeleteRows { sheet: String, row: u32, count: u32 },
    InsertCols { sheet: String, col: u32, count: u32 },
    DeleteCols { sheet: String, col: u32, count: u32 },
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct RangeMapEdit {
    pub sheet: String,
    pub moved_region: GridRange,
    pub delta_row: i32,
    pub delta_col: i32,
    pub deleted_region: Option<GridRange>,
}

pub fn rewrite_formula_for_structural_edit(
    formula: &str,
    ctx_sheet: &str,
    cell_origin: CellAddr,
    edit: &StructuralEdit,
) -> (String, bool) {
    rewrite_formula_for_structural_edit_with_resolver(formula, ctx_sheet, cell_origin, edit, |_| {
        None
    })
}

/// Rewrites `formula` so its references remain correct after a row/column insertion/deletion.
///
/// For 3D sheet spans like `Sheet1:Sheet3!A1`, Excel defines the span using *sheet tab order*.
/// Callers must provide a resolver that maps a sheet name to its 0-based tab position.
pub fn rewrite_formula_for_structural_edit_with_sheet_order_resolver(
    formula: &str,
    ctx_sheet: &str,
    cell_origin: CellAddr,
    edit: &StructuralEdit,
    mut resolve_sheet_order_index: impl FnMut(&str) -> Option<usize>,
) -> (String, bool) {
    rewrite_formula_via_ast(formula, cell_origin, |expr| {
        rewrite_expr_for_structural_edit(expr, ctx_sheet, edit, &mut resolve_sheet_order_index)
    })
}

/// Backwards-compatible alias for [`rewrite_formula_for_structural_edit_with_sheet_order_resolver`].
///
/// Note: The resolver is expected to return the sheet's 0-based *tab order index* (not a stable
/// sheet identifier).
pub fn rewrite_formula_for_structural_edit_with_resolver(
    formula: &str,
    ctx_sheet: &str,
    cell_origin: CellAddr,
    edit: &StructuralEdit,
    resolve_sheet_order_index: impl FnMut(&str) -> Option<usize>,
) -> (String, bool) {
    rewrite_formula_for_structural_edit_with_sheet_order_resolver(
        formula,
        ctx_sheet,
        cell_origin,
        edit,
        resolve_sheet_order_index,
    )
}

pub fn rewrite_formula_for_copy_delta(
    formula: &str,
    _ctx_sheet: &str,
    cell_origin: CellAddr,
    delta_row: i32,
    delta_col: i32,
) -> (String, bool) {
    rewrite_formula_via_ast(formula, cell_origin, |expr| {
        rewrite_expr_for_copy_delta(expr, delta_row, delta_col)
    })
}

pub fn rewrite_formula_for_range_map(
    formula: &str,
    ctx_sheet: &str,
    cell_origin: CellAddr,
    edit: &RangeMapEdit,
) -> (String, bool) {
    rewrite_formula_for_range_map_with_resolver(formula, ctx_sheet, cell_origin, edit, |_| None)
}

pub fn rewrite_formula_for_range_map_with_resolver(
    formula: &str,
    ctx_sheet: &str,
    cell_origin: CellAddr,
    edit: &RangeMapEdit,
    mut resolve_sheet_order_index: impl FnMut(&str) -> Option<usize>,
) -> (String, bool) {
    rewrite_formula_via_ast(formula, cell_origin, |expr| {
        rewrite_expr_for_range_map(expr, ctx_sheet, edit, &mut resolve_sheet_order_index)
    })
}

pub fn rewrite_formula_for_sheet_delete(
    formula: &str,
    cell_origin: CellAddr,
    deleted_sheet: &str,
    sheet_order: &[String],
) -> (String, bool) {
    rewrite_formula_via_ast(formula, cell_origin, |expr| {
        rewrite_expr_for_sheet_delete(expr, deleted_sheet, sheet_order)
    })
}

fn rewrite_formula_via_ast<F>(formula: &str, cell_origin: CellAddr, f: F) -> (String, bool)
where
    F: FnOnce(&Expr) -> (Expr, bool),
{
    // Editor/eval paths accept leading whitespace before `=`, so tolerate it here as well.
    // We preserve any leading whitespace on output for minimal diffs.
    let trimmed = formula.trim_start();
    let leading_len = formula.len().saturating_sub(trimmed.len());
    let (leading_ws, canonical_src) = formula.split_at(leading_len);

    let ast = match parse_formula(canonical_src, ParseOptions::default()) {
        Ok(ast) => ast,
        Err(_) => return (formula.to_string(), false),
    };

    let (expr, changed) = f(&ast.expr);
    if !changed {
        return (formula.to_string(), false);
    }

    let new_ast = Ast {
        has_equals: ast.has_equals,
        expr,
    };

    let mut opts = SerializeOptions::default();
    // Preserve `_xlfn.` prefixes for newer Excel functions.
    opts.include_xlfn_prefix = true;
    opts.origin = Some(cell_origin);

    match new_ast.to_string(opts) {
        Ok(out) => (format!("{leading_ws}{out}"), true),
        Err(_) => (formula.to_string(), false),
    }
}

fn rewrite_expr_for_structural_edit(
    expr: &Expr,
    ctx_sheet: &str,
    edit: &StructuralEdit,
    resolve_sheet_order_index: &mut impl FnMut(&str) -> Option<usize>,
) -> (Expr, bool) {
    match expr {
        Expr::FieldAccess(access) => {
            let (base, changed) = rewrite_expr_for_structural_edit(
                access.base.as_ref(),
                ctx_sheet,
                edit,
                resolve_sheet_order_index,
            );
            if !changed {
                return (expr.clone(), false);
            }
            (
                Expr::FieldAccess(FieldAccessExpr {
                    base: Box::new(base),
                    field: access.field.clone(),
                }),
                true,
            )
        }
        Expr::CellRef(r) => {
            rewrite_cell_ref_for_structural_edit(r, ctx_sheet, edit, resolve_sheet_order_index)
        }
        Expr::RowRef(r) => {
            rewrite_row_ref_for_structural_edit(r, ctx_sheet, edit, resolve_sheet_order_index)
        }
        Expr::ColRef(r) => {
            rewrite_col_ref_for_structural_edit(r, ctx_sheet, edit, resolve_sheet_order_index)
        }
        Expr::Binary(b) if b.op == BinaryOp::Range => {
            if let Some(result) = rewrite_range_for_structural_edit(
                expr,
                b,
                ctx_sheet,
                edit,
                resolve_sheet_order_index,
            ) {
                return result;
            }
            rewrite_expr_children(expr, |child| {
                rewrite_expr_for_structural_edit(child, ctx_sheet, edit, resolve_sheet_order_index)
            })
        }
        _ => rewrite_expr_children(expr, |child| {
            rewrite_expr_for_structural_edit(child, ctx_sheet, edit, resolve_sheet_order_index)
        }),
    }
}

fn rewrite_expr_for_copy_delta(expr: &Expr, delta_row: i32, delta_col: i32) -> (Expr, bool) {
    match expr {
        Expr::FieldAccess(access) => {
            let (base, changed) =
                rewrite_expr_for_copy_delta(access.base.as_ref(), delta_row, delta_col);
            if !changed {
                return (expr.clone(), false);
            }
            (
                Expr::FieldAccess(FieldAccessExpr {
                    base: Box::new(base),
                    field: access.field.clone(),
                }),
                true,
            )
        }
        Expr::CellRef(r) => rewrite_cell_ref_for_copy_delta(r, delta_row, delta_col),
        Expr::RowRef(r) => rewrite_row_ref_for_copy_delta(r, delta_row),
        Expr::ColRef(r) => rewrite_col_ref_for_copy_delta(r, delta_col),
        Expr::Binary(b) if b.op == BinaryOp::Range => {
            if let Some(result) = rewrite_range_for_copy_delta(expr, b, delta_row, delta_col) {
                return result;
            }
            rewrite_expr_children(expr, |child| {
                rewrite_expr_for_copy_delta(child, delta_row, delta_col)
            })
        }
        _ => rewrite_expr_children(expr, |child| {
            rewrite_expr_for_copy_delta(child, delta_row, delta_col)
        }),
    }
}

fn rewrite_expr_for_range_map(
    expr: &Expr,
    ctx_sheet: &str,
    edit: &RangeMapEdit,
    resolve_sheet_order_index: &mut impl FnMut(&str) -> Option<usize>,
) -> (Expr, bool) {
    match expr {
        Expr::FieldAccess(access) => {
            let (base, changed) = rewrite_expr_for_range_map(
                access.base.as_ref(),
                ctx_sheet,
                edit,
                resolve_sheet_order_index,
            );
            if !changed {
                return (expr.clone(), false);
            }
            (
                Expr::FieldAccess(FieldAccessExpr {
                    base: Box::new(base),
                    field: access.field.clone(),
                }),
                true,
            )
        }
        Expr::CellRef(r) => {
            rewrite_cell_ref_for_range_map(r, ctx_sheet, edit, resolve_sheet_order_index)
        }
        Expr::Binary(b) if b.op == BinaryOp::Range => {
            if let Some(result) = rewrite_cell_range_for_range_map(
                expr,
                b,
                ctx_sheet,
                edit,
                resolve_sheet_order_index,
            ) {
                return result;
            }
            rewrite_expr_children(expr, |child| {
                rewrite_expr_for_range_map(child, ctx_sheet, edit, resolve_sheet_order_index)
            })
        }
        _ => rewrite_expr_children(expr, |child| {
            rewrite_expr_for_range_map(child, ctx_sheet, edit, resolve_sheet_order_index)
        }),
    }
}

fn rewrite_expr_children<F>(expr: &Expr, mut f: F) -> (Expr, bool)
where
    F: FnMut(&Expr) -> (Expr, bool),
{
    match expr {
        Expr::FieldAccess(FieldAccessExpr { base, field }) => {
            let (base, changed) = f(base);
            if !changed {
                return (expr.clone(), false);
            }
            (
                Expr::FieldAccess(FieldAccessExpr {
                    base: Box::new(base),
                    field: field.clone(),
                }),
                true,
            )
        }
        Expr::FunctionCall(FunctionCall { name, args }) => {
            let mut changed = false;
            let args: Vec<Expr> = args
                .iter()
                .map(|arg| {
                    let (expr, c) = f(arg);
                    changed |= c;
                    expr
                })
                .collect();

            if !changed {
                return (expr.clone(), false);
            }

            (
                Expr::FunctionCall(FunctionCall {
                    name: name.clone(),
                    args,
                }),
                true,
            )
        }
        Expr::Call(CallExpr { callee, args }) => {
            let mut changed = false;

            let (callee, callee_changed) = f(callee);
            changed |= callee_changed;

            let args: Vec<Expr> = args
                .iter()
                .map(|arg| {
                    let (expr, c) = f(arg);
                    changed |= c;
                    expr
                })
                .collect();

            if !changed {
                return (expr.clone(), false);
            }

            (
                Expr::Call(CallExpr {
                    callee: Box::new(callee),
                    args,
                }),
                true,
            )
        }
        Expr::Unary(UnaryExpr { op, expr: inner }) => {
            let (inner, changed) = f(inner);
            if !changed {
                return (expr.clone(), false);
            }
            (
                Expr::Unary(UnaryExpr {
                    op: *op,
                    expr: Box::new(inner),
                }),
                true,
            )
        }
        Expr::Postfix(PostfixExpr { op, expr: inner }) => {
            let (inner, changed) = f(inner);
            if !changed {
                return (expr.clone(), false);
            }
            // The spill-range operator (`#`) cannot be applied to error literals in our canonical
            // grammar (e.g. `#REF!#`), and Excel drops the operator once the base reference becomes
            // invalid. If a rewrite turns the operand into an error, emit just the error.
            if *op == crate::PostfixOp::SpillRange && matches!(inner, Expr::Error(_)) {
                return (inner, true);
            }
            (
                Expr::Postfix(PostfixExpr {
                    op: *op,
                    expr: Box::new(inner),
                }),
                true,
            )
        }
        Expr::Binary(BinaryExpr { op, left, right }) => {
            let (left, left_changed) = f(left);
            let (right, right_changed) = f(right);
            if !left_changed && !right_changed {
                return (expr.clone(), false);
            }
            (
                Expr::Binary(BinaryExpr {
                    op: *op,
                    left: Box::new(left),
                    right: Box::new(right),
                }),
                true,
            )
        }
        Expr::Array(ArrayLiteral { rows }) => {
            let mut changed = false;
            let rows: Vec<Vec<Expr>> = rows
                .iter()
                .map(|row| {
                    row.iter()
                        .map(|el| {
                            let (expr, c) = f(el);
                            changed |= c;
                            expr
                        })
                        .collect()
                })
                .collect();

            if !changed {
                return (expr.clone(), false);
            }

            (Expr::Array(ArrayLiteral { rows }), true)
        }
        _ => (expr.clone(), false),
    }
}

fn sheet_index_in_order(sheet_order: &[String], name: &str) -> Option<usize> {
    sheet_order
        .iter()
        .position(|sheet_name| sheet_name_eq_case_insensitive(sheet_name, name))
}

#[derive(Clone, Debug, PartialEq, Eq)]
enum DeleteSheetRefRewrite {
    Unchanged,
    Adjusted(SheetRef),
    Invalidate,
}

fn rewrite_sheet_ref_for_delete(
    sheet: &SheetRef,
    deleted_sheet: &str,
    sheet_order: &[String],
) -> DeleteSheetRefRewrite {
    match sheet {
        SheetRef::Sheet(name) => {
            if sheet_name_eq_case_insensitive(name, deleted_sheet) {
                DeleteSheetRefRewrite::Invalidate
            } else {
                DeleteSheetRefRewrite::Unchanged
            }
        }
        SheetRef::SheetRange { start, end } => {
            let start_matches = sheet_name_eq_case_insensitive(start, deleted_sheet);
            let end_matches = sheet_name_eq_case_insensitive(end, deleted_sheet);
            if !start_matches && !end_matches {
                return DeleteSheetRefRewrite::Unchanged;
            }

            let Some(start_idx) = sheet_index_in_order(sheet_order, start) else {
                return DeleteSheetRefRewrite::Invalidate;
            };
            let Some(end_idx) = sheet_index_in_order(sheet_order, end) else {
                return DeleteSheetRefRewrite::Invalidate;
            };

            // The span references only the deleted sheet.
            if start_idx == end_idx {
                return DeleteSheetRefRewrite::Invalidate;
            }

            let dir: isize = if end_idx > start_idx { 1 } else { -1 };
            let mut new_start_idx = start_idx as isize;
            let mut new_end_idx = end_idx as isize;

            // When deleting a 3D boundary, Excel shifts it one sheet toward the other boundary.
            if start_matches {
                new_start_idx += dir;
            }
            if end_matches {
                new_end_idx -= dir;
            }

            let Some(new_start) = new_start_idx
                .try_into()
                .ok()
                .and_then(|idx: usize| sheet_order.get(idx))
            else {
                return DeleteSheetRefRewrite::Invalidate;
            };
            let Some(new_end) = new_end_idx
                .try_into()
                .ok()
                .and_then(|idx: usize| sheet_order.get(idx))
            else {
                return DeleteSheetRefRewrite::Invalidate;
            };

            if sheet_name_eq_case_insensitive(new_start, new_end) {
                DeleteSheetRefRewrite::Adjusted(SheetRef::Sheet(new_start.clone()))
            } else {
                DeleteSheetRefRewrite::Adjusted(SheetRef::SheetRange {
                    start: new_start.clone(),
                    end: new_end.clone(),
                })
            }
        }
    }
}

fn rewrite_cell_ref_for_sheet_delete(
    r: &AstCellRef,
    deleted_sheet: &str,
    sheet_order: &[String],
) -> (Expr, bool) {
    if r.workbook.is_some() {
        return (expr_ref(r.clone()), false);
    }
    let Some(sheet_ref) = r.sheet.as_ref() else {
        return (expr_ref(r.clone()), false);
    };

    match rewrite_sheet_ref_for_delete(sheet_ref, deleted_sheet, sheet_order) {
        DeleteSheetRefRewrite::Unchanged => (expr_ref(r.clone()), false),
        DeleteSheetRefRewrite::Adjusted(new_sheet) => {
            let mut out = r.clone();
            out.sheet = Some(new_sheet);
            (expr_ref(out), true)
        }
        DeleteSheetRefRewrite::Invalidate => (Expr::Error(REF_ERROR.to_string()), true),
    }
}

fn rewrite_row_ref_for_sheet_delete(
    r: &AstRowRef,
    deleted_sheet: &str,
    sheet_order: &[String],
) -> (Expr, bool) {
    if r.workbook.is_some() {
        return (Expr::RowRef(r.clone()), false);
    }
    let Some(sheet_ref) = r.sheet.as_ref() else {
        return (Expr::RowRef(r.clone()), false);
    };

    match rewrite_sheet_ref_for_delete(sheet_ref, deleted_sheet, sheet_order) {
        DeleteSheetRefRewrite::Unchanged => (Expr::RowRef(r.clone()), false),
        DeleteSheetRefRewrite::Adjusted(new_sheet) => {
            let mut out = r.clone();
            out.sheet = Some(new_sheet);
            (Expr::RowRef(out), true)
        }
        DeleteSheetRefRewrite::Invalidate => (Expr::Error(REF_ERROR.to_string()), true),
    }
}

fn rewrite_col_ref_for_sheet_delete(
    r: &AstColRef,
    deleted_sheet: &str,
    sheet_order: &[String],
) -> (Expr, bool) {
    if r.workbook.is_some() {
        return (Expr::ColRef(r.clone()), false);
    }
    let Some(sheet_ref) = r.sheet.as_ref() else {
        return (Expr::ColRef(r.clone()), false);
    };

    match rewrite_sheet_ref_for_delete(sheet_ref, deleted_sheet, sheet_order) {
        DeleteSheetRefRewrite::Unchanged => (Expr::ColRef(r.clone()), false),
        DeleteSheetRefRewrite::Adjusted(new_sheet) => {
            let mut out = r.clone();
            out.sheet = Some(new_sheet);
            (Expr::ColRef(out), true)
        }
        DeleteSheetRefRewrite::Invalidate => (Expr::Error(REF_ERROR.to_string()), true),
    }
}

fn rewrite_name_ref_for_sheet_delete(
    r: &crate::NameRef,
    deleted_sheet: &str,
    sheet_order: &[String],
) -> (Expr, bool) {
    if r.workbook.is_some() {
        return (Expr::NameRef(r.clone()), false);
    }
    let Some(sheet_ref) = r.sheet.as_ref() else {
        return (Expr::NameRef(r.clone()), false);
    };

    match rewrite_sheet_ref_for_delete(sheet_ref, deleted_sheet, sheet_order) {
        DeleteSheetRefRewrite::Unchanged => (Expr::NameRef(r.clone()), false),
        DeleteSheetRefRewrite::Adjusted(new_sheet) => {
            let mut out = r.clone();
            out.sheet = Some(new_sheet);
            (Expr::NameRef(out), true)
        }
        DeleteSheetRefRewrite::Invalidate => (Expr::Error(REF_ERROR.to_string()), true),
    }
}

fn rewrite_structured_ref_for_sheet_delete(
    r: &crate::StructuredRef,
    deleted_sheet: &str,
    sheet_order: &[String],
) -> (Expr, bool) {
    if r.workbook.is_some() {
        return (Expr::StructuredRef(r.clone()), false);
    }
    let Some(sheet_ref) = r.sheet.as_ref() else {
        return (Expr::StructuredRef(r.clone()), false);
    };

    match rewrite_sheet_ref_for_delete(sheet_ref, deleted_sheet, sheet_order) {
        DeleteSheetRefRewrite::Unchanged => (Expr::StructuredRef(r.clone()), false),
        DeleteSheetRefRewrite::Adjusted(new_sheet) => {
            let mut out = r.clone();
            out.sheet = Some(new_sheet);
            (Expr::StructuredRef(out), true)
        }
        DeleteSheetRefRewrite::Invalidate => (Expr::Error(REF_ERROR.to_string()), true),
    }
}

fn rewrite_expr_for_sheet_delete(
    expr: &Expr,
    deleted_sheet: &str,
    sheet_order: &[String],
) -> (Expr, bool) {
    match expr {
        Expr::FieldAccess(access) => {
            let (base, changed) =
                rewrite_expr_for_sheet_delete(access.base.as_ref(), deleted_sheet, sheet_order);
            if !changed {
                return (expr.clone(), false);
            }
            (
                Expr::FieldAccess(FieldAccessExpr {
                    base: Box::new(base),
                    field: access.field.clone(),
                }),
                true,
            )
        }
        Expr::CellRef(r) => rewrite_cell_ref_for_sheet_delete(r, deleted_sheet, sheet_order),
        Expr::RowRef(r) => rewrite_row_ref_for_sheet_delete(r, deleted_sheet, sheet_order),
        Expr::ColRef(r) => rewrite_col_ref_for_sheet_delete(r, deleted_sheet, sheet_order),
        Expr::NameRef(r) => rewrite_name_ref_for_sheet_delete(r, deleted_sheet, sheet_order),
        Expr::StructuredRef(r) => {
            rewrite_structured_ref_for_sheet_delete(r, deleted_sheet, sheet_order)
        }
        Expr::Binary(b) if b.op == BinaryOp::Range => {
            let (left, left_changed) =
                rewrite_expr_for_sheet_delete(&b.left, deleted_sheet, sheet_order);
            let (right, right_changed) =
                rewrite_expr_for_sheet_delete(&b.right, deleted_sheet, sheet_order);

            if matches!(left, Expr::Error(_)) || matches!(right, Expr::Error(_)) {
                return (Expr::Error(REF_ERROR.to_string()), true);
            }
            if !left_changed && !right_changed {
                return (expr.clone(), false);
            }
            (
                Expr::Binary(BinaryExpr {
                    op: b.op,
                    left: Box::new(left),
                    right: Box::new(right),
                }),
                true,
            )
        }
        _ => rewrite_expr_children(expr, |child| {
            rewrite_expr_for_sheet_delete(child, deleted_sheet, sheet_order)
        }),
    }
}

fn rewrite_cell_ref_for_structural_edit(
    r: &AstCellRef,
    ctx_sheet: &str,
    edit: &StructuralEdit,
    resolve_sheet_order_index: &mut impl FnMut(&str) -> Option<usize>,
) -> (Expr, bool) {
    if r.workbook.is_some() {
        return (expr_ref(r.clone()), false);
    }
    let edit_sheet = match edit {
        StructuralEdit::InsertRows { sheet, .. }
        | StructuralEdit::DeleteRows { sheet, .. }
        | StructuralEdit::InsertCols { sheet, .. }
        | StructuralEdit::DeleteCols { sheet, .. } => sheet.as_str(),
    };

    if !sheet_ref_applies_for_sheet_edit(
        r.sheet.as_ref(),
        ctx_sheet,
        edit_sheet,
        resolve_sheet_order_index,
    ) {
        return (expr_ref(r.clone()), false);
    }

    let Some((col, col_abs)) = coord_a1(&r.col) else {
        return (expr_ref(r.clone()), false);
    };
    let Some((row, row_abs)) = coord_a1(&r.row) else {
        return (expr_ref(r.clone()), false);
    };

    let mut new_row = row;
    let mut new_col = col;

    match edit {
        StructuralEdit::InsertRows { row: at, count, .. } => {
            new_row = adjust_row_insert(row, *at, *count).unwrap_or(row);
        }
        StructuralEdit::DeleteRows { row: at, count, .. } => {
            let del_end = at.saturating_add(count.saturating_sub(1));
            new_row = match adjust_row_delete(row, *at, del_end, *count) {
                Some(v) => v,
                None => return (Expr::Error(REF_ERROR.to_string()), true),
            };
        }
        StructuralEdit::InsertCols { col: at, count, .. } => {
            new_col = adjust_col_insert(col, *at, *count).unwrap_or(col);
        }
        StructuralEdit::DeleteCols { col: at, count, .. } => {
            let del_end = at.saturating_add(count.saturating_sub(1));
            new_col = match adjust_col_delete(col, *at, del_end, *count) {
                Some(v) => v,
                None => return (Expr::Error(REF_ERROR.to_string()), true),
            };
        }
    }

    let new_ref = AstCellRef {
        workbook: r.workbook.clone(),
        sheet: r.sheet.clone(),
        col: Coord::A1 {
            index: new_col,
            abs: col_abs,
        },
        row: Coord::A1 {
            index: new_row,
            abs: row_abs,
        },
    };

    let changed = &new_ref != r;
    (expr_ref(new_ref), changed)
}

fn rewrite_row_ref_for_structural_edit(
    r: &AstRowRef,
    ctx_sheet: &str,
    edit: &StructuralEdit,
    resolve_sheet_order_index: &mut impl FnMut(&str) -> Option<usize>,
) -> (Expr, bool) {
    if r.workbook.is_some() {
        return (Expr::RowRef(r.clone()), false);
    }
    let edit_sheet = match edit {
        StructuralEdit::InsertRows { sheet, .. } | StructuralEdit::DeleteRows { sheet, .. } => {
            sheet.as_str()
        }
        _ => return (Expr::RowRef(r.clone()), false),
    };

    if !sheet_ref_applies_for_sheet_edit(
        r.sheet.as_ref(),
        ctx_sheet,
        edit_sheet,
        resolve_sheet_order_index,
    ) {
        return (Expr::RowRef(r.clone()), false);
    }

    let Some((row, abs)) = coord_a1(&r.row) else {
        return (Expr::RowRef(r.clone()), false);
    };

    let new_row = match edit {
        StructuralEdit::InsertRows { row: at, count, .. } => {
            adjust_row_insert(row, *at, *count).unwrap_or(row)
        }
        StructuralEdit::DeleteRows { row: at, count, .. } => {
            let del_end = at.saturating_add(count.saturating_sub(1));
            match adjust_row_delete(row, *at, del_end, *count) {
                Some(v) => v,
                None => return (Expr::Error(REF_ERROR.to_string()), true),
            }
        }
        _ => row,
    };

    let new_ref = AstRowRef {
        workbook: r.workbook.clone(),
        sheet: r.sheet.clone(),
        row: Coord::A1 {
            index: new_row,
            abs,
        },
    };

    let changed = &new_ref != r;
    (Expr::RowRef(new_ref), changed)
}

fn rewrite_col_ref_for_structural_edit(
    r: &AstColRef,
    ctx_sheet: &str,
    edit: &StructuralEdit,
    resolve_sheet_order_index: &mut impl FnMut(&str) -> Option<usize>,
) -> (Expr, bool) {
    if r.workbook.is_some() {
        return (Expr::ColRef(r.clone()), false);
    }
    let edit_sheet = match edit {
        StructuralEdit::InsertCols { sheet, .. } | StructuralEdit::DeleteCols { sheet, .. } => {
            sheet.as_str()
        }
        _ => return (Expr::ColRef(r.clone()), false),
    };

    if !sheet_ref_applies_for_sheet_edit(
        r.sheet.as_ref(),
        ctx_sheet,
        edit_sheet,
        resolve_sheet_order_index,
    ) {
        return (Expr::ColRef(r.clone()), false);
    }

    let Some((col, abs)) = coord_a1(&r.col) else {
        return (Expr::ColRef(r.clone()), false);
    };

    let new_col = match edit {
        StructuralEdit::InsertCols { col: at, count, .. } => {
            adjust_col_insert(col, *at, *count).unwrap_or(col)
        }
        StructuralEdit::DeleteCols { col: at, count, .. } => {
            let del_end = at.saturating_add(count.saturating_sub(1));
            match adjust_col_delete(col, *at, del_end, *count) {
                Some(v) => v,
                None => return (Expr::Error(REF_ERROR.to_string()), true),
            }
        }
        _ => col,
    };

    let new_ref = AstColRef {
        workbook: r.workbook.clone(),
        sheet: r.sheet.clone(),
        col: Coord::A1 {
            index: new_col,
            abs,
        },
    };

    let changed = &new_ref != r;
    (Expr::ColRef(new_ref), changed)
}

fn rewrite_range_for_structural_edit(
    original: &Expr,
    b: &BinaryExpr,
    ctx_sheet: &str,
    edit: &StructuralEdit,
    resolve_sheet_order_index: &mut impl FnMut(&str) -> Option<usize>,
) -> Option<(Expr, bool)> {
    match (&*b.left, &*b.right) {
        (Expr::CellRef(start), Expr::CellRef(end)) => Some(rewrite_cell_range_for_structural_edit(
            original,
            start,
            end,
            ctx_sheet,
            edit,
            resolve_sheet_order_index,
        )),
        (Expr::RowRef(start), Expr::RowRef(end)) => Some(rewrite_row_range_for_structural_edit(
            original,
            start,
            end,
            ctx_sheet,
            edit,
            resolve_sheet_order_index,
        )),
        (Expr::ColRef(start), Expr::ColRef(end)) => Some(rewrite_col_range_for_structural_edit(
            original,
            start,
            end,
            ctx_sheet,
            edit,
            resolve_sheet_order_index,
        )),
        _ => None,
    }
}

fn rewrite_cell_range_for_structural_edit(
    original: &Expr,
    start: &AstCellRef,
    end: &AstCellRef,
    ctx_sheet: &str,
    edit: &StructuralEdit,
    resolve_sheet_order_index: &mut impl FnMut(&str) -> Option<usize>,
) -> (Expr, bool) {
    if start.workbook.is_some() || end.workbook.is_some() {
        return (original.clone(), false);
    }
    let edit_sheet = match edit {
        StructuralEdit::InsertRows { sheet, .. }
        | StructuralEdit::DeleteRows { sheet, .. }
        | StructuralEdit::InsertCols { sheet, .. }
        | StructuralEdit::DeleteCols { sheet, .. } => sheet.as_str(),
    };
    let sheet_ref = start.sheet.as_ref().or(end.sheet.as_ref());
    if !sheet_ref_applies_for_sheet_edit(
        sheet_ref,
        ctx_sheet,
        edit_sheet,
        resolve_sheet_order_index,
    ) {
        return (original.clone(), false);
    }

    let Some((start_col, start_col_abs)) = coord_a1(&start.col) else {
        return (original.clone(), false);
    };
    let Some((start_row, start_row_abs)) = coord_a1(&start.row) else {
        return (original.clone(), false);
    };
    let Some((end_col, end_col_abs)) = coord_a1(&end.col) else {
        return (original.clone(), false);
    };
    let Some((end_row, end_row_abs)) = coord_a1(&end.row) else {
        return (original.clone(), false);
    };

    let mut sr = min(start_row, end_row);
    let mut er = max(start_row, end_row);
    let mut sc = min(start_col, end_col);
    let mut ec = max(start_col, end_col);

    match edit {
        StructuralEdit::InsertRows { row: at, count, .. } => {
            sr = adjust_row_insert(sr, *at, *count).unwrap_or(sr);
            er = adjust_row_insert(er, *at, *count).unwrap_or(er);
        }
        StructuralEdit::DeleteRows { row: at, count, .. } => {
            let del_end = at.saturating_add(count.saturating_sub(1));
            let Some((new_sr, new_er)) = adjust_row_range_delete(sr, er, *at, del_end, *count)
            else {
                return (Expr::Error(REF_ERROR.to_string()), true);
            };
            sr = new_sr;
            er = new_er;
        }
        StructuralEdit::InsertCols { col: at, count, .. } => {
            sc = adjust_col_insert(sc, *at, *count).unwrap_or(sc);
            ec = adjust_col_insert(ec, *at, *count).unwrap_or(ec);
        }
        StructuralEdit::DeleteCols { col: at, count, .. } => {
            let del_end = at.saturating_add(count.saturating_sub(1));
            let Some((new_sc, new_ec)) = adjust_col_range_delete(sc, ec, *at, del_end, *count)
            else {
                return (Expr::Error(REF_ERROR.to_string()), true);
            };
            sc = new_sc;
            ec = new_ec;
        }
    }

    let start_ref = AstCellRef {
        workbook: start.workbook.clone(),
        sheet: start.sheet.clone(),
        col: Coord::A1 {
            index: sc,
            abs: start_col_abs,
        },
        row: Coord::A1 {
            index: sr,
            abs: start_row_abs,
        },
    };

    if sr == er && sc == ec {
        let out = expr_ref(start_ref);
        return (out.clone(), out != *original);
    }

    let end_ref = AstCellRef {
        workbook: end.workbook.clone(),
        sheet: end.sheet.clone(),
        col: Coord::A1 {
            index: ec,
            abs: end_col_abs,
        },
        row: Coord::A1 {
            index: er,
            abs: end_row_abs,
        },
    };

    let out = Expr::Binary(BinaryExpr {
        op: BinaryOp::Range,
        left: Box::new(expr_ref(start_ref)),
        right: Box::new(expr_ref(end_ref)),
    });

    (out.clone(), out != *original)
}

fn rewrite_row_range_for_structural_edit(
    original: &Expr,
    start: &AstRowRef,
    end: &AstRowRef,
    ctx_sheet: &str,
    edit: &StructuralEdit,
    resolve_sheet_order_index: &mut impl FnMut(&str) -> Option<usize>,
) -> (Expr, bool) {
    if start.workbook.is_some() || end.workbook.is_some() {
        return (original.clone(), false);
    }
    let edit_sheet = match edit {
        StructuralEdit::InsertRows { sheet, .. } | StructuralEdit::DeleteRows { sheet, .. } => {
            sheet.as_str()
        }
        _ => return (original.clone(), false),
    };
    let sheet_ref = start.sheet.as_ref().or(end.sheet.as_ref());
    if !sheet_ref_applies_for_sheet_edit(
        sheet_ref,
        ctx_sheet,
        edit_sheet,
        resolve_sheet_order_index,
    ) {
        return (original.clone(), false);
    }

    let Some((start_row, start_abs)) = coord_a1(&start.row) else {
        return (original.clone(), false);
    };
    let Some((end_row, end_abs)) = coord_a1(&end.row) else {
        return (original.clone(), false);
    };

    let mut sr = min(start_row, end_row);
    let mut er = max(start_row, end_row);

    match edit {
        StructuralEdit::InsertRows { row: at, count, .. } => {
            sr = adjust_row_insert(sr, *at, *count).unwrap_or(sr);
            er = adjust_row_insert(er, *at, *count).unwrap_or(er);
        }
        StructuralEdit::DeleteRows { row: at, count, .. } => {
            let del_end = at.saturating_add(count.saturating_sub(1));
            let Some((new_sr, new_er)) = adjust_row_range_delete(sr, er, *at, del_end, *count)
            else {
                return (Expr::Error(REF_ERROR.to_string()), true);
            };
            sr = new_sr;
            er = new_er;
        }
        _ => {}
    }

    let out = Expr::Binary(BinaryExpr {
        op: BinaryOp::Range,
        left: Box::new(Expr::RowRef(AstRowRef {
            workbook: start.workbook.clone(),
            sheet: start.sheet.clone(),
            row: Coord::A1 {
                index: sr,
                abs: start_abs,
            },
        })),
        right: Box::new(Expr::RowRef(AstRowRef {
            workbook: end.workbook.clone(),
            sheet: end.sheet.clone(),
            row: Coord::A1 {
                index: er,
                abs: end_abs,
            },
        })),
    });

    (out.clone(), out != *original)
}

fn rewrite_col_range_for_structural_edit(
    original: &Expr,
    start: &AstColRef,
    end: &AstColRef,
    ctx_sheet: &str,
    edit: &StructuralEdit,
    resolve_sheet_order_index: &mut impl FnMut(&str) -> Option<usize>,
) -> (Expr, bool) {
    if start.workbook.is_some() || end.workbook.is_some() {
        return (original.clone(), false);
    }
    let edit_sheet = match edit {
        StructuralEdit::InsertCols { sheet, .. } | StructuralEdit::DeleteCols { sheet, .. } => {
            sheet.as_str()
        }
        _ => return (original.clone(), false),
    };
    let sheet_ref = start.sheet.as_ref().or(end.sheet.as_ref());
    if !sheet_ref_applies_for_sheet_edit(
        sheet_ref,
        ctx_sheet,
        edit_sheet,
        resolve_sheet_order_index,
    ) {
        return (original.clone(), false);
    }

    let Some((start_col, start_abs)) = coord_a1(&start.col) else {
        return (original.clone(), false);
    };
    let Some((end_col, end_abs)) = coord_a1(&end.col) else {
        return (original.clone(), false);
    };

    let mut sc = min(start_col, end_col);
    let mut ec = max(start_col, end_col);

    match edit {
        StructuralEdit::InsertCols { col: at, count, .. } => {
            sc = adjust_col_insert(sc, *at, *count).unwrap_or(sc);
            ec = adjust_col_insert(ec, *at, *count).unwrap_or(ec);
        }
        StructuralEdit::DeleteCols { col: at, count, .. } => {
            let del_end = at.saturating_add(count.saturating_sub(1));
            let Some((new_sc, new_ec)) = adjust_col_range_delete(sc, ec, *at, del_end, *count)
            else {
                return (Expr::Error(REF_ERROR.to_string()), true);
            };
            sc = new_sc;
            ec = new_ec;
        }
        _ => {}
    }

    let out = Expr::Binary(BinaryExpr {
        op: BinaryOp::Range,
        left: Box::new(Expr::ColRef(AstColRef {
            workbook: start.workbook.clone(),
            sheet: start.sheet.clone(),
            col: Coord::A1 {
                index: sc,
                abs: start_abs,
            },
        })),
        right: Box::new(Expr::ColRef(AstColRef {
            workbook: end.workbook.clone(),
            sheet: end.sheet.clone(),
            col: Coord::A1 {
                index: ec,
                abs: end_abs,
            },
        })),
    });

    (out.clone(), out != *original)
}

fn rewrite_range_for_copy_delta(
    original: &Expr,
    b: &BinaryExpr,
    delta_row: i32,
    delta_col: i32,
) -> Option<(Expr, bool)> {
    match (&*b.left, &*b.right) {
        (Expr::CellRef(start), Expr::CellRef(end)) => Some(rewrite_cell_range_for_copy_delta(
            original, start, end, delta_row, delta_col,
        )),
        (Expr::RowRef(start), Expr::RowRef(end)) => Some(rewrite_row_range_for_copy_delta(
            original, start, end, delta_row,
        )),
        (Expr::ColRef(start), Expr::ColRef(end)) => Some(rewrite_col_range_for_copy_delta(
            original, start, end, delta_col,
        )),
        _ => None,
    }
}

fn rewrite_cell_range_for_copy_delta(
    original: &Expr,
    start: &AstCellRef,
    end: &AstCellRef,
    delta_row: i32,
    delta_col: i32,
) -> (Expr, bool) {
    let Some((start_col, start_col_abs)) = coord_a1(&start.col) else {
        return (original.clone(), false);
    };
    let Some((start_row, start_row_abs)) = coord_a1(&start.row) else {
        return (original.clone(), false);
    };
    let Some((end_col, end_col_abs)) = coord_a1(&end.col) else {
        return (original.clone(), false);
    };
    let Some((end_row, end_row_abs)) = coord_a1(&end.row) else {
        return (original.clone(), false);
    };

    let new_start_row = if start_row_abs {
        start_row as i64
    } else {
        start_row as i64 + delta_row as i64
    };
    let new_start_col = if start_col_abs {
        start_col as i64
    } else {
        start_col as i64 + delta_col as i64
    };
    let new_end_row = if end_row_abs {
        end_row as i64
    } else {
        end_row as i64 + delta_row as i64
    };
    let new_end_col = if end_col_abs {
        end_col as i64
    } else {
        end_col as i64 + delta_col as i64
    };

    if new_start_row < 0 || new_start_col < 0 || new_end_row < 0 || new_end_col < 0 {
        return (Expr::Error(REF_ERROR.to_string()), true);
    }

    let out = Expr::Binary(BinaryExpr {
        op: BinaryOp::Range,
        left: Box::new(expr_ref(AstCellRef {
            workbook: start.workbook.clone(),
            sheet: start.sheet.clone(),
            col: Coord::A1 {
                index: new_start_col as u32,
                abs: start_col_abs,
            },
            row: Coord::A1 {
                index: new_start_row as u32,
                abs: start_row_abs,
            },
        })),
        right: Box::new(expr_ref(AstCellRef {
            workbook: end.workbook.clone(),
            sheet: end.sheet.clone(),
            col: Coord::A1 {
                index: new_end_col as u32,
                abs: end_col_abs,
            },
            row: Coord::A1 {
                index: new_end_row as u32,
                abs: end_row_abs,
            },
        })),
    });

    (out.clone(), out != *original)
}

fn rewrite_row_range_for_copy_delta(
    original: &Expr,
    start: &AstRowRef,
    end: &AstRowRef,
    delta_row: i32,
) -> (Expr, bool) {
    let Some((start_row, start_abs)) = coord_a1(&start.row) else {
        return (original.clone(), false);
    };
    let Some((end_row, end_abs)) = coord_a1(&end.row) else {
        return (original.clone(), false);
    };

    let new_start_row = if start_abs {
        start_row as i64
    } else {
        start_row as i64 + delta_row as i64
    };
    let new_end_row = if end_abs {
        end_row as i64
    } else {
        end_row as i64 + delta_row as i64
    };

    if new_start_row < 0 || new_end_row < 0 {
        return (Expr::Error(REF_ERROR.to_string()), true);
    }

    let out = Expr::Binary(BinaryExpr {
        op: BinaryOp::Range,
        left: Box::new(Expr::RowRef(AstRowRef {
            workbook: start.workbook.clone(),
            sheet: start.sheet.clone(),
            row: Coord::A1 {
                index: new_start_row as u32,
                abs: start_abs,
            },
        })),
        right: Box::new(Expr::RowRef(AstRowRef {
            workbook: end.workbook.clone(),
            sheet: end.sheet.clone(),
            row: Coord::A1 {
                index: new_end_row as u32,
                abs: end_abs,
            },
        })),
    });

    (out.clone(), out != *original)
}

fn rewrite_col_range_for_copy_delta(
    original: &Expr,
    start: &AstColRef,
    end: &AstColRef,
    delta_col: i32,
) -> (Expr, bool) {
    let Some((start_col, start_abs)) = coord_a1(&start.col) else {
        return (original.clone(), false);
    };
    let Some((end_col, end_abs)) = coord_a1(&end.col) else {
        return (original.clone(), false);
    };

    let new_start_col = if start_abs {
        start_col as i64
    } else {
        start_col as i64 + delta_col as i64
    };
    let new_end_col = if end_abs {
        end_col as i64
    } else {
        end_col as i64 + delta_col as i64
    };

    if new_start_col < 0 || new_end_col < 0 {
        return (Expr::Error(REF_ERROR.to_string()), true);
    }

    let out = Expr::Binary(BinaryExpr {
        op: BinaryOp::Range,
        left: Box::new(Expr::ColRef(AstColRef {
            workbook: start.workbook.clone(),
            sheet: start.sheet.clone(),
            col: Coord::A1 {
                index: new_start_col as u32,
                abs: start_abs,
            },
        })),
        right: Box::new(Expr::ColRef(AstColRef {
            workbook: end.workbook.clone(),
            sheet: end.sheet.clone(),
            col: Coord::A1 {
                index: new_end_col as u32,
                abs: end_abs,
            },
        })),
    });

    (out.clone(), out != *original)
}

fn rewrite_cell_ref_for_copy_delta(r: &AstCellRef, delta_row: i32, delta_col: i32) -> (Expr, bool) {
    let Some((col, col_abs)) = coord_a1(&r.col) else {
        return (expr_ref(r.clone()), false);
    };
    let Some((row, row_abs)) = coord_a1(&r.row) else {
        return (expr_ref(r.clone()), false);
    };

    let new_row = if row_abs {
        row as i64
    } else {
        row as i64 + delta_row as i64
    };
    let new_col = if col_abs {
        col as i64
    } else {
        col as i64 + delta_col as i64
    };

    if new_row < 0 || new_col < 0 {
        return (Expr::Error(REF_ERROR.to_string()), true);
    }

    let new_ref = AstCellRef {
        workbook: r.workbook.clone(),
        sheet: r.sheet.clone(),
        col: Coord::A1 {
            index: new_col as u32,
            abs: col_abs,
        },
        row: Coord::A1 {
            index: new_row as u32,
            abs: row_abs,
        },
    };

    let changed = &new_ref != r;
    (expr_ref(new_ref), changed)
}

fn rewrite_row_ref_for_copy_delta(r: &AstRowRef, delta_row: i32) -> (Expr, bool) {
    let Some((row, abs)) = coord_a1(&r.row) else {
        return (Expr::RowRef(r.clone()), false);
    };

    let new_row = if abs {
        row as i64
    } else {
        row as i64 + delta_row as i64
    };
    if new_row < 0 {
        return (Expr::Error(REF_ERROR.to_string()), true);
    }

    let new_ref = AstRowRef {
        workbook: r.workbook.clone(),
        sheet: r.sheet.clone(),
        row: Coord::A1 {
            index: new_row as u32,
            abs,
        },
    };

    let changed = &new_ref != r;
    (Expr::RowRef(new_ref), changed)
}

fn rewrite_col_ref_for_copy_delta(r: &AstColRef, delta_col: i32) -> (Expr, bool) {
    let Some((col, abs)) = coord_a1(&r.col) else {
        return (Expr::ColRef(r.clone()), false);
    };

    let new_col = if abs {
        col as i64
    } else {
        col as i64 + delta_col as i64
    };
    if new_col < 0 {
        return (Expr::Error(REF_ERROR.to_string()), true);
    }

    let new_ref = AstColRef {
        workbook: r.workbook.clone(),
        sheet: r.sheet.clone(),
        col: Coord::A1 {
            index: new_col as u32,
            abs,
        },
    };

    let changed = &new_ref != r;
    (Expr::ColRef(new_ref), changed)
}

fn rewrite_cell_ref_for_range_map(
    r: &AstCellRef,
    ctx_sheet: &str,
    edit: &RangeMapEdit,
    resolve_sheet_order_index: &mut impl FnMut(&str) -> Option<usize>,
) -> (Expr, bool) {
    if r.workbook.is_some() {
        return (expr_ref(r.clone()), false);
    }
    if !sheet_ref_applies_for_sheet_edit(
        r.sheet.as_ref(),
        ctx_sheet,
        &edit.sheet,
        resolve_sheet_order_index,
    ) {
        return (expr_ref(r.clone()), false);
    }

    let Some((col, col_abs)) = coord_a1(&r.col) else {
        return (expr_ref(r.clone()), false);
    };
    let Some((row, row_abs)) = coord_a1(&r.row) else {
        return (expr_ref(r.clone()), false);
    };

    if let Some(deleted) = edit.deleted_region {
        if deleted.contains(row, col) {
            return (Expr::Error(REF_ERROR.to_string()), true);
        }
    }

    if !edit.moved_region.contains(row, col) {
        return (expr_ref(r.clone()), false);
    }

    let new_row = row as i64 + edit.delta_row as i64;
    let new_col = col as i64 + edit.delta_col as i64;
    if new_row < 0 || new_col < 0 {
        return (Expr::Error(REF_ERROR.to_string()), true);
    }

    let new_ref = AstCellRef {
        workbook: r.workbook.clone(),
        sheet: r.sheet.clone(),
        col: Coord::A1 {
            index: new_col as u32,
            abs: col_abs,
        },
        row: Coord::A1 {
            index: new_row as u32,
            abs: row_abs,
        },
    };

    let changed = &new_ref != r;
    (expr_ref(new_ref), changed)
}

fn rewrite_cell_range_for_range_map(
    original: &Expr,
    b: &BinaryExpr,
    ctx_sheet: &str,
    edit: &RangeMapEdit,
    resolve_sheet_order_index: &mut impl FnMut(&str) -> Option<usize>,
) -> Option<(Expr, bool)> {
    let Expr::CellRef(start) = b.left.as_ref() else {
        return None;
    };
    let Expr::CellRef(end) = b.right.as_ref() else {
        return None;
    };

    if start.workbook.is_some() || end.workbook.is_some() {
        return Some((original.clone(), false));
    }
    let sheet_ref = start.sheet.as_ref().or(end.sheet.as_ref());
    if !sheet_ref_applies_for_sheet_edit(
        sheet_ref,
        ctx_sheet,
        &edit.sheet,
        resolve_sheet_order_index,
    ) {
        return Some((original.clone(), false));
    }

    let Some((start_col, start_col_abs)) = coord_a1(&start.col) else {
        return Some((original.clone(), false));
    };
    let Some((start_row, start_row_abs)) = coord_a1(&start.row) else {
        return Some((original.clone(), false));
    };
    let Some((end_col, end_col_abs)) = coord_a1(&end.col) else {
        return Some((original.clone(), false));
    };
    let Some((end_row, end_row_abs)) = coord_a1(&end.row) else {
        return Some((original.clone(), false));
    };

    let original_range = GridRange::new(start_row, start_col, end_row, end_col);
    let mut areas = vec![original_range];

    if let Some(deleted) = edit.deleted_region {
        areas = subtract_region(&areas, deleted);
        if areas.is_empty() {
            return Some((Expr::Error(REF_ERROR.to_string()), true));
        }
    }

    areas = apply_move_region(&areas, edit.moved_region, edit.delta_row, edit.delta_col);
    if areas.is_empty() {
        return Some((Expr::Error(REF_ERROR.to_string()), true));
    }

    let exprs: Vec<Expr> = areas
        .into_iter()
        .map(|area| {
            if area.start_row == area.end_row && area.start_col == area.end_col {
                expr_ref(AstCellRef {
                    workbook: start.workbook.clone(),
                    sheet: start.sheet.clone(),
                    col: Coord::A1 {
                        index: area.start_col,
                        abs: start_col_abs,
                    },
                    row: Coord::A1 {
                        index: area.start_row,
                        abs: start_row_abs,
                    },
                })
            } else {
                Expr::Binary(BinaryExpr {
                    op: BinaryOp::Range,
                    left: Box::new(expr_ref(AstCellRef {
                        workbook: start.workbook.clone(),
                        sheet: start.sheet.clone(),
                        col: Coord::A1 {
                            index: area.start_col,
                            abs: start_col_abs,
                        },
                        row: Coord::A1 {
                            index: area.start_row,
                            abs: start_row_abs,
                        },
                    })),
                    right: Box::new(expr_ref(AstCellRef {
                        workbook: end.workbook.clone(),
                        sheet: end.sheet.clone(),
                        col: Coord::A1 {
                            index: area.end_col,
                            abs: end_col_abs,
                        },
                        row: Coord::A1 {
                            index: area.end_row,
                            abs: end_row_abs,
                        },
                    })),
                })
            }
        })
        .collect();

    let out = union_expr(exprs);
    let changed = out != *original;
    Some((out, changed))
}

fn union_expr(mut exprs: Vec<Expr>) -> Expr {
    match exprs.len() {
        0 => Expr::Error(REF_ERROR.to_string()),
        1 => exprs.pop().expect("len == 1"),
        _ => {
            let mut iter = exprs.into_iter();
            let mut out = iter.next().expect("len > 1");
            for next in iter {
                out = Expr::Binary(BinaryExpr {
                    op: BinaryOp::Union,
                    left: Box::new(out),
                    right: Box::new(next),
                });
            }
            out
        }
    }
}

fn expr_ref(r: AstCellRef) -> Expr {
    Expr::CellRef(r)
}

fn coord_a1(coord: &Coord) -> Option<(u32, bool)> {
    match coord {
        Coord::A1 { index, abs } => Some((*index, *abs)),
        Coord::Offset(_) => None,
    }
}

fn adjust_row_insert(row: u32, at: u32, count: u32) -> Option<u32> {
    if row >= at {
        row.checked_add(count)
    } else {
        Some(row)
    }
}

fn adjust_col_insert(col: u32, at: u32, count: u32) -> Option<u32> {
    if col >= at {
        col.checked_add(count)
    } else {
        Some(col)
    }
}

fn adjust_row_delete(row: u32, del_start: u32, del_end: u32, count: u32) -> Option<u32> {
    if row < del_start {
        Some(row)
    } else if row > del_end {
        Some(row - count)
    } else {
        None
    }
}

fn adjust_col_delete(col: u32, del_start: u32, del_end: u32, count: u32) -> Option<u32> {
    if col < del_start {
        Some(col)
    } else if col > del_end {
        Some(col - count)
    } else {
        None
    }
}

fn adjust_row_range_delete(
    start: u32,
    end: u32,
    del_start: u32,
    del_end: u32,
    count: u32,
) -> Option<(u32, u32)> {
    if end < del_start {
        return Some((start, end));
    }
    if start > del_end {
        return Some((start - count, end - count));
    }
    if start >= del_start && end <= del_end {
        return None;
    }

    let mut new_start = start;
    let mut new_end = end;

    if start >= del_start && start <= del_end {
        new_start = del_start;
    }

    if end >= del_start && end <= del_end {
        if del_start == 0 {
            return None;
        }
        new_end = del_start - 1;
    } else if end > del_end {
        new_end = end - count;
    }

    if new_start > new_end {
        None
    } else {
        Some((new_start, new_end))
    }
}

fn adjust_col_range_delete(
    start: u32,
    end: u32,
    del_start: u32,
    del_end: u32,
    count: u32,
) -> Option<(u32, u32)> {
    if end < del_start {
        return Some((start, end));
    }
    if start > del_end {
        return Some((start - count, end - count));
    }
    if start >= del_start && end <= del_end {
        return None;
    }

    let mut new_start = start;
    let mut new_end = end;

    if start >= del_start && start <= del_end {
        new_start = del_start;
    }

    if end >= del_start && end <= del_end {
        if del_start == 0 {
            return None;
        }
        new_end = del_start - 1;
    } else if end > del_end {
        new_end = end - count;
    }

    if new_start > new_end {
        None
    } else {
        Some((new_start, new_end))
    }
}

fn subtract_region(areas: &[GridRange], deleted: GridRange) -> Vec<GridRange> {
    let mut out = Vec::new();
    for area in areas {
        let overlap = rect_intersection(*area, deleted);
        match overlap {
            None => out.push(*area),
            Some(over) if over == *area => {}
            Some(over) => out.extend(rect_difference(*area, over)),
        }
    }
    out
}

fn apply_move_region(
    areas: &[GridRange],
    moved_region: GridRange,
    delta_row: i32,
    delta_col: i32,
) -> Vec<GridRange> {
    let mut out = Vec::new();
    for area in areas {
        let overlap = rect_intersection(*area, moved_region);
        match overlap {
            None => out.push(*area),
            Some(over) if over == *area => {
                if let Some(shifted) = shift_range(*area, delta_row, delta_col) {
                    out.push(shifted);
                }
            }
            Some(over) => {
                out.extend(rect_difference(*area, over));
                if let Some(shifted) = shift_range(over, delta_row, delta_col) {
                    out.push(shifted);
                }
            }
        }
    }
    out
}

fn rect_intersection(a: GridRange, b: GridRange) -> Option<GridRange> {
    let start_row = max(a.start_row, b.start_row);
    let start_col = max(a.start_col, b.start_col);
    let end_row = min(a.end_row, b.end_row);
    let end_col = min(a.end_col, b.end_col);
    if start_row > end_row || start_col > end_col {
        None
    } else {
        Some(GridRange::new(start_row, start_col, end_row, end_col))
    }
}

fn rect_difference(range: GridRange, overlap: GridRange) -> Vec<GridRange> {
    let mut out = Vec::new();
    if range.start_row < overlap.start_row {
        out.push(GridRange::new(
            range.start_row,
            range.start_col,
            overlap.start_row - 1,
            range.end_col,
        ));
    }
    if overlap.end_row < range.end_row {
        out.push(GridRange::new(
            overlap.end_row + 1,
            range.start_col,
            range.end_row,
            range.end_col,
        ));
    }
    let mid_start_row = overlap.start_row;
    let mid_end_row = overlap.end_row;
    if range.start_col < overlap.start_col {
        out.push(GridRange::new(
            mid_start_row,
            range.start_col,
            mid_end_row,
            overlap.start_col - 1,
        ));
    }
    if overlap.end_col < range.end_col {
        out.push(GridRange::new(
            mid_start_row,
            overlap.end_col + 1,
            mid_end_row,
            range.end_col,
        ));
    }
    out
}

fn shift_range(area: GridRange, delta_row: i32, delta_col: i32) -> Option<GridRange> {
    let start_row = area.start_row as i64 + delta_row as i64;
    let start_col = area.start_col as i64 + delta_col as i64;
    let end_row = area.end_row as i64 + delta_row as i64;
    let end_col = area.end_col as i64 + delta_col as i64;
    if start_row < 0 || start_col < 0 || end_row < 0 || end_col < 0 {
        return None;
    }
    Some(GridRange::new(
        start_row as u32,
        start_col as u32,
        end_row as u32,
        end_col as u32,
    ))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn copy_delta_updates_row_and_column_ranges() {
        let (out, changed) =
            rewrite_formula_for_copy_delta("=A:A", "Sheet1", CellAddr::new(0, 0), 0, 1);
        assert!(changed);
        assert_eq!(out, "=B:B");

        let (out, changed) =
            rewrite_formula_for_copy_delta("=1:1", "Sheet1", CellAddr::new(0, 0), 1, 0);
        assert!(changed);
        assert_eq!(out, "=2:2");
    }

    #[test]
    fn copy_delta_turns_entire_range_into_ref_error_if_any_endpoint_underflows() {
        let (out, changed) =
            rewrite_formula_for_copy_delta("=A1:B2", "Sheet1", CellAddr::new(0, 0), -1, 0);
        assert!(changed);
        assert_eq!(out, "=#REF!");

        let (out, changed) =
            rewrite_formula_for_copy_delta("=1:2", "Sheet1", CellAddr::new(0, 0), -1, 0);
        assert!(changed);
        assert_eq!(out, "=#REF!");

        let (out, changed) =
            rewrite_formula_for_copy_delta("=A:B", "Sheet1", CellAddr::new(0, 0), 0, -1);
        assert!(changed);
        assert_eq!(out, "=#REF!");
    }

    #[test]
    fn structural_edit_rewrites_sheet_qualified_ranges() {
        let edit = StructuralEdit::InsertRows {
            sheet: "Sheet1".to_string(),
            row: 0,
            count: 1,
        };
        let (out, changed) = rewrite_formula_for_structural_edit(
            "=SUM(Sheet1!A1:B2)",
            "Other",
            CellAddr::new(0, 0),
            &edit,
        );
        assert!(changed);
        assert_eq!(out, "=SUM(Sheet1!A2:B3)");
    }

    #[test]
    fn structural_edit_updates_column_ranges_and_preserves_ref_errors() {
        let insert = StructuralEdit::InsertCols {
            sheet: "Sheet1".to_string(),
            col: 0,
            count: 1,
        };
        let (out, changed) =
            rewrite_formula_for_structural_edit("=A:A", "Sheet1", CellAddr::new(0, 0), &insert);
        assert!(changed);
        assert_eq!(out, "=B:B");

        let delete = StructuralEdit::DeleteCols {
            sheet: "Sheet1".to_string(),
            col: 0,
            count: 1,
        };
        let (out, changed) =
            rewrite_formula_for_structural_edit("=A:B", "Sheet1", CellAddr::new(0, 0), &delete);
        assert!(changed);
        assert_eq!(out, "=A:A");

        let (out, changed) =
            rewrite_formula_for_structural_edit("=A:A", "Sheet1", CellAddr::new(0, 0), &delete);
        assert!(changed);
        assert_eq!(out, "=#REF!");
    }

    #[test]
    fn external_workbook_refs_are_not_rewritten() {
        let edit = StructuralEdit::InsertRows {
            sheet: "Sheet1".to_string(),
            row: 0,
            count: 1,
        };
        let (out, changed) = rewrite_formula_for_structural_edit(
            "=[Book.xlsx]Sheet1!A1",
            "Sheet1",
            CellAddr::new(0, 0),
            &edit,
        );
        assert!(!changed);
        assert_eq!(out, "=[Book.xlsx]Sheet1!A1");
    }

    #[test]
    fn range_map_can_split_ranges_into_union() {
        let edit = RangeMapEdit {
            sheet: "Sheet1".to_string(),
            moved_region: GridRange::new(10, 10, 10, 10),
            delta_row: 0,
            delta_col: 0,
            deleted_region: Some(GridRange::new(0, 1, 0, 1)), // delete B1
        };
        let (out, changed) =
            rewrite_formula_for_range_map("=A1:C1", "Sheet1", CellAddr::new(0, 0), &edit);
        assert!(changed);
        assert_eq!(out, "=A1,C1");
    }

    #[test]
    fn range_map_moves_cell_references() {
        let edit = RangeMapEdit {
            sheet: "Sheet1".to_string(),
            moved_region: GridRange::new(0, 0, 0, 0),
            delta_row: 1,
            delta_col: 1,
            deleted_region: None,
        };
        let (out, changed) =
            rewrite_formula_for_range_map("=A1", "Sheet1", CellAddr::new(0, 0), &edit);
        assert!(changed);
        assert_eq!(out, "=B2");
    }

    #[test]
    fn rewrite_preserves_leading_whitespace_before_equals() {
        let (out, changed) =
            rewrite_formula_for_copy_delta("   =A1", "Sheet1", CellAddr::new(0, 0), 1, 0);
        assert!(changed);
        assert_eq!(out, "   =A2");
    }

    #[test]
    fn range_map_splits_ranges_inside_function_args_with_parentheses() {
        // Simulate `DeleteCellsShiftLeft` of cell B1: delete B1 and shift everything right of it
        // left by 1. The range `A1:C1` becomes a union of A1 and (moved) B1, which must be
        // parenthesized inside `SUM(...)` to avoid being parsed as multiple arguments.
        let edit = RangeMapEdit {
            sheet: "Sheet1".to_string(),
            moved_region: GridRange::new(0, 2, 0, u32::MAX), // C1 and beyond shift left
            delta_row: 0,
            delta_col: -1,
            deleted_region: Some(GridRange::new(0, 1, 0, 1)), // delete B1
        };
        let (out, changed) =
            rewrite_formula_for_range_map("=SUM(A1:C1)", "Sheet1", CellAddr::new(0, 0), &edit);
        assert!(changed);
        assert_eq!(out, "=SUM((A1,B1))");
    }

    #[test]
    fn range_map_splits_ranges_inside_call_expr_args_with_parentheses() {
        // Same scenario as `range_map_splits_ranges_inside_function_args_with_parentheses`, but
        // the range appears as an argument to a postfix call expression (lambda invocation).
        let edit = RangeMapEdit {
            sheet: "Sheet1".to_string(),
            moved_region: GridRange::new(0, 2, 0, u32::MAX), // C1 and beyond shift left
            delta_row: 0,
            delta_col: -1,
            deleted_region: Some(GridRange::new(0, 1, 0, 1)), // delete B1
        };
        let (out, changed) = rewrite_formula_for_range_map(
            "=LAMBDA(x,x)(A1:C1)",
            "Sheet1",
            CellAddr::new(0, 0),
            &edit,
        );
        assert!(changed);
        assert_eq!(out, "=LAMBDA(x,x)((A1,B1))");
    }

    #[test]
    fn spill_postfix_is_dropped_when_reference_becomes_ref_error() {
        let delete_col = StructuralEdit::DeleteCols {
            sheet: "Sheet1".to_string(),
            col: 0,
            count: 1,
        };
        let (out, changed) =
            rewrite_formula_for_structural_edit("=A1#", "Sheet1", CellAddr::new(0, 0), &delete_col);
        assert!(changed);
        assert_eq!(out, "=#REF!");

        let (out, changed) =
            rewrite_formula_for_copy_delta("=A1#", "Sheet1", CellAddr::new(0, 0), 0, -1);
        assert!(changed);
        assert_eq!(out, "=#REF!");

        let range_map = RangeMapEdit {
            sheet: "Sheet1".to_string(),
            moved_region: GridRange::new(0, 0, u32::MAX, u32::MAX),
            delta_row: 0,
            delta_col: 0,
            deleted_region: Some(GridRange::new(0, 0, 0, 0)),
        };
        let (out, changed) =
            rewrite_formula_for_range_map("=A1#", "Sheet1", CellAddr::new(0, 0), &range_map);
        assert!(changed);
        assert_eq!(out, "=#REF!");
    }

    #[test]
    fn copy_delta_rewrites_references_inside_call_expressions() {
        let (out, changed) = rewrite_formula_for_copy_delta(
            "=LAMBDA(x,x+1)(A1)",
            "Sheet1",
            CellAddr::new(0, 0),
            1,
            0,
        );
        assert!(changed);
        assert_eq!(out, "=LAMBDA(x,x+1)(A2)");
    }

    #[test]
    fn copy_delta_rewrites_call_callee_references() {
        let (out, changed) = rewrite_formula_for_copy_delta(
            "=LAMBDA(x,A1+x)(1)",
            "Sheet1",
            CellAddr::new(0, 0),
            1,
            0,
        );
        assert!(changed);
        assert_eq!(out, "=LAMBDA(x,A2+x)(1)");
    }

    #[test]
    fn structural_edit_rewrites_references_inside_field_access() {
        let edit = StructuralEdit::InsertRows {
            sheet: "Sheet1".to_string(),
            row: 0,
            count: 1,
        };
        let (out, changed) =
            rewrite_formula_for_structural_edit("=A1.Price", "Sheet1", CellAddr::new(0, 1), &edit);
        assert!(changed);
        assert_eq!(out, "=A2.Price");

        let (out, changed) = rewrite_formula_for_structural_edit(
            "=A1.[Unit Price]",
            "Sheet1",
            CellAddr::new(0, 1),
            &edit,
        );
        assert!(changed);
        assert_eq!(out, r#"=A2.["Unit Price"]"#);
    }

    #[test]
    fn copy_delta_rewrites_references_inside_field_access() {
        let (out, changed) =
            rewrite_formula_for_copy_delta("=A1.Price", "Sheet1", CellAddr::new(0, 1), 1, 0);
        assert!(changed);
        assert_eq!(out, "=A2.Price");
    }

    #[test]
    fn copy_delta_drops_sheet_prefix_when_reference_becomes_ref_error() {
        // The formula grammar does not allow sheet-qualified error literals like `Sheet1!#REF!`.
        // When rewriting creates a `#REF!`, the sheet prefix should be dropped.
        let (out, changed) =
            rewrite_formula_for_copy_delta("=Sheet1!A1", "Sheet1", CellAddr::new(0, 0), 0, -1);
        assert!(changed);
        assert_eq!(out, "=#REF!");

        let (out, changed) = rewrite_formula_for_copy_delta(
            "='Sheet Name'!A1",
            "Sheet1",
            CellAddr::new(0, 0),
            0,
            -1,
        );
        assert!(changed);
        assert_eq!(out, "=#REF!");
    }

    #[test]
    fn copy_delta_rewrites_external_workbook_refs() {
        let (out, changed) = rewrite_formula_for_copy_delta(
            "=[Book.xlsx]Sheet1!A1",
            "Sheet1",
            CellAddr::new(0, 0),
            1,
            1,
        );
        assert!(changed);
        assert_eq!(out, "=[Book.xlsx]Sheet1!B2");

        // Absolute references should not move when copying.
        let (out, changed) = rewrite_formula_for_copy_delta(
            "=[Book.xlsx]Sheet1!$A$1",
            "Sheet1",
            CellAddr::new(0, 0),
            1,
            1,
        );
        assert!(!changed);
        assert_eq!(out, "=[Book.xlsx]Sheet1!$A$1");
    }
}
