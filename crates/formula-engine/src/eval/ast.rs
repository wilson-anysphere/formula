use crate::eval::address::CellAddr;
use crate::value::ErrorKind;
use std::sync::Arc;

pub type ParsedExpr = Expr<String>;
pub type CompiledExpr = Expr<usize>;

/// A 2D cell coordinate that can be either absolute (A1-style index) or relative (offset from the
/// formula origin cell).
///
/// This is designed to be a compact, "bytecode-like" representation so compiled ASTs can be
/// shared across filled formulas. The evaluator resolves relative coordinates using the current
/// evaluation cell as the base.
///
/// ## Sentinel semantics
///
/// Excel whole-row/whole-column references (e.g. `A:A` / `1:1`) are represented using a sentinel
/// "sheet end" coordinate for the open-ended axis. In the evaluation layer this is represented as
/// [`CellAddr::SHEET_END`] (`u32::MAX`). In this IR we reserve [`i32::MAX`] as the corresponding
/// sentinel for *absolute* coordinates.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Ref {
    /// Row index or offset.
    ///
    /// - When `row_abs == true`, this is an absolute 0-indexed row number, or [`Ref::SHEET_END`]
    ///   for whole-column endpoints.
    /// - When `row_abs == false`, this is a signed row offset from the formula origin cell.
    pub row: i32,
    /// Column index or offset.
    ///
    /// - When `col_abs == true`, this is an absolute 0-indexed column number, or
    ///   [`Ref::SHEET_END`] for whole-row endpoints.
    /// - When `col_abs == false`, this is a signed column offset from the formula origin cell.
    pub col: i32,
    pub row_abs: bool,
    pub col_abs: bool,
}

impl Ref {
    /// Sentinel value used for absolute coordinates to represent "the last row/column of the
    /// sheet" in whole-row/whole-column references.
    pub const SHEET_END: i32 = i32::MAX;

    /// Convert an absolute [`CellAddr`] component (`row` or `col`) into a compact i32 encoding.
    ///
    /// [`CellAddr::SHEET_END`] maps to [`Ref::SHEET_END`]. Values greater than or equal to
    /// `Ref::SHEET_END` are rejected because `Ref::SHEET_END` is reserved.
    pub fn encode_abs_component(value: u32) -> Option<i32> {
        if value == CellAddr::SHEET_END {
            return Some(Self::SHEET_END);
        }
        if value >= Self::SHEET_END as u32 {
            return None;
        }
        Some(value as i32)
    }

    /// Convert a compact i32 absolute coordinate component back into the evaluation layer's u32
    /// representation.
    ///
    /// [`Ref::SHEET_END`] maps to [`CellAddr::SHEET_END`].
    pub fn decode_abs_component(value: i32) -> Option<u32> {
        if value == Self::SHEET_END {
            return Some(CellAddr::SHEET_END);
        }
        if value < 0 {
            return None;
        }
        Some(value as u32)
    }

    /// Build a fully-absolute reference from an evaluation-layer [`CellAddr`].
    pub fn from_abs_cell_addr(addr: CellAddr) -> Option<Self> {
        Some(Self {
            row: Self::encode_abs_component(addr.row)?,
            col: Self::encode_abs_component(addr.col)?,
            row_abs: true,
            col_abs: true,
        })
    }

    /// If this reference is fully absolute, decode it into a [`CellAddr`].
    pub fn as_abs_cell_addr(&self) -> Option<CellAddr> {
        if !self.row_abs || !self.col_abs {
            return None;
        }
        Some(CellAddr {
            row: Self::decode_abs_component(self.row)?,
            col: Self::decode_abs_component(self.col)?,
        })
    }

    /// Resolve this reference relative to `base` (the formula origin cell).
    ///
    /// Returns `None` when:
    /// - A relative offset produces a negative coordinate.
    /// - The computation overflows `u32`.
    /// - The resolved coordinate lands on [`CellAddr::SHEET_END`] via relative arithmetic (the
    ///   sentinel is reserved for absolute whole-row/whole-column endpoints).
    pub fn resolve(&self, base: CellAddr) -> Option<CellAddr> {
        let row = if self.row_abs {
            Self::decode_abs_component(self.row)?
        } else {
            base.row.checked_add_signed(self.row)?
        };
        let col = if self.col_abs {
            Self::decode_abs_component(self.col)?
        } else {
            base.col.checked_add_signed(self.col)?
        };

        if !self.row_abs && row == CellAddr::SHEET_END {
            return None;
        }
        if !self.col_abs && col == CellAddr::SHEET_END {
            return None;
        }

        Some(CellAddr { row, col })
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum SheetReference<S> {
    Current,
    Sheet(S),
    /// 3D sheet span reference like `Sheet1:Sheet3!A1`.
    ///
    /// Resolution is based on workbook sheet order.
    SheetRange(S, S),
    /// External workbook reference like `[Book.xlsx]Sheet1!A1`.
    ///
    /// The evaluator resolves these through an external value provider (if configured).
    /// Missing external links evaluate to `#REF!`.
    External(String),
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct CellRef<S> {
    pub sheet: SheetReference<S>,
    pub addr: Ref,
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct RangeRef<S> {
    pub sheet: SheetReference<S>,
    pub start: Ref,
    pub end: Ref,
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct StructuredRefExpr<S> {
    pub sheet: SheetReference<S>,
    pub sref: crate::structured_refs::StructuredRef,
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct NameRef<S> {
    pub sheet: SheetReference<S>,
    pub name: String,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum UnaryOp {
    Plus,
    Minus,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PostfixOp {
    Percent,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum BinaryOp {
    /// The range operator (`:`).
    Range,
    /// The intersection operator (space).
    Intersect,
    /// The union operator (`,`; lexed separately from function argument separators).
    Union,
    /// Exponentiation (`^`).
    Pow,
    Add,
    Sub,
    Mul,
    Div,
    /// String concatenation (`&`).
    Concat,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum CompareOp {
    Eq,
    Ne,
    Lt,
    Le,
    Gt,
    Ge,
}

#[derive(Debug, Clone, PartialEq)]
pub enum Expr<S> {
    Number(f64),
    Text(String),
    Bool(bool),
    Blank,
    Error(ErrorKind),
    /// Array literal constant written as `{...}`.
    ///
    /// The values are stored in row-major order and each element is a compiled sub-expression
    /// evaluated at runtime.
    ArrayLiteral {
        rows: usize,
        cols: usize,
        values: Arc<[Expr<S>]>,
    },
    NameRef(NameRef<S>),
    CellRef(CellRef<S>),
    RangeRef(RangeRef<S>),
    StructuredRef(StructuredRefExpr<S>),
    /// Postfix field access on a scalar value, e.g. `A1.Price` or `A1.["Change%"]`.
    ///
    /// This is currently intended for Excel "entity" / rich value fields. The `field` string is
    /// preserved exactly from the canonical parser (excluding the leading dot).
    FieldAccess {
        base: Box<Expr<S>>,
        field: String,
    },
    Unary {
        op: UnaryOp,
        expr: Box<Expr<S>>,
    },
    Postfix {
        op: PostfixOp,
        expr: Box<Expr<S>>,
    },
    Binary {
        op: BinaryOp,
        left: Box<Expr<S>>,
        right: Box<Expr<S>>,
    },
    Compare {
        op: CompareOp,
        left: Box<Expr<S>>,
        right: Box<Expr<S>>,
    },
    FunctionCall {
        name: String,
        original_name: String,
        args: Vec<Expr<S>>,
    },
    /// Call a runtime lambda value produced by another expression.
    ///
    /// This models Excel's LAMBDA invocation syntax, e.g. `LAMBDA(x, x + 1)(3)`.
    Call {
        callee: Box<Expr<S>>,
        args: Vec<Expr<S>>,
    },
    /// Excel's implicit intersection operator (`@`).
    ImplicitIntersection(Box<Expr<S>>),
    /// Excel spill-range reference operator (`#`), e.g. `A1#`.
    SpillRange(Box<Expr<S>>),
}

impl<S: Clone> Expr<S> {
    pub fn map_sheets<T: Clone, F>(&self, f: &mut F) -> Expr<T>
    where
        F: FnMut(&SheetReference<S>) -> SheetReference<T>,
    {
        match self {
            Expr::Number(n) => Expr::Number(*n),
            Expr::Text(s) => Expr::Text(s.clone()),
            Expr::Bool(b) => Expr::Bool(*b),
            Expr::Blank => Expr::Blank,
            Expr::Error(e) => Expr::Error(*e),
            Expr::ArrayLiteral { rows, cols, values } => Expr::ArrayLiteral {
                rows: *rows,
                cols: *cols,
                values: {
                    let mut out: Vec<Expr<T>> = Vec::new();
                    if out.try_reserve_exact(values.len()).is_err() {
                        debug_assert!(
                            false,
                            "allocation failed (map_sheets array literal, len={})",
                            values.len()
                        );
                        return Expr::Error(ErrorKind::Num);
                    }
                    for v in values.iter() {
                        out.push(v.map_sheets(f));
                    }
                    out.into()
                },
            },
            Expr::NameRef(r) => Expr::NameRef(NameRef {
                sheet: f(&r.sheet),
                name: r.name.clone(),
            }),
            Expr::CellRef(r) => Expr::CellRef(CellRef {
                sheet: f(&r.sheet),
                addr: r.addr,
            }),
            Expr::RangeRef(r) => Expr::RangeRef(RangeRef {
                sheet: f(&r.sheet),
                start: r.start,
                end: r.end,
            }),
            Expr::StructuredRef(r) => Expr::StructuredRef(StructuredRefExpr {
                sheet: f(&r.sheet),
                sref: r.sref.clone(),
            }),
            Expr::FieldAccess { base, field } => Expr::FieldAccess {
                base: Box::new(base.map_sheets(f)),
                field: field.clone(),
            },
            Expr::Unary { op, expr } => Expr::Unary {
                op: *op,
                expr: Box::new(expr.map_sheets(f)),
            },
            Expr::Postfix { op, expr } => Expr::Postfix {
                op: *op,
                expr: Box::new(expr.map_sheets(f)),
            },
            Expr::Binary { op, left, right } => Expr::Binary {
                op: *op,
                left: Box::new(left.map_sheets(f)),
                right: Box::new(right.map_sheets(f)),
            },
            Expr::Compare { op, left, right } => Expr::Compare {
                op: *op,
                left: Box::new(left.map_sheets(f)),
                right: Box::new(right.map_sheets(f)),
            },
            Expr::FunctionCall {
                name,
                original_name,
                args,
            } => Expr::FunctionCall {
                name: name.clone(),
                original_name: original_name.clone(),
                args: {
                    let mut out: Vec<Expr<T>> = Vec::new();
                    if out.try_reserve_exact(args.len()).is_err() {
                        debug_assert!(
                            false,
                            "allocation failed (map_sheets function call args, len={})",
                            args.len()
                        );
                        return Expr::Error(ErrorKind::Num);
                    }
                    for a in args.iter() {
                        out.push(a.map_sheets(f));
                    }
                    out
                },
            },
            Expr::Call { callee, args } => Expr::Call {
                callee: Box::new(callee.map_sheets(f)),
                args: {
                    let mut out: Vec<Expr<T>> = Vec::new();
                    if out.try_reserve_exact(args.len()).is_err() {
                        debug_assert!(
                            false,
                            "allocation failed (map_sheets call args, len={})",
                            args.len()
                        );
                        return Expr::Error(ErrorKind::Num);
                    }
                    for a in args.iter() {
                        out.push(a.map_sheets(f));
                    }
                    out
                },
            },
            Expr::ImplicitIntersection(inner) => {
                Expr::ImplicitIntersection(Box::new(inner.map_sheets(f)))
            }
            Expr::SpillRange(inner) => Expr::SpillRange(Box::new(inner.map_sheets(f))),
        }
    }
}
