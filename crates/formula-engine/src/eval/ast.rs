use crate::eval::address::CellAddr;
use crate::value::ErrorKind;

pub type ParsedExpr = Expr<String>;
pub type CompiledExpr = Expr<usize>;

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum SheetReference<S> {
    Current,
    Sheet(S),
    /// External workbook reference like `[Book.xlsx]Sheet1!A1`.
    /// Not implemented yet; evaluating it yields `#REF!`.
    External(String),
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct CellRef<S> {
    pub sheet: SheetReference<S>,
    pub addr: CellAddr,
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct RangeRef<S> {
    pub sheet: SheetReference<S>,
    pub start: CellAddr,
    pub end: CellAddr,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum UnaryOp {
    Plus,
    Minus,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum BinaryOp {
    Add,
    Sub,
    Mul,
    Div,
    Pow,
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
    CellRef(CellRef<S>),
    RangeRef(RangeRef<S>),
    StructuredRef(crate::structured_refs::StructuredRef),
    Unary {
        op: UnaryOp,
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
    /// Excel's implicit intersection operator (`@`).
    ImplicitIntersection(Box<Expr<S>>),
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
            Expr::CellRef(r) => Expr::CellRef(CellRef {
                sheet: f(&r.sheet),
                addr: r.addr,
            }),
            Expr::RangeRef(r) => Expr::RangeRef(RangeRef {
                sheet: f(&r.sheet),
                start: r.start,
                end: r.end,
            }),
            Expr::StructuredRef(r) => Expr::StructuredRef(r.clone()),
            Expr::Unary { op, expr } => Expr::Unary {
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
                args: args.iter().map(|a| a.map_sheets(f)).collect(),
            },
            Expr::ImplicitIntersection(inner) => {
                Expr::ImplicitIntersection(Box::new(inner.map_sheets(f)))
            }
        }
    }
}
