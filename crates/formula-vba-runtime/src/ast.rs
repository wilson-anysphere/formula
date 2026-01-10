use std::collections::HashMap;

use crate::value::VbaValue;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ProcedureKind {
    Sub,
    Function,
}

#[derive(Debug, Clone)]
pub struct VbaProgram {
    pub procedures: HashMap<String, ProcedureDef>,
}

impl VbaProgram {
    pub fn new() -> Self {
        Self {
            procedures: HashMap::new(),
        }
    }

    pub fn get(&self, name: &str) -> Option<&ProcedureDef> {
        self.procedures.get(&name.to_ascii_lowercase())
    }
}

#[derive(Debug, Clone)]
pub struct ProcedureDef {
    pub name: String,
    pub kind: ProcedureKind,
    pub params: Vec<ParamDef>,
    pub body: Vec<Stmt>,
}

#[derive(Debug, Clone)]
pub struct ParamDef {
    pub name: String,
    pub by_ref: bool,
}

#[derive(Debug, Clone)]
pub enum Stmt {
    Dim(Vec<String>),
    Assign {
        target: Expr,
        value: Expr,
    },
    Set {
        target: Expr,
        value: Expr,
    },
    ExprStmt(Expr),
    If {
        cond: Expr,
        then_body: Vec<Stmt>,
        elseifs: Vec<(Expr, Vec<Stmt>)>,
        else_body: Vec<Stmt>,
    },
    For {
        var: String,
        start: Expr,
        end: Expr,
        step: Option<Expr>,
        body: Vec<Stmt>,
    },
    DoWhile {
        cond: Expr,
        body: Vec<Stmt>,
    },
    ExitSub,
    ExitFunction,
    ExitFor,
    OnErrorResumeNext,
    OnErrorGoto0,
    OnErrorGotoLabel(String),
    Label(String),
    Goto(String),
}

#[derive(Debug, Clone)]
pub enum Expr {
    Literal(VbaValue),
    Var(String),
    Unary {
        op: UnOp,
        expr: Box<Expr>,
    },
    Binary {
        op: BinOp,
        left: Box<Expr>,
        right: Box<Expr>,
    },
    Call {
        callee: Box<Expr>,
        args: Vec<Expr>,
    },
    Member {
        object: Box<Expr>,
        member: String,
    },
    Index {
        array: Box<Expr>,
        indices: Vec<Expr>,
    },
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum UnOp {
    Neg,
    Not,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BinOp {
    Add,
    Sub,
    Mul,
    Div,
    Concat,
    Eq,
    Ne,
    Lt,
    Le,
    Gt,
    Ge,
    And,
    Or,
}
