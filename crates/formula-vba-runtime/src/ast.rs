use std::collections::HashMap;

use crate::value::VbaValue;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VbaType {
    Variant,
    Integer,
    Long,
    Double,
    String,
    Boolean,
    Date,
}

#[derive(Debug, Clone)]
pub struct ArrayDim {
    pub lower: Option<Expr>,
    pub upper: Expr,
}

#[derive(Debug, Clone)]
pub struct VarDecl {
    pub name: String,
    pub ty: VbaType,
    pub dims: Vec<ArrayDim>,
}

#[derive(Debug, Clone)]
pub struct ConstDecl {
    pub name: String,
    pub ty: Option<VbaType>,
    pub value: Expr,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ProcedureKind {
    Sub,
    Function,
}

#[derive(Debug, Clone)]
pub struct VbaProgram {
    pub option_explicit: bool,
    pub module_vars: Vec<VarDecl>,
    pub module_consts: Vec<ConstDecl>,
    pub procedures: HashMap<String, ProcedureDef>,
}

impl VbaProgram {
    pub fn new() -> Self {
        Self {
            option_explicit: false,
            module_vars: Vec::new(),
            module_consts: Vec::new(),
            procedures: HashMap::new(),
        }
    }

    pub fn get(&self, name: &str) -> Option<&ProcedureDef> {
        self.procedures.get(&name.to_ascii_lowercase())
    }
}

impl Default for VbaProgram {
    fn default() -> Self {
        Self::new()
    }
}

#[derive(Debug, Clone)]
pub struct ProcedureDef {
    pub name: String,
    pub kind: ProcedureKind,
    pub params: Vec<ParamDef>,
    pub return_type: Option<VbaType>,
    pub body: Vec<Stmt>,
}

#[derive(Debug, Clone)]
pub struct ParamDef {
    pub name: String,
    pub by_ref: bool,
    pub ty: Option<VbaType>,
}

#[derive(Debug, Clone)]
pub enum Stmt {
    Dim(Vec<VarDecl>),
    Const(Vec<ConstDecl>),
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
    ForEach {
        var: String,
        iterable: Expr,
        body: Vec<Stmt>,
    },
    DoLoop {
        pre_condition: Option<(LoopConditionKind, Expr)>,
        post_condition: Option<(LoopConditionKind, Expr)>,
        body: Vec<Stmt>,
    },
    While {
        cond: Expr,
        body: Vec<Stmt>,
    },
    SelectCase {
        expr: Expr,
        cases: Vec<SelectCaseArm>,
        else_body: Vec<Stmt>,
    },
    With {
        object: Expr,
        body: Vec<Stmt>,
    },
    ExitSub,
    ExitFunction,
    ExitFor,
    ExitDo,
    OnErrorResumeNext,
    OnErrorGoto0,
    OnErrorGotoLabel(String),
    ResumeNext,
    Resume,
    ResumeLabel(String),
    Label(String),
    Goto(String),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LoopConditionKind {
    While,
    Until,
}

#[derive(Debug, Clone)]
pub struct SelectCaseArm {
    pub conditions: Vec<CaseCondition>,
    pub body: Vec<Stmt>,
}

#[derive(Debug, Clone)]
pub enum CaseCondition {
    Expr(Expr),
    Range { start: Expr, end: Expr },
    Is { op: CaseComparisonOp, expr: Expr },
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CaseComparisonOp {
    Eq,
    Ne,
    Lt,
    Le,
    Gt,
    Ge,
}

#[derive(Debug, Clone)]
pub enum Expr {
    Literal(VbaValue),
    Missing,
    Var(String),
    With,
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
        args: Vec<CallArg>,
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

#[derive(Debug, Clone)]
pub struct CallArg {
    pub name: Option<String>,
    pub expr: Expr,
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
    IntDiv,
    Mod,
    Pow,
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
