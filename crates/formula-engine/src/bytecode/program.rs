use super::ast::Function;
use super::value::{MultiRangeRef, RangeRef, Ref, Value};
use std::sync::Arc;

#[repr(u8)]
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum OpCode {
    PushConst = 0,
    LoadCell = 1,
    LoadRange = 2,
    UnaryPlus = 3,
    UnaryNeg = 4,
    Add = 5,
    Sub = 6,
    Mul = 7,
    Div = 8,
    Pow = 9,
    Eq = 10,
    Ne = 11,
    Lt = 12,
    Le = 13,
    Gt = 14,
    Ge = 15,
    CallFunc = 16,
    ImplicitIntersection = 17,
    StoreLocal = 18,
    LoadLocal = 19,
    /// Unconditional jump to instruction index `a`.
    Jump = 20,
    /// Pop a value, coerce it to bool, and if FALSE jump to instruction index `a`.
    ///
    /// If coercion fails with an error, pushes the error value and jumps to instruction index `b`.
    JumpIfFalseOrError = 21,
    /// Inspect the top of the stack: if it's NOT an error, jump to instruction index `a`.
    /// Otherwise (it is an error), pop it and continue.
    JumpIfNotError = 22,
    /// Inspect the top of the stack: if it's NOT `#N/A`, jump to instruction index `a`.
    /// Otherwise (it is `#N/A`), pop it and continue.
    JumpIfNotNaError = 23,
    LoadMultiRange = 24,
    SpillRange = 25,
    /// Reference union (`,`).
    Union = 26,
    /// Reference intersection (whitespace).
    Intersect = 27,
    /// Create a closure capturing the current lexical environment.
    MakeLambda = 28,
    /// Call a closure value with `b` arguments.
    CallValue = 29,
    /// Sentinel opcode used when decoding invalid bytecode.
    ///
    /// This should never be produced by the compiler, but can appear if the program is corrupted.
    Invalid = 255,
}

/// Packed instruction:
/// - bits 56..63: opcode
/// - bits 28..55: operand a
/// - bits 0..27: operand b
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Instruction(u64);

impl Instruction {
    #[inline]
    pub fn new(op: OpCode, a: u32, b: u32) -> Self {
        debug_assert!(a < (1 << 28));
        debug_assert!(b < (1 << 28));
        Instruction(((op as u64) << 56) | ((a as u64) << 28) | (b as u64))
    }

    #[inline]
    pub fn op(self) -> OpCode {
        let op = ((self.0 >> 56) & 0xFF) as u8;
        match op {
            0 => OpCode::PushConst,
            1 => OpCode::LoadCell,
            2 => OpCode::LoadRange,
            3 => OpCode::UnaryPlus,
            4 => OpCode::UnaryNeg,
            5 => OpCode::Add,
            6 => OpCode::Sub,
            7 => OpCode::Mul,
            8 => OpCode::Div,
            9 => OpCode::Pow,
            10 => OpCode::Eq,
            11 => OpCode::Ne,
            12 => OpCode::Lt,
            13 => OpCode::Le,
            14 => OpCode::Gt,
            15 => OpCode::Ge,
            16 => OpCode::CallFunc,
            17 => OpCode::ImplicitIntersection,
            18 => OpCode::StoreLocal,
            19 => OpCode::LoadLocal,
            20 => OpCode::Jump,
            21 => OpCode::JumpIfFalseOrError,
            22 => OpCode::JumpIfNotError,
            23 => OpCode::JumpIfNotNaError,
            24 => OpCode::LoadMultiRange,
            25 => OpCode::SpillRange,
            26 => OpCode::Union,
            27 => OpCode::Intersect,
            28 => OpCode::MakeLambda,
            29 => OpCode::CallValue,
            _ => {
                debug_assert!(false, "invalid opcode: {op}");
                OpCode::Invalid
            }
        }
    }

    #[inline]
    pub fn a(self) -> u32 {
        ((self.0 >> 28) & 0x0FFF_FFFF) as u32
    }

    #[inline]
    pub fn b(self) -> u32 {
        (self.0 & 0x0FFF_FFFF) as u32
    }
}

#[derive(Clone, Debug)]
pub enum ConstValue {
    Value(Value),
}

impl ConstValue {
    #[inline]
    pub fn to_value(&self) -> Value {
        match self {
            ConstValue::Value(v) => v.clone(),
        }
    }
}

#[derive(Clone, Debug)]
pub struct Program {
    pub(crate) instrs: Vec<Instruction>,
    pub(crate) consts: Vec<ConstValue>,
    pub(crate) cell_refs: Vec<Ref>,
    pub(crate) range_refs: Vec<RangeRef>,
    pub(crate) multi_range_refs: Vec<MultiRangeRef>,
    pub(crate) funcs: Vec<Function>,
    pub(crate) locals: Vec<Arc<str>>,
    pub(crate) lambdas: Vec<Arc<LambdaTemplate>>,
    pub(crate) key: Arc<str>,
}

#[derive(Clone, Debug)]
pub struct LambdaTemplate {
    pub params: Arc<[Arc<str>]>,
    pub body: Arc<Program>,
    pub param_locals: Arc<[u32]>,
    /// Per-parameter `ISOMITTED` flags aligned with `params` / `param_locals`.
    pub omitted_param_locals: Arc<[u32]>,
    pub captures: Arc<[Capture]>,
    pub self_local: Option<u32>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Capture {
    pub outer_local: u32,
    pub inner_local: u32,
}

impl Program {
    pub fn new(key: Arc<str>) -> Self {
        Self {
            instrs: Vec::new(),
            consts: Vec::new(),
            cell_refs: Vec::new(),
            range_refs: Vec::new(),
            multi_range_refs: Vec::new(),
            funcs: Vec::new(),
            locals: Vec::new(),
            lambdas: Vec::new(),
            key,
        }
    }

    #[inline]
    pub fn key(&self) -> &str {
        &self.key
    }

    #[inline]
    pub fn instrs(&self) -> &[Instruction] {
        &self.instrs
    }
}
