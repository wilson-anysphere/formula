use super::ast::Function;
use super::value::{RangeRef, Ref, Value};
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
        match ((self.0 >> 56) & 0xFF) as u8 {
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
            _ => unreachable!("invalid opcode"),
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
    pub(crate) funcs: Vec<Function>,
    pub(crate) key: Arc<str>,
}

impl Program {
    pub fn new(key: Arc<str>) -> Self {
        Self {
            instrs: Vec::new(),
            consts: Vec::new(),
            cell_refs: Vec::new(),
            range_refs: Vec::new(),
            funcs: Vec::new(),
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
