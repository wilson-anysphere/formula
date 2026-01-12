use super::ast::{BinaryOp, UnaryOp};
use super::grid::Grid;
use super::runtime::{apply_binary, apply_unary, call_function};
use super::value::{CellCoord, Value};
use crate::locale::ValueLocaleConfig;

use super::program::{OpCode, Program};

/// Stack-based bytecode interpreter.
#[derive(Default)]
pub struct Vm {
    stack: Vec<Value>,
}

impl Vm {
    pub fn new() -> Self {
        Self { stack: Vec::new() }
    }

    pub fn with_capacity(stack: usize) -> Self {
        Self {
            stack: Vec::with_capacity(stack),
        }
    }

    pub fn eval(
        &mut self,
        program: &Program,
        grid: &dyn Grid,
        base: CellCoord,
    ) -> Value {
        self.stack.clear();
        for inst in program.instrs() {
            match inst.op() {
                OpCode::PushConst => {
                    let v = program.consts[inst.a() as usize].to_value();
                    self.stack.push(v);
                }
                OpCode::LoadCell => {
                    let r = program.cell_refs[inst.a() as usize];
                    self.stack.push(grid.get_value(r.resolve(base)));
                }
                OpCode::LoadRange => {
                    let r = program.range_refs[inst.a() as usize];
                    self.stack.push(Value::Range(r));
                }
                OpCode::UnaryPlus => {
                    let v = self.stack.pop().unwrap_or(Value::Empty);
                    self.stack.push(apply_unary(UnaryOp::Plus, v));
                }
                OpCode::UnaryNeg => {
                    let v = self.stack.pop().unwrap_or(Value::Empty);
                    self.stack.push(apply_unary(UnaryOp::Neg, v));
                }
                OpCode::Add
                | OpCode::Sub
                | OpCode::Mul
                | OpCode::Div
                | OpCode::Pow
                | OpCode::Eq
                | OpCode::Ne
                | OpCode::Lt
                | OpCode::Le
                | OpCode::Gt
                | OpCode::Ge => {
                    let right = self.stack.pop().unwrap_or(Value::Empty);
                    let left = self.stack.pop().unwrap_or(Value::Empty);
                    let op = match inst.op() {
                        OpCode::Add => BinaryOp::Add,
                        OpCode::Sub => BinaryOp::Sub,
                        OpCode::Mul => BinaryOp::Mul,
                        OpCode::Div => BinaryOp::Div,
                        OpCode::Pow => BinaryOp::Pow,
                        OpCode::Eq => BinaryOp::Eq,
                        OpCode::Ne => BinaryOp::Ne,
                        OpCode::Lt => BinaryOp::Lt,
                        OpCode::Le => BinaryOp::Le,
                        OpCode::Gt => BinaryOp::Gt,
                        OpCode::Ge => BinaryOp::Ge,
                        _ => unreachable!(),
                    };
                    self.stack.push(apply_binary(op, left, right));
                }
                OpCode::CallFunc => {
                    let func = &program.funcs[inst.a() as usize];
                    let argc = inst.b() as usize;
                    let start = self.stack.len().saturating_sub(argc);
                    let result = call_function(func, &self.stack[start..], grid, base);
                    self.stack.truncate(start);
                    self.stack.push(result);
                }
            }
        }
        self.stack.pop().unwrap_or(Value::Empty)
    }

    pub fn eval_with_value_locale(
        &mut self,
        program: &Program,
        grid: &dyn Grid,
        base: CellCoord,
        _value_locale: ValueLocaleConfig,
    ) -> Value {
        self.eval(program, grid, base)
    }
}
