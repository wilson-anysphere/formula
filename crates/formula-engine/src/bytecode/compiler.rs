use super::ast::{BinaryOp, Expr, Function, UnaryOp};
use super::program::{ConstValue, Instruction, OpCode, Program};
use super::value::{RangeRef, Value};
use std::sync::Arc;

#[derive(Default)]
pub struct Compiler;

impl Compiler {
    pub fn compile(key: Arc<str>, expr: &Expr) -> Program {
        let mut program = Program::new(key);
        Self::compile_expr(expr, &mut program);
        program
    }

    fn compile_expr(expr: &Expr, program: &mut Program) {
        match expr {
            Expr::Literal(v) => {
                let idx = program.consts.len() as u32;
                program.consts.push(ConstValue::Value(v.clone()));
                program
                    .instrs
                    .push(Instruction::new(OpCode::PushConst, idx, 0));
            }
            Expr::CellRef(r) => {
                let idx = program.cell_refs.len() as u32;
                program.cell_refs.push(*r);
                program
                    .instrs
                    .push(Instruction::new(OpCode::LoadCell, idx, 0));
            }
            Expr::RangeRef(r) => {
                let idx = program.range_refs.len() as u32;
                program.range_refs.push(*r);
                program
                    .instrs
                    .push(Instruction::new(OpCode::LoadRange, idx, 0));
            }
            Expr::Unary { op, expr } => {
                Self::compile_expr(expr, program);
                match op {
                    UnaryOp::Plus => program
                        .instrs
                        .push(Instruction::new(OpCode::UnaryPlus, 0, 0)),
                    UnaryOp::Neg => program
                        .instrs
                        .push(Instruction::new(OpCode::UnaryNeg, 0, 0)),
                    UnaryOp::ImplicitIntersection => program
                        .instrs
                        .push(Instruction::new(OpCode::ImplicitIntersection, 0, 0)),
                }
            }
            Expr::Binary { op, left, right } => {
                Self::compile_expr(left, program);
                Self::compile_expr(right, program);
                let opcode = match op {
                    BinaryOp::Add => OpCode::Add,
                    BinaryOp::Sub => OpCode::Sub,
                    BinaryOp::Mul => OpCode::Mul,
                    BinaryOp::Div => OpCode::Div,
                    BinaryOp::Pow => OpCode::Pow,
                    BinaryOp::Eq => OpCode::Eq,
                    BinaryOp::Ne => OpCode::Ne,
                    BinaryOp::Lt => OpCode::Lt,
                    BinaryOp::Le => OpCode::Le,
                    BinaryOp::Gt => OpCode::Gt,
                    BinaryOp::Ge => OpCode::Ge,
                };
                program.instrs.push(Instruction::new(opcode, 0, 0));
            }
            Expr::FuncCall { func, args } => {
                for (arg_idx, arg) in args.iter().enumerate() {
                    Self::compile_func_arg(func, arg_idx, arg, program);
                }
                let func_idx = intern_func(program, func.clone());
                let argc = args.len() as u32;
                program
                    .instrs
                    .push(Instruction::new(OpCode::CallFunc, func_idx, argc));
            }
        }
    }

    fn compile_func_arg(func: &Function, arg_idx: usize, arg: &Expr, program: &mut Program) {
        // Excel-style aggregate functions have a quirk: a cell reference passed directly as an
        // argument (e.g. `SUM(A1)`) is treated as a *reference* argument, not a scalar, which means
        // logical/text values in the referenced cell are ignored (unlike `SUM(TRUE)` / `SUM("5")`).
        //
        // Lower single-cell references to a range so the runtime can apply reference semantics.
        let treat_cell_as_range = match func {
            Function::Sum | Function::Average | Function::Min | Function::Max | Function::Count => {
                true
            }
            Function::CountIf => arg_idx == 0,
            Function::SumProduct => true,
            Function::VLookup | Function::HLookup | Function::Match => arg_idx == 1,
            Function::Abs
            | Function::Int
            | Function::Round
            | Function::RoundUp
            | Function::RoundDown
            | Function::Mod
            | Function::Sign
            | Function::Concat
            | Function::Not
            | Function::Unknown(_) => false,
        };

        if treat_cell_as_range {
            if let Expr::CellRef(r) = arg {
                let idx = program.range_refs.len() as u32;
                program.range_refs.push(RangeRef::new(*r, *r));
                program
                    .instrs
                    .push(Instruction::new(OpCode::LoadRange, idx, 0));
                return;
            }
        }

        Self::compile_expr(arg, program);
    }
}

fn intern_func(program: &mut Program, func: Function) -> u32 {
    // We allow duplicates: the Vm uses the per-instruction index for dispatch without string compare.
    // If desired, this can be made a small map; for now keep it simple.
    let idx = program.funcs.len() as u32;
    program.funcs.push(func);
    idx
}

/// Canonical key for cache lookups, derived from the normalized AST.
pub fn normalized_key(expr: &Expr) -> Arc<str> {
    let mut s = String::new();
    expr_to_key(expr, &mut s);
    Arc::from(s)
}

fn expr_to_key(expr: &Expr, out: &mut String) {
    match expr {
        Expr::Literal(v) => match v {
            Value::Number(n) => {
                out.push_str("N");
                out.push_str(&format!("{:016x}", n.to_bits()));
            }
            Value::Bool(b) => {
                out.push_str(if *b { "B1" } else { "B0" });
            }
            Value::Text(t) => {
                out.push_str("S");
                out.push_str(t);
                out.push('\0');
            }
            Value::Empty => out.push_str("E"),
            Value::Error(e) => {
                out.push_str("ERR");
                out.push_str(&format!("{:?}", e));
                out.push('\0');
            }
            Value::Range(r) => {
                out.push_str("RANGE(");
                ref_to_key(r.start, out);
                out.push(',');
                ref_to_key(r.end, out);
                out.push(')');
            }
            Value::Array(a) => {
                out.push_str("ARR(");
                out.push_str(&a.rows.to_string());
                out.push('x');
                out.push_str(&a.cols.to_string());
                out.push(':');
                for v in &a.values {
                    out.push_str(&format!("{:016x}", v.to_bits()));
                    out.push(',');
                }
                out.push(')');
            }
        },
        Expr::CellRef(r) => {
            out.push_str("CELL(");
            ref_to_key(*r, out);
            out.push(')');
        }
        Expr::RangeRef(r) => {
            out.push_str("RANGE(");
            ref_to_key(r.start, out);
            out.push(',');
            ref_to_key(r.end, out);
            out.push(')');
        }
        Expr::Unary { op, expr } => {
            out.push_str("UN(");
            out.push_str(match op {
                UnaryOp::Plus => "+",
                UnaryOp::Neg => "-",
                UnaryOp::ImplicitIntersection => "@",
            });
            out.push(',');
            expr_to_key(expr, out);
            out.push(')');
        }
        Expr::Binary { op, left, right } => {
            out.push_str("BIN(");
            out.push_str(match op {
                BinaryOp::Add => "+",
                BinaryOp::Sub => "-",
                BinaryOp::Mul => "*",
                BinaryOp::Div => "/",
                BinaryOp::Pow => "^",
                BinaryOp::Eq => "=",
                BinaryOp::Ne => "<>",
                BinaryOp::Lt => "<",
                BinaryOp::Le => "<=",
                BinaryOp::Gt => ">",
                BinaryOp::Ge => ">=",
            });
            out.push(',');
            expr_to_key(left, out);
            out.push(',');
            expr_to_key(right, out);
            out.push(')');
        }
        Expr::FuncCall { func, args } => {
            out.push_str("FN(");
            out.push_str(func.name());
            out.push(',');
            out.push_str(&args.len().to_string());
            out.push(':');
            for arg in args {
                expr_to_key(arg, out);
                out.push(',');
            }
            out.push(')');
        }
    }
}

fn ref_to_key(r: super::value::Ref, out: &mut String) {
    out.push_str(if r.row_abs { "R$" } else { "R" });
    out.push_str(&r.row.to_string());
    out.push_str(if r.col_abs { "C$" } else { "C" });
    out.push_str(&r.col.to_string());
}
