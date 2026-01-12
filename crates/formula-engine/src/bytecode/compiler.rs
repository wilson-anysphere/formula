use super::ast::{BinaryOp, Expr, Function, UnaryOp};
use super::program::{ConstValue, Instruction, OpCode, Program};
use super::value::{ErrorKind, RangeRef, Value};
use std::collections::HashMap;
use std::sync::Arc;

#[derive(Default)]
pub struct Compiler;

impl Compiler {
    pub fn compile(key: Arc<str>, expr: &Expr) -> Program {
        let mut program = Program::new(key);
        let mut ctx = CompileCtx::new(&mut program);
        ctx.compile_expr(expr);
        program
    }
}

struct CompileCtx<'a> {
    program: &'a mut Program,
    lexical_scopes: Vec<HashMap<Arc<str>, u32>>,
}

impl<'a> CompileCtx<'a> {
    fn new(program: &'a mut Program) -> Self {
        Self {
            program,
            lexical_scopes: Vec::new(),
        }
    }

    fn resolve_local(&self, name: &Arc<str>) -> Option<u32> {
        for scope in self.lexical_scopes.iter().rev() {
            if let Some(idx) = scope.get(name) {
                return Some(*idx);
            }
        }
        None
    }

    fn alloc_local(&mut self, name: Arc<str>) -> u32 {
        let idx = self.program.locals.len() as u32;
        self.program.locals.push(name);
        idx
    }

    fn push_error_const(&mut self, kind: ErrorKind) {
        let idx = self.program.consts.len() as u32;
        self.program
            .consts
            .push(ConstValue::Value(Value::Error(kind)));
        self.program
            .instrs
            .push(Instruction::new(OpCode::PushConst, idx, 0));
    }

    fn compile_expr(&mut self, expr: &Expr) {
        match expr {
            Expr::Literal(v) => {
                let idx = self.program.consts.len() as u32;
                self.program.consts.push(ConstValue::Value(v.clone()));
                self.program
                    .instrs
                    .push(Instruction::new(OpCode::PushConst, idx, 0));
            }
            Expr::CellRef(r) => {
                let idx = self.program.cell_refs.len() as u32;
                self.program.cell_refs.push(*r);
                self.program
                    .instrs
                    .push(Instruction::new(OpCode::LoadCell, idx, 0));
            }
            Expr::RangeRef(r) => {
                let idx = self.program.range_refs.len() as u32;
                self.program.range_refs.push(*r);
                self.program
                    .instrs
                    .push(Instruction::new(OpCode::LoadRange, idx, 0));
            }
            Expr::NameRef(name) => {
                if let Some(idx) = self.resolve_local(name) {
                    self.program
                        .instrs
                        .push(Instruction::new(OpCode::LoadLocal, idx, 0));
                } else {
                    // Bytecode evaluation does not currently resolve defined names; treat as `#NAME?`.
                    self.push_error_const(ErrorKind::Name);
                }
            }
            Expr::Unary { op, expr } => {
                self.compile_expr(expr);
                match op {
                    UnaryOp::Plus => self
                        .program
                        .instrs
                        .push(Instruction::new(OpCode::UnaryPlus, 0, 0)),
                    UnaryOp::Neg => self
                        .program
                        .instrs
                        .push(Instruction::new(OpCode::UnaryNeg, 0, 0)),
                    UnaryOp::ImplicitIntersection => self
                        .program
                        .instrs
                        .push(Instruction::new(OpCode::ImplicitIntersection, 0, 0)),
                }
            }
            Expr::Binary { op, left, right } => {
                self.compile_expr(left);
                self.compile_expr(right);
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
                self.program.instrs.push(Instruction::new(opcode, 0, 0));
            }
            Expr::FuncCall { func, args } => match func {
                Function::Let => self.compile_let(args),
                _ => {
                    for (arg_idx, arg) in args.iter().enumerate() {
                        self.compile_func_arg(func, arg_idx, arg);
                    }
                    let func_idx = intern_func(self.program, func.clone());
                    let argc = args.len() as u32;
                    self.program
                        .instrs
                        .push(Instruction::new(OpCode::CallFunc, func_idx, argc));
                }
            },
        }
    }

    fn compile_let(&mut self, args: &[Expr]) {
        if args.len() < 3 || args.len() % 2 == 0 {
            self.push_error_const(ErrorKind::Value);
            return;
        }

        let last = args.len() - 1;
        self.lexical_scopes.push(HashMap::new());

        for pair in args[..last].chunks_exact(2) {
            let Expr::NameRef(name) = &pair[0] else {
                self.lexical_scopes.pop();
                self.push_error_const(ErrorKind::Value);
                return;
            };

            // Evaluate the binding value, then store it into a new local slot.
            // The binding name becomes visible to subsequent bindings and the final body.
            self.compile_expr(&pair[1]);
            let idx = self.alloc_local(name.clone());
            self.program
                .instrs
                .push(Instruction::new(OpCode::StoreLocal, idx, 0));
            self.lexical_scopes
                .last_mut()
                .expect("pushed scope")
                .insert(name.clone(), idx);
        }

        self.compile_expr(&args[last]);
        self.lexical_scopes.pop();
    }

    fn compile_func_arg(&mut self, func: &Function, arg_idx: usize, arg: &Expr) {
        // Excel-style aggregate functions have a quirk: a cell reference passed directly as an
        // argument (e.g. `SUM(A1)`) is treated as a *reference* argument, not a scalar, which means
        // logical/text values in the referenced cell are ignored (unlike `SUM(TRUE)` / `SUM("5")`).
        //
        // Lower single-cell references to a range so the runtime can apply reference semantics.
        let treat_cell_as_range = match func {
            // AND/OR have the same reference-vs-scalar semantics as the aggregate functions:
            // a direct cell reference argument is treated as a reference, not a scalar, so blanks
            // and text values in the referenced cell are ignored.
            Function::And | Function::Or => true,
            Function::Sum | Function::Average | Function::Min | Function::Max | Function::Count => {
                true
            }
            Function::CountIf => arg_idx == 0,
            Function::SumIf | Function::AverageIf => arg_idx == 0 || arg_idx == 2,
            Function::SumIfs | Function::AverageIfs | Function::MinIfs | Function::MaxIfs => {
                arg_idx == 0 || arg_idx % 2 == 1
            }
            Function::CountIfs => arg_idx % 2 == 0,
            Function::SumProduct => true,
            Function::VLookup | Function::HLookup | Function::Match => arg_idx == 1,
            Function::If
            | Function::IfError
            | Function::IfNa
            | Function::IsError
            | Function::IsNa
            | Function::Na
            | Function::Let
            | Function::Abs
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
                let idx = self.program.range_refs.len() as u32;
                self.program.range_refs.push(RangeRef::new(*r, *r));
                self.program
                    .instrs
                    .push(Instruction::new(OpCode::LoadRange, idx, 0));
                return;
            }
        }

        self.compile_expr(arg);
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
        Expr::NameRef(name) => {
            out.push_str("NAME(");
            out.push_str(name);
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
