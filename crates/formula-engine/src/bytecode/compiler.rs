use super::ast::{BinaryOp, Expr, Function, UnaryOp};
use super::program::{ConstValue, Instruction, OpCode, Program};
use super::value::{ErrorKind, RangeRef, SheetRangeRef, Value};
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
            Expr::MultiRangeRef(r) => {
                let idx = self.program.multi_range_refs.len() as u32;
                self.program.multi_range_refs.push(r.clone());
                self.program
                    .instrs
                    .push(Instruction::new(OpCode::LoadMultiRange, idx, 0));
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
            Expr::SpillRange(inner) => {
                self.compile_expr(inner);
                self.program
                    .instrs
                    .push(Instruction::new(OpCode::SpillRange, 0, 0));
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
                // Certain logical/error functions are lazy in Excel: they should only evaluate
                // the branch argument that is selected at runtime (e.g. `IF(FALSE, <expensive>, 0)`).
                //
                // Implement these with explicit control flow opcodes so the VM can short-circuit.
                Function::If if args.len() == 2 || args.len() == 3 => {
                    self.compile_expr(&args[0]);
                    let jump_idx = self.program.instrs.len();
                    // Patched below: a=false_target, b=end_target.
                    self.program.instrs.push(Instruction::new(
                        OpCode::JumpIfFalseOrError,
                        0,
                        0,
                    ));

                    // TRUE branch.
                    self.compile_expr(&args[1]);
                    let jump_end_idx = self.program.instrs.len();
                    self.program
                        .instrs
                        .push(Instruction::new(OpCode::Jump, 0, 0));

                    // FALSE branch.
                    let false_target = self.program.instrs.len() as u32;
                    if args.len() == 3 {
                        self.compile_expr(&args[2]);
                    } else {
                        // Engine behavior: missing false branch defaults to FALSE (not blank).
                        let idx = self.program.consts.len() as u32;
                        self.program
                            .consts
                            .push(ConstValue::Value(Value::Bool(false)));
                        self.program
                            .instrs
                            .push(Instruction::new(OpCode::PushConst, idx, 0));
                    }

                    let end_target = self.program.instrs.len() as u32;
                    self.program.instrs[jump_idx] = Instruction::new(
                        OpCode::JumpIfFalseOrError,
                        false_target,
                        end_target,
                    );
                    self.program.instrs[jump_end_idx] =
                        Instruction::new(OpCode::Jump, end_target, 0);
                }
                Function::Ifs => self.compile_ifs(args),
                Function::IfError if args.len() == 2 => {
                    // Evaluate the first argument. If it's not an error, short-circuit and
                    // return it without evaluating the fallback.
                    self.compile_expr(&args[0]);
                    let jump_idx = self.program.instrs.len();
                    self.program
                        .instrs
                        .push(Instruction::new(OpCode::JumpIfNotError, 0, 0));

                    // Error fallback branch (only evaluated if arg0 is an error).
                    self.compile_expr(&args[1]);
                    let end_target = self.program.instrs.len() as u32;
                    self.program.instrs[jump_idx] =
                        Instruction::new(OpCode::JumpIfNotError, end_target, 0);
                }
                Function::IfNa if args.len() == 2 => {
                    // Evaluate the first argument. If it's not #N/A, short-circuit and return
                    // it without evaluating the fallback.
                    self.compile_expr(&args[0]);
                    let jump_idx = self.program.instrs.len();
                    self.program.instrs.push(Instruction::new(
                        OpCode::JumpIfNotNaError,
                        0,
                        0,
                    ));

                    // #N/A fallback branch (only evaluated if arg0 is #N/A).
                    self.compile_expr(&args[1]);
                    let end_target = self.program.instrs.len() as u32;
                    self.program.instrs[jump_idx] =
                        Instruction::new(OpCode::JumpIfNotNaError, end_target, 0);
                }
                Function::Switch => self.compile_switch(args),
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

    fn alloc_temp_local(&mut self, label: &'static str) -> u32 {
        self.alloc_local(Arc::from(label))
    }

    fn compile_ifs(&mut self, args: &[Expr]) {
        // IFS requires at least one condition/value pair, and the argument count must be even.
        if args.len() < 2 {
            self.push_error_const(ErrorKind::Value);
            return;
        }
        if args.len() % 2 != 0 {
            self.push_error_const(ErrorKind::Value);
            return;
        }

        let mut jump_if_idxs: Vec<usize> = Vec::new();
        let mut jump_end_idxs: Vec<usize> = Vec::new();

        for pair in args.chunks_exact(2) {
            self.compile_expr(&pair[0]);
            let jump_idx = self.program.instrs.len();
            // Patched below: a=next_case_target, b=end_target (for error propagation).
            self.program
                .instrs
                .push(Instruction::new(OpCode::JumpIfFalseOrError, 0, 0));

            // TRUE branch.
            self.compile_expr(&pair[1]);
            let jump_end_idx = self.program.instrs.len();
            self.program
                .instrs
                .push(Instruction::new(OpCode::Jump, 0, 0));

            // Next condition/value pair (or the final #N/A block).
            let next_case_target = self.program.instrs.len() as u32;
            self.program.instrs[jump_idx] =
                Instruction::new(OpCode::JumpIfFalseOrError, next_case_target, 0);

            jump_if_idxs.push(jump_idx);
            jump_end_idxs.push(jump_end_idx);
        }

        // If no condition is true, IFS returns #N/A.
        self.push_error_const(ErrorKind::NA);
        let end_target = self.program.instrs.len() as u32;

        for idx in jump_if_idxs {
            let a = self.program.instrs[idx].a();
            self.program.instrs[idx] = Instruction::new(OpCode::JumpIfFalseOrError, a, end_target);
        }
        for idx in jump_end_idxs {
            self.program.instrs[idx] = Instruction::new(OpCode::Jump, end_target, 0);
        }
    }

    fn compile_switch(&mut self, args: &[Expr]) {
        // SWITCH(expr, value1, result1, [value2, result2], ..., [default])
        if args.len() < 3 {
            self.push_error_const(ErrorKind::Value);
            return;
        }

        let has_default = (args.len() - 1) % 2 != 0;
        let pairs_end = if has_default { args.len() - 1 } else { args.len() };
        let pairs = &args[1..pairs_end];
        let default = if has_default { Some(&args[args.len() - 1]) } else { None };

        if pairs.len() < 2 || pairs.len() % 2 != 0 {
            self.push_error_const(ErrorKind::Value);
            return;
        }

        // Evaluate the discriminant expression once and store it in a temp local.
        self.compile_expr(&args[0]);
        let expr_local = self.alloc_temp_local("\u{0}SWITCH_EXPR");
        self.program
            .instrs
            .push(Instruction::new(OpCode::StoreLocal, expr_local, 0));

        // If the discriminant is an error, return it without evaluating any case expressions.
        self.program
            .instrs
            .push(Instruction::new(OpCode::LoadLocal, expr_local, 0));
        let jump_not_error_idx = self.program.instrs.len();
        self.program
            .instrs
            .push(Instruction::new(OpCode::JumpIfNotError, 0, 0));
        // Error path: reload the error and jump to end.
        self.program
            .instrs
            .push(Instruction::new(OpCode::LoadLocal, expr_local, 0));
        let jump_error_end_idx = self.program.instrs.len();
        self.program
            .instrs
            .push(Instruction::new(OpCode::Jump, 0, 0));

        let cases_start = self.program.instrs.len() as u32;
        self.program.instrs[jump_not_error_idx] =
            Instruction::new(OpCode::JumpIfNotError, cases_start, 0);

        let mut jump_if_idxs: Vec<usize> = Vec::new();
        let mut jump_end_idxs: Vec<usize> = vec![jump_error_end_idx];

        for (idx, pair) in pairs.chunks_exact(2).enumerate() {
            if idx != 0 {
                self.program
                    .instrs
                    .push(Instruction::new(OpCode::LoadLocal, expr_local, 0));
            }

            // Compare discriminant == case value.
            self.compile_expr(&pair[0]);
            self.program
                .instrs
                .push(Instruction::new(OpCode::Eq, 0, 0));

            let jump_idx = self.program.instrs.len();
            // Patched below: a=next_case_target, b=end_target (for error propagation).
            self.program
                .instrs
                .push(Instruction::new(OpCode::JumpIfFalseOrError, 0, 0));

            // Match branch: evaluate result and jump to end.
            self.compile_expr(&pair[1]);
            let jump_end_idx = self.program.instrs.len();
            self.program
                .instrs
                .push(Instruction::new(OpCode::Jump, 0, 0));

            let next_case_target = self.program.instrs.len() as u32;
            self.program.instrs[jump_idx] =
                Instruction::new(OpCode::JumpIfFalseOrError, next_case_target, 0);
            jump_if_idxs.push(jump_idx);
            jump_end_idxs.push(jump_end_idx);
        }

        // No match: return default if provided, otherwise #N/A.
        match default {
            Some(expr) => self.compile_expr(expr),
            None => self.push_error_const(ErrorKind::NA),
        }
        let end_target = self.program.instrs.len() as u32;

        for idx in jump_if_idxs {
            let a = self.program.instrs[idx].a();
            self.program.instrs[idx] = Instruction::new(OpCode::JumpIfFalseOrError, a, end_target);
        }
        for idx in jump_end_idxs {
            self.program.instrs[idx] = Instruction::new(OpCode::Jump, end_target, 0);
        }
    }

    fn compile_func_arg(&mut self, func: &Function, arg_idx: usize, arg: &Expr) {
        // Excel-style aggregate functions have a quirk: a cell reference passed directly as an
        // argument (e.g. `SUM(A1)`) is treated as a *reference* argument, not a scalar, which means
        // logical/text values in the referenced cell are ignored (unlike `SUM(TRUE)` / `SUM("5")`).
        //
        // Lower single-cell references to a range so the runtime can apply reference semantics.
        let treat_cell_as_range = match func {
            // AND/OR are *not* treated like numeric aggregates: a direct cell reference argument is
            // treated as a scalar, so text-like values in the referenced cell behave like scalar
            // text arguments (i.e. typically #VALUE! rather than being ignored).
            Function::And | Function::Or => false,
            // XOR uses reference semantics for direct cell references (like ranges), matching the
            // engine evaluator.
            Function::Xor => true,
            Function::Sum
            | Function::Average
            | Function::Min
            | Function::Max
            | Function::Count
            | Function::CountA
            | Function::CountBlank => true,
            Function::CountIf => arg_idx == 0,
            Function::SumIf | Function::AverageIf => arg_idx == 0 || arg_idx == 2,
            Function::SumIfs | Function::AverageIfs | Function::MinIfs | Function::MaxIfs => {
                arg_idx == 0 || arg_idx % 2 == 1
            }
            Function::CountIfs => arg_idx % 2 == 0,
            Function::SumProduct => true,
            Function::VLookup | Function::HLookup | Function::Match => arg_idx == 1,
            Function::Row | Function::Column | Function::Rows | Function::Columns => true,
            _ => false,
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

fn value_to_key(v: &Value, out: &mut String) {
    match v {
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
        Value::MultiRange(r) => {
            out.push_str("MRANGE(");
            out.push_str(&r.areas.len().to_string());
            out.push(':');
            for area in r.areas.iter() {
                sheet_range_to_key(*area, out);
                out.push(',');
            }
            out.push(')');
        }
        Value::Array(a) => {
            out.push_str("ARR(");
            out.push_str(&a.rows.to_string());
            out.push('x');
            out.push_str(&a.cols.to_string());
            out.push(':');
            for v in &a.values {
                value_to_key(v, out);
                out.push(',');
            }
            out.push(')');
        }
        // Rich value variants (e.g. entity / record types) can be carried through the bytecode
        // runtime for pass-through semantics, but are not currently produced by parsing/lowering
        // as literals. Encode them as an opaque marker so cache keys stay deterministic without
        // needing to serialize the rich payload.
        #[allow(unreachable_patterns)]
        _ => out.push_str("RICH"),
    }
}

fn expr_to_key(expr: &Expr, out: &mut String) {
    match expr {
        Expr::Literal(v) => value_to_key(v, out),
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
        Expr::MultiRangeRef(r) => {
            out.push_str("MRANGE(");
            out.push_str(&r.areas.len().to_string());
            out.push(':');
            for area in r.areas.iter() {
                sheet_range_to_key(*area, out);
                out.push(',');
            }
            out.push(')');
        }
        Expr::SpillRange(inner) => {
            out.push_str("SPILL#(");
            expr_to_key(inner, out);
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

fn sheet_range_to_key(r: SheetRangeRef, out: &mut String) {
    out.push_str("S");
    out.push_str(&r.sheet.to_string());
    out.push(':');
    ref_to_key(r.range.start, out);
    out.push(',');
    ref_to_key(r.range.end, out);
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn normalized_key_is_stable_for_common_expressions() {
        let a1 = super::super::value::CellCoord::new(0, 0);
        let a2 = super::super::value::CellCoord::new(1, 0);

        // Same relative pattern filled down a row should hit the same cache key.
        let expr1 = super::super::ast::parse_formula("=B1+1", a1).expect("parse");
        let expr2 = super::super::ast::parse_formula("=B2+1", a2).expect("parse");
        assert_eq!(normalized_key(&expr1), normalized_key(&expr2));

        // Literal/text/function encoding should remain deterministic.
        let expr = super::super::ast::parse_formula("=CONCAT(\"a\",\"b\")", a1).expect("parse");
        assert_eq!(normalized_key(&expr).as_ref(), "FN(CONCAT,2:Sa\0,Sb\0,)");
    }
}
