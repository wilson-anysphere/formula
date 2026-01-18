use super::ast::{BinaryOp, Expr, Function, UnaryOp};
use super::program::{Capture, ConstValue, Instruction, LambdaTemplate, OpCode, Program};
use super::value::{ErrorKind, RangeRef, SheetId, SheetRangeRef, Value};
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

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum LocalKind {
    Scalar,
    RefSingle,
    RefMulti,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct LocalInfo {
    idx: u32,
    kind: LocalKind,
}

struct CompileCtx<'a> {
    program: &'a mut Program,
    lexical_scopes: Vec<HashMap<Arc<str>, LocalInfo>>,
    lambda_self_name: Option<Arc<str>>,
    omitted_param_locals: Option<HashMap<Arc<str>, u32>>,
    closure: Option<ClosureCtx>,
}

struct ClosureCtx {
    outer_locals: HashMap<Arc<str>, LocalInfo>,
    captures: Vec<Capture>,
}

impl<'a> CompileCtx<'a> {
    fn new(program: &'a mut Program) -> Self {
        Self {
            program,
            lexical_scopes: Vec::new(),
            lambda_self_name: None,
            omitted_param_locals: None,
            closure: None,
        }
    }

    fn resolve_local(&self, name: &Arc<str>) -> Option<LocalInfo> {
        for scope in self.lexical_scopes.iter().rev() {
            if let Some(info) = scope.get(name) {
                return Some(*info);
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
        self.compile_expr_inner(expr, false);
    }

    fn compile_expr_inner(&mut self, expr: &Expr, allow_range: bool) {
        match expr {
            Expr::Literal(v) => {
                // `Value::Missing` is used during lowering as a placeholder for syntactically blank
                // arguments (e.g. `ADDRESS(1,1,,FALSE)`), but it must not be allowed to propagate as
                // a general runtime value (e.g. via `IF(FALSE,1,)`).
                //
                // Treat literal `Missing` as a normal blank value during expression evaluation.
                // `compile_func_arg` special-cases direct blank arguments and preserves `Missing`
                // so functions that need to distinguish omitted args from blank cell values can
                // still do so.
                let v = match v {
                    Value::Missing => Value::Empty,
                    other => other.clone(),
                };
                let idx = self.program.consts.len() as u32;
                self.program.consts.push(ConstValue::Value(v));
                self.program
                    .instrs
                    .push(Instruction::new(OpCode::PushConst, idx, 0));
            }
            Expr::CellRef(r) => {
                if allow_range {
                    let idx = self.program.range_refs.len() as u32;
                    self.program.range_refs.push(RangeRef::new(*r, *r));
                    self.program
                        .instrs
                        .push(Instruction::new(OpCode::LoadRange, idx, 0));
                } else {
                    let idx = self.program.cell_refs.len() as u32;
                    self.program.cell_refs.push(*r);
                    self.program
                        .instrs
                        .push(Instruction::new(OpCode::LoadCell, idx, 0));
                }
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
                if let Some(info) = self.resolve_local(name) {
                    self.program
                        .instrs
                        .push(Instruction::new(OpCode::LoadLocal, info.idx, 0));
                    if !allow_range && info.kind == LocalKind::RefSingle {
                        self.program.instrs.push(Instruction::new(
                            OpCode::ImplicitIntersection,
                            0,
                            0,
                        ));
                    }
                } else if let Some(outer_info) = self
                    .closure
                    .as_ref()
                    .and_then(|closure| closure.outer_locals.get(name))
                    .copied()
                {
                    // Capture an outer local into this lambda.
                    let inner_idx = self.alloc_local(name.clone());
                    if self.lexical_scopes.is_empty() {
                        self.lexical_scopes.push(HashMap::new());
                    }
                    self.lexical_scopes[0].insert(
                        name.clone(),
                        LocalInfo {
                            idx: inner_idx,
                            kind: outer_info.kind,
                        },
                    );
                    if let Some(closure) = self.closure.as_mut() {
                        closure.captures.push(Capture {
                            outer_local: outer_info.idx,
                            inner_local: inner_idx,
                        });
                    }
                    self.program
                        .instrs
                        .push(Instruction::new(OpCode::LoadLocal, inner_idx, 0));
                    if !allow_range && outer_info.kind == LocalKind::RefSingle {
                        self.program.instrs.push(Instruction::new(
                            OpCode::ImplicitIntersection,
                            0,
                            0,
                        ));
                    }
                } else {
                    // Bytecode evaluation does not currently resolve defined names; treat as `#NAME?`.
                    self.push_error_const(ErrorKind::Name);
                }
            }
            Expr::SpillRange(inner) => {
                // The spill-range operator (`expr#`) evaluates its operand in a "reference context"
                // (i.e. it must preserve references rather than implicitly intersecting them).
                //
                // This matters for LET bindings like `LET(x, A1, x#)`: `x` is a reference value,
                // and `x#` should behave like `A1#`.
                self.compile_expr_inner(inner, true);
                self.program
                    .instrs
                    .push(Instruction::new(OpCode::SpillRange, 0, 0));
            }
            Expr::Unary { op, expr } => {
                match op {
                    UnaryOp::ImplicitIntersection => self.compile_expr_inner(expr, true),
                    UnaryOp::Plus | UnaryOp::Neg => self.compile_expr_inner(expr, false),
                }
                match op {
                    UnaryOp::Plus => {
                        self.program
                            .instrs
                            .push(Instruction::new(OpCode::UnaryPlus, 0, 0))
                    }
                    UnaryOp::Neg => {
                        self.program
                            .instrs
                            .push(Instruction::new(OpCode::UnaryNeg, 0, 0))
                    }
                    UnaryOp::ImplicitIntersection => self.program.instrs.push(Instruction::new(
                        OpCode::ImplicitIntersection,
                        0,
                        0,
                    )),
                }
            }
            Expr::Binary { op, left, right } => {
                // Reference algebra operators (`A1:A3,B1:B3` union; `A1:C3 B2:D4` intersection)
                // evaluate their operands in "reference context" so bare cell refs behave like
                // single-cell ranges and LET locals preserve reference semantics.
                let allow_range = matches!(op, BinaryOp::Union | BinaryOp::Intersect);
                self.compile_expr_inner(left, allow_range);
                self.compile_expr_inner(right, allow_range);
                let opcode = match op {
                    BinaryOp::Add => OpCode::Add,
                    BinaryOp::Sub => OpCode::Sub,
                    BinaryOp::Mul => OpCode::Mul,
                    BinaryOp::Div => OpCode::Div,
                    BinaryOp::Pow => OpCode::Pow,
                    BinaryOp::Union => OpCode::Union,
                    BinaryOp::Intersect => OpCode::Intersect,
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
                Function::Let => self.compile_let(args, allow_range),
                Function::IsOmitted => self.compile_isomitted(args),
                // Certain logical/error functions are lazy in Excel: they should only evaluate
                // the branch argument that is selected at runtime (e.g. `IF(FALSE, <expensive>, 0)`).
                //
                // Implement these with explicit control flow opcodes so the VM can short-circuit.
                Function::If if args.len() == 2 || args.len() == 3 => {
                    self.compile_expr_inner(&args[0], false);
                    let jump_idx = self.program.instrs.len();
                    // Patched below: a=false_target, b=end_target.
                    self.program
                        .instrs
                        .push(Instruction::new(OpCode::JumpIfFalseOrError, 0, 0));

                    // TRUE branch.
                    self.compile_expr_inner(&args[1], false);
                    let jump_end_idx = self.program.instrs.len();
                    self.program
                        .instrs
                        .push(Instruction::new(OpCode::Jump, 0, 0));

                    // FALSE branch.
                    let false_target = self.program.instrs.len() as u32;
                    if args.len() == 3 {
                        self.compile_expr_inner(&args[2], false);
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
                    self.program.instrs[jump_idx] =
                        Instruction::new(OpCode::JumpIfFalseOrError, false_target, end_target);
                    self.program.instrs[jump_end_idx] =
                        Instruction::new(OpCode::Jump, end_target, 0);
                }
                Function::Ifs => self.compile_ifs(args),
                Function::IfError if args.len() == 2 => {
                    // Evaluate the first argument. If it's not an error, short-circuit and
                    // return it without evaluating the fallback.
                    self.compile_expr_inner(&args[0], false);
                    let jump_idx = self.program.instrs.len();
                    self.program
                        .instrs
                        .push(Instruction::new(OpCode::JumpIfNotError, 0, 0));

                    // Error fallback branch (only evaluated if arg0 is an error).
                    self.compile_expr_inner(&args[1], false);
                    let end_target = self.program.instrs.len() as u32;
                    self.program.instrs[jump_idx] =
                        Instruction::new(OpCode::JumpIfNotError, end_target, 0);
                }
                Function::IfNa if args.len() == 2 => {
                    // Evaluate the first argument. If it's not #N/A, short-circuit and return
                    // it without evaluating the fallback.
                    self.compile_expr_inner(&args[0], false);
                    let jump_idx = self.program.instrs.len();
                    self.program
                        .instrs
                        .push(Instruction::new(OpCode::JumpIfNotNaError, 0, 0));

                    // #N/A fallback branch (only evaluated if arg0 is #N/A).
                    self.compile_expr_inner(&args[1], false);
                    let end_target = self.program.instrs.len() as u32;
                    self.program.instrs[jump_idx] =
                        Instruction::new(OpCode::JumpIfNotNaError, end_target, 0);
                }
                Function::Choose => self.compile_choose(args, allow_range),
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
            Expr::Lambda { params, body } => self.compile_lambda(params, body),
            Expr::Call { callee, args } => self.compile_call(callee, args),
        }
    }

    fn compile_lambda(&mut self, params: &Arc<[Arc<str>]>, body: &Expr) {
        let outer_locals = self.collect_visible_locals();
        let self_name = self.lambda_self_name.clone();

        let body_key = normalized_key(body);
        let mut body_program = Program::new(body_key);
        let mut body_ctx = CompileCtx::new(&mut body_program);
        body_ctx.closure = Some(ClosureCtx {
            outer_locals,
            captures: Vec::new(),
        });

        // Root scope: holds captures, optional self binding, and parameters.
        body_ctx.lexical_scopes.push(HashMap::new());

        let mut self_local = None;
        if let Some(name) = self_name {
            let idx = body_ctx.alloc_local(name.clone());
            body_ctx.lexical_scopes[0].insert(
                name,
                LocalInfo {
                    idx,
                    kind: LocalKind::Scalar,
                },
            );
            self_local = Some(idx);
        }

        let mut param_locals: Vec<u32> = Vec::new();
        let mut omitted_param_locals: Vec<u32> = Vec::new();
        if param_locals.try_reserve_exact(params.len()).is_err()
            || omitted_param_locals.try_reserve_exact(params.len()).is_err()
        {
            debug_assert!(
                false,
                "allocation failed (lambda param locals, len={})",
                params.len()
            );
            self.push_error_const(ErrorKind::Num);
            return;
        }
        let mut omitted_param_map: HashMap<Arc<str>, u32> = HashMap::new();
        for p in params.iter() {
            let idx = body_ctx.alloc_local(p.clone());
            body_ctx.lexical_scopes[0].insert(
                p.clone(),
                LocalInfo {
                    idx,
                    kind: LocalKind::Scalar,
                },
            );
            param_locals.push(idx);

            // Track omitted parameters for `ISOMITTED(...)` by allocating a hidden local per param.
            // The VM sets these locals to TRUE/FALSE at call time based on the number of supplied
            // arguments.
            let omitted_name: Arc<str> = Arc::from(format!("\0LAMBDA_OMITTED:{p}"));
            let omitted_local = body_ctx.alloc_local(omitted_name);
            omitted_param_locals.push(omitted_local);
            omitted_param_map.insert(p.clone(), omitted_local);
        }

        body_ctx.omitted_param_locals = Some(omitted_param_map);

        body_ctx.compile_expr(body);

        let Some(closure) = body_ctx.closure.take() else {
            debug_assert!(false, "lambda body compilation should have closure context");
            self.push_error_const(ErrorKind::Value);
            return;
        };

        let template = Arc::new(LambdaTemplate {
            params: params.clone(),
            body: Arc::new(body_program),
            param_locals: Arc::from(param_locals.into_boxed_slice()),
            omitted_param_locals: Arc::from(omitted_param_locals.into_boxed_slice()),
            captures: Arc::from(closure.captures.into_boxed_slice()),
            self_local,
        });
        let idx = self.program.lambdas.len() as u32;
        self.program.lambdas.push(template);
        self.program
            .instrs
            .push(Instruction::new(OpCode::MakeLambda, idx, 0));
    }

    fn compile_call(&mut self, callee: &Expr, args: &[Expr]) {
        self.compile_expr(callee);
        for arg in args {
            self.compile_call_arg(arg);
        }
        self.program
            .instrs
            .push(Instruction::new(OpCode::CallValue, 0, args.len() as u32));
    }

    fn compile_call_arg(&mut self, arg: &Expr) {
        // Preserve reference semantics for lambda arguments:
        // Excel evaluates references first and implicitly dereferences only when needed.
        // Passing a cell reference as a reference (not as a scalar value) is important for
        // correct behavior in nested aggregate calls (e.g. `LAMBDA(r, SUM(r))(A1)`).
        //
        // This also applies to LET/LAMBDA lexical locals that hold single-cell references:
        // `LET(r, A10, LAMBDA(x, ROW(x))(r))` should treat `r` as a reference, not a scalar.
        self.compile_expr_inner(arg, true);
    }

    fn collect_visible_locals(&self) -> HashMap<Arc<str>, LocalInfo> {
        let mut out = HashMap::new();
        for scope in self.lexical_scopes.iter().rev() {
            for (name, info) in scope {
                out.entry(name.clone()).or_insert(*info);
            }
        }
        out
    }

    fn infer_let_value_kind(&self, expr: &Expr) -> LocalKind {
        fn local_kind_in(
            scopes: &[HashMap<Arc<str>, LocalKind>],
            name: &Arc<str>,
        ) -> Option<LocalKind> {
            for scope in scopes.iter().rev() {
                if let Some(kind) = scope.get(name) {
                    return Some(*kind);
                }
            }
            None
        }

        fn infer_kind(
            ctx: &CompileCtx<'_>,
            expr: &Expr,
            scopes: &mut Vec<HashMap<Arc<str>, LocalKind>>,
        ) -> LocalKind {
            match expr {
                Expr::CellRef(_) => LocalKind::RefSingle,
                Expr::RangeRef(r) => {
                    if r.start == r.end {
                        LocalKind::RefSingle
                    } else {
                        LocalKind::RefMulti
                    }
                }
                Expr::NameRef(name) => local_kind_in(scopes, name)
                    .or_else(|| ctx.resolve_local(name).map(|info| info.kind))
                    .unwrap_or(LocalKind::Scalar),
                Expr::Unary { op, .. } if *op == UnaryOp::ImplicitIntersection => LocalKind::Scalar,
                Expr::FuncCall { func, args } => match func {
                    Function::Let => {
                        if args.len() < 3 || args.len() % 2 == 0 {
                            return LocalKind::Scalar;
                        }

                        scopes.push(HashMap::new());
                        for pair in args[..args.len() - 1].chunks_exact(2) {
                            let Expr::NameRef(name) = &pair[0] else {
                                scopes.pop();
                                return LocalKind::Scalar;
                            };
                            if name.is_empty() {
                                scopes.pop();
                                return LocalKind::Scalar;
                            }

                            let kind = infer_kind(ctx, &pair[1], scopes);
                            let Some(scope) = scopes.last_mut() else {
                                debug_assert!(false, "LET inference: expected an active scope");
                                return LocalKind::Scalar;
                            };
                            scope.insert(name.clone(), kind);
                        }

                        let kind = infer_kind(ctx, &args[args.len() - 1], scopes);
                        scopes.pop();
                        kind
                    }
                    Function::Choose => {
                        // The index argument is always scalar; CHOOSE's result kind depends on the
                        // selected value argument.
                        let mut out = LocalKind::Scalar;
                        for arg in args.iter().skip(1) {
                            match infer_kind(ctx, arg, scopes) {
                                LocalKind::Scalar => {}
                                LocalKind::RefSingle => out = LocalKind::RefSingle,
                                LocalKind::RefMulti => return LocalKind::RefMulti,
                            }
                        }
                        out
                    }
                    _ => LocalKind::Scalar,
                },
                _ => LocalKind::Scalar,
            }
        }

        infer_kind(self, expr, &mut Vec::new())
    }

    fn compile_let(&mut self, args: &[Expr], allow_range: bool) {
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
            if name.is_empty() {
                self.lexical_scopes.pop();
                self.push_error_const(ErrorKind::Value);
                return;
            }

            // Evaluate the binding value, then store it into a new local slot.
            // The binding name becomes visible to subsequent bindings and the final body.
            let kind = self.infer_let_value_kind(&pair[1]);

            // Allow the LET binding name to be referenced inside any LAMBDA bodies produced by
            // the value expression (for recursion via `f(x)`).
            let prev_self_name = self.lambda_self_name.take();
            self.lambda_self_name = Some(name.clone());

            // LET binding values are evaluated in "argument mode" (may preserve references).
            self.compile_expr_inner(&pair[1], true);

            self.lambda_self_name = prev_self_name;

            let idx = self.alloc_local(name.clone());
            self.program
                .instrs
                .push(Instruction::new(OpCode::StoreLocal, idx, 0));
            self.lexical_scopes
                .last_mut()
                .map(|scope| scope.insert(name.clone(), LocalInfo { idx, kind }))
                .unwrap_or_else(|| {
                    debug_assert!(false, "LET compile: expected an active scope");
                    None
                });
        }

        self.compile_expr_inner(&args[last], allow_range);
        self.lexical_scopes.pop();
    }

    fn compile_isomitted(&mut self, args: &[Expr]) {
        if args.len() != 1 {
            self.push_error_const(ErrorKind::Value);
            return;
        }

        let Expr::NameRef(name) = &args[0] else {
            self.push_error_const(ErrorKind::Value);
            return;
        };

        // `ISOMITTED` only reports omission for parameters of the *current* lambda invocation.
        // Outside of a lambda (or when referring to a non-parameter identifier), it returns FALSE.
        let omitted_local = self
            .omitted_param_locals
            .as_ref()
            .and_then(|map| map.get(name))
            .copied();

        match omitted_local {
            Some(idx) => {
                self.program
                    .instrs
                    .push(Instruction::new(OpCode::LoadLocal, idx, 0));
            }
            None => {
                let idx = self.program.consts.len() as u32;
                self.program
                    .consts
                    .push(ConstValue::Value(Value::Bool(false)));
                self.program
                    .instrs
                    .push(Instruction::new(OpCode::PushConst, idx, 0));
            }
        }
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

    fn compile_choose(&mut self, args: &[Expr], allow_range: bool) {
        // CHOOSE(index, value1, value2, ...)
        // Excel semantics:
        // - Evaluate `index` once.
        // - Coerce to number, truncate toward zero (NOT floor), and select the 1-based branch.
        // - If index is an error, return it without evaluating any branch.
        // - If index is out of range, return #VALUE! without evaluating any branch.
        //
        // This is compiled into explicit control-flow so unselected branches are not evaluated.
        if args.len() < 2 {
            self.push_error_const(ErrorKind::Value);
            return;
        }

        let choice_count = args.len() - 1;

        // Normalize the index once using the runtime CHOOSE implementation over constants
        // 1..=choice_count. This preserves the engine's coercion rules (notably NaN -> 0 via
        // float-to-int casts) while ensuring the actual choice expressions remain lazy.
        //
        // Avoid normalizing via `INT(index)` + comparisons: the engine's Excel-style comparisons
        // treat NaN as "equal" for ordering purposes, which can incorrectly select a branch and
        // eagerly evaluate a choice expression.
        self.compile_expr_inner(&args[0], false);
        let mut selector_consts: Vec<u32> = Vec::new();
        if selector_consts.try_reserve_exact(choice_count).is_err() {
            debug_assert!(
                false,
                "allocation failed (choose selector consts, len={choice_count})"
            );
            self.push_error_const(ErrorKind::Num);
            return;
        }
        for i in 1..=choice_count {
            let idx = self.program.consts.len() as u32;
            self.program
                .consts
                .push(ConstValue::Value(Value::Number(i as f64)));
            selector_consts.push(idx);
            self.program
                .instrs
                .push(Instruction::new(OpCode::PushConst, idx, 0));
        }
        let func_idx = intern_func(self.program, Function::Choose);
        self.program.instrs.push(Instruction::new(
            OpCode::CallFunc,
            func_idx,
            args.len() as u32,
        ));

        // Store the selection result so we can branch on it multiple times without re-evaluating
        // the index expression.
        let sel_local = self.alloc_temp_local("\u{0}CHOOSE_SEL");
        self.program
            .instrs
            .push(Instruction::new(OpCode::StoreLocal, sel_local, 0));

        // If selection is an error (invalid index/coercion), return it without evaluating any
        // branch expressions.
        self.program
            .instrs
            .push(Instruction::new(OpCode::LoadLocal, sel_local, 0));
        let jump_not_error_idx = self.program.instrs.len();
        self.program
            .instrs
            .push(Instruction::new(OpCode::JumpIfNotError, 0, 0));
        // Error path: reload the error and jump to end.
        self.program
            .instrs
            .push(Instruction::new(OpCode::LoadLocal, sel_local, 0));
        let jump_error_end_idx = self.program.instrs.len();
        self.program
            .instrs
            .push(Instruction::new(OpCode::Jump, 0, 0));

        let cases_start = self.program.instrs.len() as u32;
        self.program.instrs[jump_not_error_idx] =
            Instruction::new(OpCode::JumpIfNotError, cases_start, 0);

        let mut jump_if_idxs: Vec<usize> = Vec::new();
        let mut jump_end_idxs: Vec<usize> = vec![jump_error_end_idx];

        for (idx, choice_expr) in args[1..].iter().enumerate() {
            if idx != 0 {
                self.program
                    .instrs
                    .push(Instruction::new(OpCode::LoadLocal, sel_local, 0));
            }

            // Compare selection == (idx + 1).
            let const_idx = selector_consts[idx];
            self.program
                .instrs
                .push(Instruction::new(OpCode::PushConst, const_idx, 0));
            self.program.instrs.push(Instruction::new(OpCode::Eq, 0, 0));

            let jump_idx = self.program.instrs.len();
            // Patched below: a=next_case_target, b=end_target (for error propagation).
            self.program
                .instrs
                .push(Instruction::new(OpCode::JumpIfFalseOrError, 0, 0));

            // Match branch: evaluate the selected choice expression and jump to end.
            //
            // CHOOSE can return references; when the surrounding expression allows ranges (e.g.
            // `SUM(CHOOSE(1, A1, B1))` / `SUM(CHOOSE(1, A1:A3, B1:B3))`), propagate `allow_range`
            // so bare cell references are treated like reference arguments.
            self.compile_expr_inner(choice_expr, allow_range);

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

        // Should be unreachable because the selection call above guarantees the index is in range
        // when it's not an error, but keep a deterministic fallback.
        self.push_error_const(ErrorKind::Value);
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
        let pairs_end = if has_default {
            args.len() - 1
        } else {
            args.len()
        };
        let pairs = &args[1..pairs_end];
        let default = if has_default {
            Some(&args[args.len() - 1])
        } else {
            None
        };

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
            self.program.instrs.push(Instruction::new(OpCode::Eq, 0, 0));

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
        // Preserve `Missing` for *direct* blank arguments so functions can distinguish a
        // syntactically omitted argument from a blank cell value.
        if matches!(arg, Expr::Literal(Value::Missing)) {
            let idx = self.program.consts.len() as u32;
            self.program.consts.push(ConstValue::Value(Value::Missing));
            self.program
                .instrs
                .push(Instruction::new(OpCode::PushConst, idx, 0));
            return;
        }

        // Excel-style aggregate functions have a quirk: a cell reference passed directly as an
        // argument (e.g. `SUM(A1)`) is treated as a *reference* argument, not a scalar, which means
        // logical/text values in the referenced cell are ignored (unlike `SUM(TRUE)` / `SUM("5")`).
        //
        // Lower single-cell references to a range so the runtime can apply reference semantics.
        let allow_range = match func {
            // AND/OR accept range arguments (which ignore text-like values), but direct *cell*
            // references are treated as scalar arguments (so text-like values produce #VALUE!).
            Function::And | Function::Or => true,
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
            // Bytecode currently represents lookup vectors as range values; lower single-cell
            // references to ranges so XMATCH/XLOOKUP can treat them uniformly.
            Function::XMatch => arg_idx == 1,
            Function::XLookup => arg_idx == 1 || arg_idx == 2,
            // OFFSET's first argument is reference-valued: even a single-cell ref like `A1` must
            // preserve reference semantics so the runtime can compute the shifted rectangle.
            Function::Offset => arg_idx == 0,
            Function::Row | Function::Column | Function::Rows | Function::Columns => true,
            Function::Rand | Function::RandBetween => false,
            _ => false,
        };

        let allow_range =
            if matches!(func, Function::And | Function::Or) && matches!(arg, Expr::CellRef(_)) {
                false
            } else {
                allow_range
            };

        self.compile_expr_inner(arg, allow_range);
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
        Value::Missing => out.push_str("M"),
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
                sheet_range_to_key(area, out);
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
            for v in a.values.iter() {
                value_to_key(v, out);
                out.push(',');
            }
            out.push(')');
        }
        // Rich value variants (e.g. entity / record types) can be carried through the bytecode
        // runtime for pass-through semantics, but are not currently produced by parsing/lowering
        // as literals. Encode them as an opaque marker so cache keys stay deterministic without
        // needing to serialize the rich payload.
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
                sheet_range_to_key(area, out);
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
                BinaryOp::Union => ",",
                BinaryOp::Intersect => " ",
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
        Expr::Lambda { params, body } => {
            out.push_str("LAMBDA(");
            out.push_str(&params.len().to_string());
            out.push(':');
            for p in params.iter() {
                out.push_str(p);
                out.push(',');
            }
            out.push_str("BODY=");
            expr_to_key(body, out);
            out.push(')');
        }
        Expr::Call { callee, args } => {
            out.push_str("CALL(");
            expr_to_key(callee, out);
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

fn sheet_id_to_key(sheet: &SheetId, out: &mut String) {
    match sheet {
        SheetId::Local(id) => {
            out.push('L');
            out.push_str(&id.to_string());
        }
        SheetId::External(key) => {
            out.push('E');
            out.push_str(key);
            out.push('\0');
        }
        SheetId::ExternalSpan {
            workbook,
            start,
            end,
        } => {
            out.push('X');
            out.push_str(workbook);
            out.push(']');
            out.push_str(start);
            out.push(':');
            out.push_str(end);
            out.push('\0');
        }
    }
}

fn sheet_range_to_key(r: &SheetRangeRef, out: &mut String) {
    out.push_str("S");
    sheet_id_to_key(&r.sheet, out);
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
