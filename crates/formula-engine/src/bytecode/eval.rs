use super::ast::{BinaryOp, UnaryOp};
use super::grid::Grid;
use super::runtime::{
    apply_binary, apply_implicit_intersection, apply_unary, call_function, deref_value_dynamic,
};
use super::value::{CellCoord, ErrorKind, Lambda, Value};
use crate::date::ExcelDateSystem;
use crate::locale::ValueLocaleConfig;
use chrono::{DateTime, Utc};
use std::sync::Arc;

use super::program::{OpCode, Program};

/// Stack-based bytecode interpreter.
#[derive(Default)]
pub struct Vm {
    stack: Vec<Value>,
    locals: Vec<Value>,
    lambda_depth: u32,
    sheet_id: usize,
}

// Keep lambda recursion bounded well below the Rust stack limit to avoid process aborts for
// accidental infinite recursion (matching the AST evaluator).
const LAMBDA_RECURSION_LIMIT: u32 = 64;

// Small stack reservation used by the engine's bytecode execution loop.
pub(crate) const DEFAULT_VM_STACK_RESERVE: usize = 32;

impl Vm {
    pub fn new() -> Self {
        Self {
            stack: Vec::new(),
            locals: Vec::new(),
            lambda_depth: 0,
            sheet_id: 0,
        }
    }

    pub(crate) fn new_with_default_stack() -> Self {
        Self::with_capacity(DEFAULT_VM_STACK_RESERVE)
    }

    #[inline]
    fn push_stack(&mut self, v: Value) -> bool {
        if self.stack.len() == self.stack.capacity() {
            // Grow exponentially (bounded) to amortize capacity checks.
            let additional = self.stack.len().max(1).min(1024);
            if self.stack.try_reserve(additional).is_err() {
                debug_assert!(
                    false,
                    "allocation failed (Vm stack grow, len={})",
                    self.stack.len()
                );
                return false;
            }
        }
        self.stack.push(v);
        true
    }

    pub fn with_capacity(stack: usize) -> Self {
        let mut stack_vec: Vec<Value> = Vec::new();
        if stack_vec.try_reserve_exact(stack).is_err() {
            debug_assert!(false, "allocation failed (Vm stack, len={stack})");
        }
        Self {
            stack: stack_vec,
            locals: Vec::new(),
            lambda_depth: 0,
            sheet_id: 0,
        }
    }

    pub fn eval(
        &mut self,
        program: &Program,
        grid: &dyn Grid,
        sheet_id: usize,
        base: CellCoord,
        locale: &crate::LocaleConfig,
    ) -> Value {
        super::runtime::set_thread_current_sheet_id(sheet_id);
        super::runtime::reset_thread_rng_counter();
        self.sheet_id = sheet_id;
        self.stack.clear();
        self.locals.clear();
        let locals_len = program.locals.len();
        if self.locals.try_reserve_exact(locals_len).is_err() {
            debug_assert!(false, "allocation failed (Vm locals, len={locals_len})");
            return Value::Error(ErrorKind::Num);
        }
        self.locals.resize(locals_len, Value::Empty);
        self.eval_program(program, grid, base, locale)
    }

    fn eval_program(
        &mut self,
        program: &Program,
        grid: &dyn Grid,
        base: CellCoord,
        locale: &crate::LocaleConfig,
    ) -> Value {
        let v = self.eval_program_raw(program, grid, base, locale);
        // Match the AST evaluator: the final result uses dynamic dereference, so range references
        // spill instead of producing a scalar `#SPILL!`.
        deref_value_dynamic(v, grid, base)
    }

    fn eval_program_raw(
        &mut self,
        program: &Program,
        grid: &dyn Grid,
        base: CellCoord,
        locale: &crate::LocaleConfig,
    ) -> Value {
        let instrs = program.instrs();
        let mut pc: usize = 0;
        while pc < instrs.len() {
            let inst = instrs[pc];
            let op = inst.op();
            match op {
                OpCode::Invalid => return Value::Error(ErrorKind::Value),
                OpCode::PushConst => {
                    let v = program.consts[inst.a() as usize].to_value();
                    if !self.push_stack(v) {
                        return Value::Error(ErrorKind::Num);
                    }
                }
                OpCode::LoadCell => {
                    let r = program.cell_refs[inst.a() as usize];
                    if !self.push_stack(grid.get_value(r.resolve(base))) {
                        return Value::Error(ErrorKind::Num);
                    }
                }
                OpCode::LoadRange => {
                    let r = program.range_refs[inst.a() as usize];
                    if !self.push_stack(Value::Range(r)) {
                        return Value::Error(ErrorKind::Num);
                    }
                }
                OpCode::LoadMultiRange => {
                    let r = program.multi_range_refs[inst.a() as usize].clone();
                    if !self.push_stack(Value::MultiRange(r)) {
                        return Value::Error(ErrorKind::Num);
                    }
                }
                OpCode::StoreLocal => {
                    let v = self.stack.pop().unwrap_or(Value::Empty);
                    let idx = inst.a() as usize;
                    if idx >= self.locals.len() {
                        let new_len = idx + 1;
                        let additional = new_len.saturating_sub(self.locals.len());
                        if self.locals.try_reserve_exact(additional).is_err() {
                            debug_assert!(false, "allocation failed (Vm locals grow, len={new_len})");
                            return Value::Error(ErrorKind::Num);
                        }
                        self.locals.resize(new_len, Value::Empty);
                    }
                    self.locals[idx] = v;
                }
                OpCode::LoadLocal => {
                    let idx = inst.a() as usize;
                    let v = self.locals.get(idx).cloned().unwrap_or(Value::Empty);
                    if !self.push_stack(v) {
                        return Value::Error(ErrorKind::Num);
                    }
                }
                OpCode::Jump => {
                    pc = inst.a() as usize;
                    continue;
                }
                OpCode::UnaryPlus => {
                    let v = self.stack.pop().unwrap_or(Value::Empty);
                    if !self.push_stack(apply_unary(UnaryOp::Plus, v, grid, base)) {
                        return Value::Error(ErrorKind::Num);
                    }
                }
                OpCode::UnaryNeg => {
                    let v = self.stack.pop().unwrap_or(Value::Empty);
                    if !self.push_stack(apply_unary(UnaryOp::Neg, v, grid, base)) {
                        return Value::Error(ErrorKind::Num);
                    }
                }
                OpCode::ImplicitIntersection => {
                    let v = self.stack.pop().unwrap_or(Value::Empty);
                    if !self.push_stack(apply_implicit_intersection(v, grid, base)) {
                        return Value::Error(ErrorKind::Num);
                    }
                }
                OpCode::Add
                | OpCode::Sub
                | OpCode::Mul
                | OpCode::Div
                | OpCode::Pow
                | OpCode::Union
                | OpCode::Intersect
                | OpCode::Eq
                | OpCode::Ne
                | OpCode::Lt
                | OpCode::Le
                | OpCode::Gt
                | OpCode::Ge => {
                    let right = self.stack.pop().unwrap_or(Value::Empty);
                    let left = self.stack.pop().unwrap_or(Value::Empty);
                    let op = match op {
                        OpCode::Add => BinaryOp::Add,
                        OpCode::Sub => BinaryOp::Sub,
                        OpCode::Mul => BinaryOp::Mul,
                        OpCode::Div => BinaryOp::Div,
                        OpCode::Pow => BinaryOp::Pow,
                        OpCode::Union => BinaryOp::Union,
                        OpCode::Intersect => BinaryOp::Intersect,
                        OpCode::Eq => BinaryOp::Eq,
                        OpCode::Ne => BinaryOp::Ne,
                        OpCode::Lt => BinaryOp::Lt,
                        OpCode::Le => BinaryOp::Le,
                        OpCode::Gt => BinaryOp::Gt,
                        OpCode::Ge => BinaryOp::Ge,
                        _ => {
                            debug_assert!(false, "invalid binary opcode: {op:?}");
                            return Value::Error(ErrorKind::Value);
                        }
                    };
                    let v = apply_binary(op, left, right, grid, self.sheet_id, base);
                    if !self.push_stack(v) {
                        return Value::Error(ErrorKind::Num);
                    }
                }
                OpCode::CallFunc => {
                    let func = &program.funcs[inst.a() as usize];
                    let argc = inst.b() as usize;
                    let start = self.stack.len().saturating_sub(argc);
                    let result = call_function(func, &self.stack[start..], grid, base, locale);
                    // `Value::Missing` is an internal placeholder for syntactically blank arguments.
                    // It must not escape as a runtime value result, otherwise downstream calls may
                    // misinterpret it as an omitted argument and apply incorrect defaulting.
                    let result = if matches!(result, Value::Missing) {
                        Value::Empty
                    } else {
                        result
                    };
                    self.stack.truncate(start);
                    if !self.push_stack(result) {
                        return Value::Error(ErrorKind::Num);
                    }
                }
                OpCode::SpillRange => {
                    let v = self.stack.pop().unwrap_or(Value::Empty);
                    let spilled = super::runtime::apply_spill_range(
                        v,
                        grid,
                        self.sheet_id,
                        base,
                    );
                    if !self.push_stack(spilled) {
                        return Value::Error(ErrorKind::Num);
                    }
                }
                OpCode::MakeLambda => {
                    let template = program.lambdas[inst.a() as usize].clone();
                    let mut captures: Vec<Value> = Vec::new();
                    if captures.try_reserve_exact(template.captures.len()).is_err() {
                        debug_assert!(
                            false,
                            "allocation failed (lambda captures, len={})",
                            template.captures.len()
                        );
                        if !self.push_stack(Value::Error(ErrorKind::Num)) {
                            return Value::Error(ErrorKind::Num);
                        }
                        pc += 1;
                        continue;
                    }
                    for cap in template.captures.iter() {
                        let outer_idx = cap.outer_local as usize;
                        captures.push(self.locals.get(outer_idx).cloned().unwrap_or(Value::Empty));
                    }
                    if !self.push_stack(Value::Lambda(Lambda {
                        template,
                        captures: Arc::from(captures.into_boxed_slice()),
                    })) {
                        return Value::Error(ErrorKind::Num);
                    }
                }
                OpCode::CallValue => {
                    let argc = inst.b() as usize;
                    if argc > crate::EXCEL_MAX_ARGS {
                        // Should be prevented by parsing/lowering.
                        self.stack.truncate(self.stack.len().saturating_sub(argc));
                        if !self.push_stack(Value::Error(ErrorKind::Value)) {
                            return Value::Error(ErrorKind::Num);
                        }
                        pc += 1;
                        continue;
                    }

                    let mut args: Vec<Value> = Vec::new();
                    if args.try_reserve_exact(argc).is_err() {
                        debug_assert!(false, "allocation failed (call args, len={argc})");
                        self.stack.truncate(self.stack.len().saturating_sub(argc + 1));
                        if !self.push_stack(Value::Error(ErrorKind::Num)) {
                            return Value::Error(ErrorKind::Num);
                        }
                        pc += 1;
                        continue;
                    }
                    for _ in 0..argc {
                        args.push(self.stack.pop().unwrap_or(Value::Empty));
                    }
                    args.reverse();
                    let callee = self.stack.pop().unwrap_or(Value::Empty);

                    let result = match callee {
                        Value::Lambda(lambda) => self.call_lambda(lambda, args, grid, base, locale),
                        Value::Error(e) => Value::Error(e),
                        _ => Value::Error(ErrorKind::Value),
                    };
                    if !self.push_stack(result) {
                        return Value::Error(ErrorKind::Num);
                    }
                }
                OpCode::JumpIfFalseOrError => {
                    let v = self.stack.pop().unwrap_or(Value::Empty);
                    // Match the evaluator semantics used by logical functions like IF/IFS:
                    // a single-cell reference should behave like a scalar value (while multi-cell
                    // references still surface #SPILL! via array coercion).
                    let v = super::runtime::deref_value_dynamic(v, grid, base);
                    match super::runtime::coerce_to_bool(&v) {
                        Ok(true) => {}
                        Ok(false) => {
                            pc = inst.a() as usize;
                            continue;
                        }
                        Err(e) => {
                            if !self.push_stack(Value::Error(e)) {
                                return Value::Error(ErrorKind::Num);
                            }
                            pc = inst.b() as usize;
                            continue;
                        }
                    }
                }
                OpCode::JumpIfNotError => {
                    let is_error = matches!(self.stack.last(), Some(Value::Error(_)));
                    if !is_error {
                        pc = inst.a() as usize;
                        continue;
                    }
                    let _ = self.stack.pop();
                }
                OpCode::JumpIfNotNaError => {
                    let is_na = matches!(
                        self.stack.last(),
                        Some(Value::Error(super::value::ErrorKind::NA))
                    );
                    if !is_na {
                        pc = inst.a() as usize;
                        continue;
                    }
                    let _ = self.stack.pop();
                }
            }
            pc += 1;
        }
        self.stack.pop().unwrap_or(Value::Empty)
    }

    fn call_lambda(
        &mut self,
        lambda: Lambda,
        args: Vec<Value>,
        grid: &dyn Grid,
        base: CellCoord,
        locale: &crate::LocaleConfig,
    ) -> Value {
        if args.len() > crate::EXCEL_MAX_ARGS {
            return Value::Error(ErrorKind::Value);
        }

        if args.len() > lambda.template.params.len() {
            return Value::Error(ErrorKind::Value);
        }

        if self.lambda_depth >= LAMBDA_RECURSION_LIMIT {
            return Value::Error(ErrorKind::Calc);
        }

        self.lambda_depth += 1;

        let body_program = lambda.template.body.clone();
        let locals_len = body_program.locals.len();
        let mut locals: Vec<Value> = Vec::new();
        if locals.try_reserve_exact(locals_len).is_err() {
            debug_assert!(
                false,
                "allocation failed (lambda locals, len={locals_len})"
            );
            self.lambda_depth = self.lambda_depth.saturating_sub(1);
            return Value::Error(ErrorKind::Num);
        }
        locals.resize(locals_len, Value::Empty);

        // Populate captured values.
        debug_assert_eq!(lambda.template.captures.len(), lambda.captures.len());
        for (cap, value) in lambda.template.captures.iter().zip(lambda.captures.iter()) {
            if let Some(slot) = locals.get_mut(cap.inner_local as usize) {
                *slot = value.clone();
            }
        }

        // Bind the lambda value itself for recursion (if requested by the compiler).
        if let Some(self_idx) = lambda.template.self_local {
            if let Some(slot) = locals.get_mut(self_idx as usize) {
                *slot = Value::Lambda(lambda.clone());
            }
        }

        // Bind parameters. Missing args are treated as blank.
        for (idx, local_idx) in lambda.template.param_locals.iter().copied().enumerate() {
            if let Some(slot) = locals.get_mut(local_idx as usize) {
                *slot = args.get(idx).cloned().unwrap_or(Value::Empty);
            }
        }

        // Track omitted parameters for `ISOMITTED(...)`.
        debug_assert_eq!(
            lambda.template.params.len(),
            lambda.template.omitted_param_locals.len()
        );
        for (idx, local_idx) in lambda
            .template
            .omitted_param_locals
            .iter()
            .copied()
            .enumerate()
        {
            if let Some(slot) = locals.get_mut(local_idx as usize) {
                *slot = Value::Bool(idx >= args.len());
            }
        }

        // Evaluate the lambda body with a fresh stack + locals.
        let saved_stack = std::mem::take(&mut self.stack);
        let saved_locals = std::mem::take(&mut self.locals);

        self.stack = Vec::new();
        self.locals = locals;
        // Lambdas can return references, which should be preserved so that reference-only
        // functions (e.g. ROW/COLUMN) can consume the reference value without forcing an eager
        // dereference/spill inside the lambda body.
        let result = self.eval_program_raw(&body_program, grid, base, locale);

        self.stack = saved_stack;
        self.locals = saved_locals;

        self.lambda_depth = self.lambda_depth.saturating_sub(1);
        result
    }

    pub fn eval_with_value_locale(
        &mut self,
        program: &Program,
        grid: &dyn Grid,
        sheet_id: usize,
        base: CellCoord,
        value_locale: ValueLocaleConfig,
    ) -> Value {
        // Preserve the existing public API while ensuring locale-aware coercion for text values
        // matches the main evaluator. This uses Excel's default 1900 date system and the current
        // wall-clock time for any date strings that omit a year.
        self.eval_with_coercion_context(
            program,
            grid,
            sheet_id,
            base,
            ExcelDateSystem::EXCEL_1900,
            value_locale,
            Utc::now(),
        )
    }

    pub fn eval_with_coercion_context(
        &mut self,
        program: &Program,
        grid: &dyn Grid,
        sheet_id: usize,
        base: CellCoord,
        date_system: ExcelDateSystem,
        value_locale: ValueLocaleConfig,
        now_utc: DateTime<Utc>,
    ) -> Value {
        // Treat each explicit evaluation call as its own "recalc tick" for volatile functions.
        // Use the supplied `now_utc` (which callers can freeze) to derive a deterministic id.
        let recalc_id = now_utc.timestamp_nanos_opt().unwrap_or(0) as u64;
        let _guard =
            super::runtime::set_thread_eval_context(date_system, value_locale, now_utc, recalc_id);

        // Criteria strings inside quotes should follow the workbook/value locale for numeric parsing.
        let mut locale_config = crate::LocaleConfig::en_us();
        locale_config.decimal_separator = value_locale.separators.decimal_sep;
        locale_config.thousands_separator = Some(value_locale.separators.thousands_sep);
        self.eval(program, grid, sheet_id, base, &locale_config)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::date::{ymd_to_serial, ExcelDate};
    use chrono::TimeZone;
    use std::sync::atomic::{AtomicUsize, Ordering};

    #[test]
    fn eval_with_value_locale_parses_numeric_text_using_locale() {
        let origin = CellCoord::new(0, 0);
        let expr = super::super::parse_formula("=\"1.234,56\"+1", origin).expect("parse");
        let cache = super::super::BytecodeCache::new();
        let program = cache.get_or_compile(&expr);
        let grid = super::super::grid::ColumnarGrid::new(1, 1);

        let mut vm = Vm::with_capacity(32);
        let value =
            vm.eval_with_value_locale(&program, &grid, 0, origin, ValueLocaleConfig::de_de());
        match value {
            Value::Number(n) => assert!((n - 1235.56).abs() < 1e-9, "got {n}"),
            other => panic!("expected Value::Number, got {other:?}"),
        }
    }

    #[test]
    fn eval_with_coercion_context_respects_date_system() {
        let origin = CellCoord::new(0, 0);
        let expr = super::super::parse_formula("=\"2020-01-01\"+0", origin).expect("parse");
        let cache = super::super::BytecodeCache::new();
        let program = cache.get_or_compile(&expr);
        let grid = super::super::grid::ColumnarGrid::new(1, 1);

        let expected =
            ymd_to_serial(ExcelDate::new(2020, 1, 1), ExcelDateSystem::Excel1904).unwrap() as f64;

        let mut vm = Vm::with_capacity(32);
        let value = vm.eval_with_coercion_context(
            &program,
            &grid,
            0,
            origin,
            ExcelDateSystem::Excel1904,
            ValueLocaleConfig::en_us(),
            Utc.with_ymd_and_hms(2024, 1, 1, 0, 0, 0).unwrap(),
        );

        match value {
            Value::Number(n) => assert!((n - expected).abs() < 1e-9, "got {n}"),
            other => panic!("expected Value::Number, got {other:?}"),
        }
    }

    #[test]
    fn eval_with_coercion_context_uses_now_year_for_missing_year_dates() {
        let origin = CellCoord::new(0, 0);
        let expr = super::super::parse_formula("=\"1/2\"+0", origin).expect("parse");
        let cache = super::super::BytecodeCache::new();
        let program = cache.get_or_compile(&expr);
        let grid = super::super::grid::ColumnarGrid::new(1, 1);

        let now_utc = Utc.with_ymd_and_hms(2024, 6, 15, 0, 0, 0).unwrap();
        let expected =
            ymd_to_serial(ExcelDate::new(2024, 1, 2), ExcelDateSystem::EXCEL_1900).unwrap() as f64;
        let mut vm = Vm::with_capacity(32);
        let value = vm.eval_with_coercion_context(
            &program,
            &grid,
            0,
            origin,
            ExcelDateSystem::EXCEL_1900,
            ValueLocaleConfig::en_us(),
            now_utc,
        );

        match value {
            Value::Number(n) => assert!((n - expected).abs() < 1e-9, "got {n}"),
            other => panic!("expected Value::Number, got {other:?}"),
        }
    }

    struct CountingGrid {
        inner: super::super::grid::ColumnarGrid,
        reads: AtomicUsize,
    }

    impl CountingGrid {
        fn new(rows: i32, cols: i32) -> Self {
            Self {
                inner: super::super::grid::ColumnarGrid::new(rows, cols),
                reads: AtomicUsize::new(0),
            }
        }

        fn reads(&self) -> usize {
            self.reads.load(Ordering::SeqCst)
        }
    }

    impl super::super::grid::Grid for CountingGrid {
        fn get_value(&self, coord: CellCoord) -> Value {
            self.reads.fetch_add(1, Ordering::SeqCst);
            self.inner.get_value(coord)
        }

        fn column_slice(&self, col: i32, row_start: i32, row_end: i32) -> Option<&[f64]> {
            // Count columnar reads too (used by bulk range functions like SUM) so short-circuit
            // tests can catch unused branch evaluation even when it uses column slices instead of
            // per-cell `get_value` calls.
            self.reads.fetch_add(1, Ordering::SeqCst);
            self.inner.column_slice(col, row_start, row_end)
        }

        fn bounds(&self) -> (i32, i32) {
            self.inner.bounds()
        }
    }

    struct NanIndexGrid {
        reads: AtomicUsize,
    }

    impl NanIndexGrid {
        fn new() -> Self {
            Self {
                reads: AtomicUsize::new(0),
            }
        }

        fn reads(&self) -> usize {
            self.reads.load(Ordering::SeqCst)
        }
    }

    impl super::super::grid::Grid for NanIndexGrid {
        fn get_value(&self, coord: CellCoord) -> Value {
            self.reads.fetch_add(1, Ordering::SeqCst);
            if coord == CellCoord::new(0, 0) {
                return Value::Number(f64::NAN);
            }
            if coord == CellCoord::new(1, 0) {
                panic!("unexpected evaluation of CHOOSE branch expression");
            }
            Value::Empty
        }

        fn column_slice(&self, _col: i32, _row_start: i32, _row_end: i32) -> Option<&[f64]> {
            None
        }

        fn bounds(&self) -> (i32, i32) {
            (10, 10)
        }
    }

    #[test]
    fn vm_short_circuits_if_branches() {
        let origin = CellCoord::new(0, 0);
        let locale = crate::LocaleConfig::en_us();

        // IF(FALSE, A1, 1) should not evaluate the TRUE branch.
        let expr = super::super::parse_formula("=IF(FALSE, A1, 1)", origin).expect("parse");
        let program = super::super::BytecodeCache::new().get_or_compile(&expr);

        let grid = CountingGrid::new(10, 10);
        let mut vm = Vm::with_capacity(32);
        let value = vm.eval(&program, &grid, 0, origin, &locale);
        assert_eq!(value, Value::Number(1.0));
        assert_eq!(grid.reads(), 0, "unused IF branch should not be evaluated");

        // IF(FALSE, SUM(A1:A10), 1) should not evaluate the TRUE branch, even though it would use
        // column-slice access when evaluated.
        let expr =
            super::super::parse_formula("=IF(FALSE, SUM(A1:A10), 1)", origin).expect("parse");
        let program = super::super::BytecodeCache::new().get_or_compile(&expr);
        let grid = CountingGrid::new(10, 10);
        let mut vm = Vm::with_capacity(32);
        let value = vm.eval(&program, &grid, 0, origin, &locale);
        assert_eq!(value, Value::Number(1.0));
        assert_eq!(
            grid.reads(),
            0,
            "unused IF branch should not be evaluated (including range reads)"
        );

        // IF(TRUE, 1, A1) should not evaluate the FALSE branch.
        let expr = super::super::parse_formula("=IF(TRUE, 1, A1)", origin).expect("parse");
        let program = super::super::BytecodeCache::new().get_or_compile(&expr);

        let grid = CountingGrid::new(10, 10);
        let mut vm = Vm::with_capacity(32);
        let value = vm.eval(&program, &grid, 0, origin, &locale);

        assert_eq!(value, Value::Number(1.0));
        assert_eq!(grid.reads(), 0, "unused IF branch should not be evaluated");

        // IF(TRUE, 1, SUM(A1:A10)) should not evaluate the FALSE branch, even though it would use
        // column-slice access when evaluated.
        let expr = super::super::parse_formula("=IF(TRUE, 1, SUM(A1:A10))", origin).expect("parse");
        let program = super::super::BytecodeCache::new().get_or_compile(&expr);
        let grid = CountingGrid::new(10, 10);
        let mut vm = Vm::with_capacity(32);
        let value = vm.eval(&program, &grid, 0, origin, &locale);
        assert_eq!(value, Value::Number(1.0));
        assert_eq!(
            grid.reads(),
            0,
            "unused IF branch should not be evaluated (including range reads)"
        );

        // If the IF condition is an error, neither branch should be evaluated and the error should
        // be returned.
        let expr = super::super::parse_formula("=IF(1/0, A1, 1)", origin).expect("parse");
        let program = super::super::BytecodeCache::new().get_or_compile(&expr);

        let grid = CountingGrid::new(10, 10);
        let mut vm = Vm::with_capacity(32);
        let value = vm.eval(&program, &grid, 0, origin, &locale);

        assert_eq!(value, Value::Error(super::super::value::ErrorKind::Div0));
        assert_eq!(
            grid.reads(),
            0,
            "IF branches should not be evaluated when condition is an error"
        );
    }

    #[test]
    fn vm_short_circuits_iferror_and_ifna_fallbacks() {
        let origin = CellCoord::new(0, 0);
        let locale = crate::LocaleConfig::en_us();

        // IFERROR(1, A1) should not evaluate the fallback.
        let expr = super::super::parse_formula("=IFERROR(1, A1)", origin).expect("parse");
        let program = super::super::BytecodeCache::new().get_or_compile(&expr);
        let grid = CountingGrid::new(10, 10);
        let mut vm = Vm::with_capacity(32);
        let value = vm.eval(&program, &grid, 0, origin, &locale);
        assert_eq!(value, Value::Number(1.0));
        assert_eq!(
            grid.reads(),
            0,
            "unused IFERROR fallback should not be evaluated"
        );

        // IFERROR(1, SUM(A1:A10)) should not evaluate the fallback.
        let expr = super::super::parse_formula("=IFERROR(1, SUM(A1:A10))", origin).expect("parse");
        let program = super::super::BytecodeCache::new().get_or_compile(&expr);
        let grid = CountingGrid::new(10, 10);
        let mut vm = Vm::with_capacity(32);
        let value = vm.eval(&program, &grid, 0, origin, &locale);
        assert_eq!(value, Value::Number(1.0));
        assert_eq!(
            grid.reads(),
            0,
            "unused IFERROR fallback should not be evaluated (including range reads)"
        );

        // IFERROR(1/0, A1) should evaluate the fallback.
        let expr = super::super::parse_formula("=IFERROR(1/0, A1)", origin).expect("parse");
        let program = super::super::BytecodeCache::new().get_or_compile(&expr);
        let grid = CountingGrid::new(10, 10);
        let mut vm = Vm::with_capacity(32);
        let value = vm.eval(&program, &grid, 0, origin, &locale);
        assert_eq!(value, Value::Empty); // A1 is empty.
        assert_eq!(
            grid.reads(),
            1,
            "IFERROR fallback should be evaluated for errors"
        );

        // IFERROR(1/0, SUM(A1:A10)) should evaluate the fallback (range read).
        let expr =
            super::super::parse_formula("=IFERROR(1/0, SUM(A1:A10))", origin).expect("parse");
        let program = super::super::BytecodeCache::new().get_or_compile(&expr);
        let grid = CountingGrid::new(10, 10);
        let mut vm = Vm::with_capacity(32);
        let value = vm.eval(&program, &grid, 0, origin, &locale);
        assert_eq!(value, Value::Number(0.0));
        assert!(
            grid.reads() > 0,
            "IFERROR fallback should be evaluated for errors (including range reads)"
        );
        // IFNA(1, A1) should not evaluate the fallback.
        let expr = super::super::parse_formula("=IFNA(1, A1)", origin).expect("parse");
        let program = super::super::BytecodeCache::new().get_or_compile(&expr);
        let grid = CountingGrid::new(10, 10);
        let mut vm = Vm::with_capacity(32);
        let value = vm.eval(&program, &grid, 0, origin, &locale);
        assert_eq!(value, Value::Number(1.0));
        assert_eq!(
            grid.reads(),
            0,
            "unused IFNA fallback should not be evaluated"
        );

        // IFNA(1, SUM(A1:A10)) should not evaluate the fallback.
        let expr = super::super::parse_formula("=IFNA(1, SUM(A1:A10))", origin).expect("parse");
        let program = super::super::BytecodeCache::new().get_or_compile(&expr);
        let grid = CountingGrid::new(10, 10);
        let mut vm = Vm::with_capacity(32);
        let value = vm.eval(&program, &grid, 0, origin, &locale);
        assert_eq!(value, Value::Number(1.0));
        assert_eq!(
            grid.reads(),
            0,
            "unused IFNA fallback should not be evaluated (including range reads)"
        );

        // IFNA(1/0, A1) should not evaluate the fallback because the error is not #N/A.
        let expr = super::super::parse_formula("=IFNA(1/0, A1)", origin).expect("parse");
        let program = super::super::BytecodeCache::new().get_or_compile(&expr);
        let grid = CountingGrid::new(10, 10);
        let mut vm = Vm::with_capacity(32);
        let value = vm.eval(&program, &grid, 0, origin, &locale);
        assert_eq!(value, Value::Error(super::super::value::ErrorKind::Div0));
        assert_eq!(
            grid.reads(),
            0,
            "IFNA fallback should not be evaluated for non-#N/A errors"
        );

        // IFNA(1/0, SUM(A1:A10)) should not evaluate the fallback because the error is not #N/A.
        let expr = super::super::parse_formula("=IFNA(1/0, SUM(A1:A10))", origin).expect("parse");
        let program = super::super::BytecodeCache::new().get_or_compile(&expr);
        let grid = CountingGrid::new(10, 10);
        let mut vm = Vm::with_capacity(32);
        let value = vm.eval(&program, &grid, 0, origin, &locale);
        assert_eq!(value, Value::Error(super::super::value::ErrorKind::Div0));
        assert_eq!(
            grid.reads(),
            0,
            "IFNA fallback should not be evaluated for non-#N/A errors (including range reads)"
        );
        // IFNA(NA(), A1) should evaluate the fallback.
        let expr = super::super::parse_formula("=IFNA(NA(), A1)", origin).expect("parse");
        let program = super::super::BytecodeCache::new().get_or_compile(&expr);
        let grid = CountingGrid::new(10, 10);
        let mut vm = Vm::with_capacity(32);
        let value = vm.eval(&program, &grid, 0, origin, &locale);
        assert_eq!(value, Value::Empty); // A1 is empty.
        assert_eq!(
            grid.reads(),
            1,
            "IFNA fallback should be evaluated for #N/A"
        );

        // IFNA(NA(), SUM(A1:A10)) should evaluate the fallback (range read).
        let expr = super::super::parse_formula("=IFNA(NA(), SUM(A1:A10))", origin).expect("parse");
        let program = super::super::BytecodeCache::new().get_or_compile(&expr);
        let grid = CountingGrid::new(10, 10);
        let mut vm = Vm::with_capacity(32);
        let value = vm.eval(&program, &grid, 0, origin, &locale);
        assert_eq!(value, Value::Number(0.0));
        assert!(
            grid.reads() > 0,
            "IFNA fallback should be evaluated for #N/A (including range reads)"
        );
    }

    #[test]
    fn vm_short_circuits_ifs_pairs() {
        let origin = CellCoord::new(0, 0);
        let locale = crate::LocaleConfig::en_us();

        // IFS(TRUE, 1, A1, 2) should not evaluate later conditions.
        let expr = super::super::parse_formula("=IFS(TRUE, 1, A1, 2)", origin).expect("parse");
        let program = super::super::BytecodeCache::new().get_or_compile(&expr);
        let grid = CountingGrid::new(10, 10);
        let mut vm = Vm::with_capacity(32);
        let value = vm.eval(&program, &grid, 0, origin, &locale);
        assert_eq!(value, Value::Number(1.0));
        assert_eq!(
            grid.reads(),
            0,
            "unused IFS condition should not be evaluated"
        );

        // IFS(TRUE, 1, SUM(A1:A10), 2) should not evaluate later conditions, even if they would use
        // range reads.
        let expr =
            super::super::parse_formula("=IFS(TRUE, 1, SUM(A1:A10), 2)", origin).expect("parse");
        let program = super::super::BytecodeCache::new().get_or_compile(&expr);
        let grid = CountingGrid::new(10, 10);
        let mut vm = Vm::with_capacity(32);
        let value = vm.eval(&program, &grid, 0, origin, &locale);
        assert_eq!(value, Value::Number(1.0));
        assert_eq!(
            grid.reads(),
            0,
            "unused IFS condition should not be evaluated (including range reads)"
        );
        // IFS(FALSE, A1, TRUE, 2) should not evaluate the value for a FALSE condition.
        let expr = super::super::parse_formula("=IFS(FALSE, A1, TRUE, 2)", origin).expect("parse");
        let program = super::super::BytecodeCache::new().get_or_compile(&expr);
        let grid = CountingGrid::new(10, 10);
        let mut vm = Vm::with_capacity(32);
        let value = vm.eval(&program, &grid, 0, origin, &locale);
        assert_eq!(value, Value::Number(2.0));
        assert_eq!(
            grid.reads(),
            0,
            "unused IFS value expression should not be evaluated"
        );

        // IFS(FALSE, SUM(A1:A10), TRUE, 2) should not evaluate the value for a FALSE condition,
        // even if it would use range reads.
        let expr = super::super::parse_formula("=IFS(FALSE, SUM(A1:A10), TRUE, 2)", origin)
            .expect("parse");
        let program = super::super::BytecodeCache::new().get_or_compile(&expr);
        let grid = CountingGrid::new(10, 10);
        let mut vm = Vm::with_capacity(32);
        let value = vm.eval(&program, &grid, 0, origin, &locale);
        assert_eq!(value, Value::Number(2.0));
        assert_eq!(
            grid.reads(),
            0,
            "unused IFS value expression should not be evaluated (including range reads)"
        );
    }

    #[test]
    fn vm_short_circuits_switch_cases() {
        let origin = CellCoord::new(0, 0);
        let locale = crate::LocaleConfig::en_us();

        // SWITCH(1, 1, 10, A1, 20) should not evaluate later case values after a match.
        let expr = super::super::parse_formula("=SWITCH(1, 1, 10, A1, 20)", origin).expect("parse");
        let program = super::super::BytecodeCache::new().get_or_compile(&expr);
        let grid = CountingGrid::new(10, 10);
        let mut vm = Vm::with_capacity(32);
        let value = vm.eval(&program, &grid, 0, origin, &locale);
        assert_eq!(value, Value::Number(10.0));
        assert_eq!(
            grid.reads(),
            0,
            "unused SWITCH case value should not be evaluated"
        );

        // SWITCH(1, 1, 10, SUM(A1:A10), 20) should not evaluate later case values after a match,
        // even if they would use range reads.
        let expr = super::super::parse_formula("=SWITCH(1, 1, 10, SUM(A1:A10), 20)", origin)
            .expect("parse");
        let program = super::super::BytecodeCache::new().get_or_compile(&expr);
        let grid = CountingGrid::new(10, 10);
        let mut vm = Vm::with_capacity(32);
        let value = vm.eval(&program, &grid, 0, origin, &locale);
        assert_eq!(value, Value::Number(10.0));
        assert_eq!(
            grid.reads(),
            0,
            "unused SWITCH case value should not be evaluated (including range reads)"
        );
        // SWITCH(2, 1, A1, 2, 20) should not evaluate the result for a non-matching case.
        let expr = super::super::parse_formula("=SWITCH(2, 1, A1, 2, 20)", origin).expect("parse");
        let program = super::super::BytecodeCache::new().get_or_compile(&expr);
        let grid = CountingGrid::new(10, 10);
        let mut vm = Vm::with_capacity(32);
        let value = vm.eval(&program, &grid, 0, origin, &locale);
        assert_eq!(value, Value::Number(20.0));
        assert_eq!(
            grid.reads(),
            0,
            "unused SWITCH result expression should not be evaluated"
        );

        // SWITCH(2, 1, SUM(A1:A10), 2, 20) should not evaluate the result for a non-matching case,
        // even if it would use range reads.
        let expr = super::super::parse_formula("=SWITCH(2, 1, SUM(A1:A10), 2, 20)", origin)
            .expect("parse");
        let program = super::super::BytecodeCache::new().get_or_compile(&expr);
        let grid = CountingGrid::new(10, 10);
        let mut vm = Vm::with_capacity(32);
        let value = vm.eval(&program, &grid, 0, origin, &locale);
        assert_eq!(value, Value::Number(20.0));
        assert_eq!(
            grid.reads(),
            0,
            "unused SWITCH result expression should not be evaluated (including range reads)"
        );
        // If the discriminant expression is an error, SWITCH should not evaluate any case values.
        let expr =
            super::super::parse_formula("=SWITCH(1/0, 1, 10, A1, 20)", origin).expect("parse");
        let program = super::super::BytecodeCache::new().get_or_compile(&expr);
        let grid = CountingGrid::new(10, 10);
        let mut vm = Vm::with_capacity(32);
        let value = vm.eval(&program, &grid, 0, origin, &locale);
        assert_eq!(value, Value::Error(super::super::value::ErrorKind::Div0));
        assert_eq!(
            grid.reads(),
            0,
            "SWITCH should not evaluate case values when discriminant is an error"
        );

        // If the discriminant expression is an error, SWITCH should not evaluate any case values,
        // even if they would use range reads.
        let expr = super::super::parse_formula("=SWITCH(1/0, 1, 10, SUM(A1:A10), 20)", origin)
            .expect("parse");
        let program = super::super::BytecodeCache::new().get_or_compile(&expr);
        let grid = CountingGrid::new(10, 10);
        let mut vm = Vm::with_capacity(32);
        let value = vm.eval(&program, &grid, 0, origin, &locale);
        assert_eq!(value, Value::Error(super::super::value::ErrorKind::Div0));
        assert_eq!(
            grid.reads(),
            0,
            "SWITCH should not evaluate case values when discriminant is an error (including range reads)"
        );
    }

    #[test]
    fn vm_short_circuits_choose_branches() {
        let origin = CellCoord::new(0, 0);
        let locale = crate::LocaleConfig::en_us();

        let expr = super::super::parse_formula("=CHOOSE(2, 1/0, 7)", origin).expect("parse");
        let program = super::super::BytecodeCache::new().get_or_compile(&expr);
        let grid = super::super::grid::ColumnarGrid::new(1, 1);

        let mut vm = Vm::with_capacity(32);
        let value = vm.eval(&program, &grid, 0, origin, &locale);
        assert_eq!(value, Value::Number(7.0));
    }

    #[test]
    fn vm_choose_index_error_propagates_without_evaluating_choices() {
        let origin = CellCoord::new(0, 0);
        let expr = super::super::parse_formula("=CHOOSE(1/0, A1, 7)", origin).expect("parse");
        let program = super::super::BytecodeCache::new().get_or_compile(&expr);
        let grid = CountingGrid::new(10, 10);
        let locale = crate::LocaleConfig::en_us();

        let mut vm = Vm::with_capacity(32);
        let value = vm.eval(&program, &grid, 0, origin, &locale);
        assert_eq!(value, Value::Error(super::super::value::ErrorKind::Div0));
        assert_eq!(
            grid.reads(),
            0,
            "CHOOSE should not evaluate choices when index is an error"
        );
    }

    #[test]
    fn vm_choose_coerces_fractional_index() {
        let origin = CellCoord::new(0, 0);
        // Excel truncates CHOOSE indices to integers: 1.9 selects the first value.
        let expr = super::super::parse_formula("=CHOOSE(1.9, 7, 8)", origin).expect("parse");
        let program = super::super::BytecodeCache::new().get_or_compile(&expr);
        let grid = super::super::grid::ColumnarGrid::new(1, 1);
        let locale = crate::LocaleConfig::en_us();

        let mut vm = Vm::with_capacity(32);
        let value = vm.eval(&program, &grid, 0, origin, &locale);
        assert_eq!(value, Value::Number(7.0));
    }

    #[test]
    fn vm_choose_truncates_fractional_index_without_evaluating_other_branches() {
        let origin = CellCoord::new(0, 0);
        // Index 2.9 should truncate to 2; the other branches must not be evaluated.
        let expr =
            super::super::parse_formula("=CHOOSE(2.9, 1/0, 20, 1/0)", origin).expect("parse");
        let program = super::super::BytecodeCache::new().get_or_compile(&expr);
        let grid = super::super::grid::ColumnarGrid::new(1, 1);
        let locale = crate::LocaleConfig::en_us();

        let mut vm = Vm::with_capacity(32);
        let value = vm.eval(&program, &grid, 0, origin, &locale);
        assert_eq!(value, Value::Number(20.0));
    }

    #[test]
    fn vm_choose_out_of_range_returns_value_error_without_evaluating_choices() {
        let origin = CellCoord::new(0, 0);
        let locale = crate::LocaleConfig::en_us();

        // Index 3 is out of range; the choice expressions should not be evaluated.
        let expr = super::super::parse_formula("=CHOOSE(3, 1/0, 7)", origin).expect("parse");
        let program = super::super::BytecodeCache::new().get_or_compile(&expr);
        let grid = super::super::grid::ColumnarGrid::new(1, 1);

        let mut vm = Vm::with_capacity(32);
        let value = vm.eval(&program, &grid, 0, origin, &locale);
        assert_eq!(value, Value::Error(super::super::ErrorKind::Value));
    }

    #[test]
    fn vm_choose_nan_index_returns_value_error_without_evaluating_choices() {
        let origin = CellCoord::new(0, 0);
        let expr = super::super::parse_formula("=CHOOSE(A1, A2, 7)", origin).expect("parse");
        let program = super::super::BytecodeCache::new().get_or_compile(&expr);
        let grid = NanIndexGrid::new();
        let locale = crate::LocaleConfig::en_us();

        let mut vm = Vm::with_capacity(32);
        let value = vm.eval(&program, &grid, 0, origin, &locale);
        assert_eq!(value, Value::Error(super::super::ErrorKind::Value));
        assert_eq!(grid.reads(), 1, "CHOOSE should only evaluate the index");
    }

    #[test]
    fn vm_choose_truncates_non_integer_index_toward_zero() {
        let origin = CellCoord::new(0, 0);
        let locale = crate::LocaleConfig::en_us();
        let grid = super::super::grid::ColumnarGrid::new(1, 1);

        // Excel truncates CHOOSE's index (toward zero).
        let expr = super::super::parse_formula("=CHOOSE(1.9, 10, 20)", origin).expect("parse");
        let program = super::super::BytecodeCache::new().get_or_compile(&expr);
        let mut vm = Vm::with_capacity(32);
        let value = vm.eval(&program, &grid, 0, origin, &locale);
        assert_eq!(value, Value::Number(10.0));

        // After truncation this becomes 0, which is out of range.
        let expr = super::super::parse_formula("=CHOOSE(0.9, 10, 20)", origin).expect("parse");
        let program = super::super::BytecodeCache::new().get_or_compile(&expr);
        let mut vm = Vm::with_capacity(32);
        let value = vm.eval(&program, &grid, 0, origin, &locale);
        assert_eq!(value, Value::Error(super::super::ErrorKind::Value));
    }

    #[test]
    fn vm_choose_propagates_index_error_without_evaluating_choices() {
        let origin = CellCoord::new(0, 0);
        let locale = crate::LocaleConfig::en_us();

        // If the index expression is an error, CHOOSE should not evaluate any choice branches.
        let expr = super::super::parse_formula("=CHOOSE(1/0, A1, 7)", origin).expect("parse");
        let program = super::super::BytecodeCache::new().get_or_compile(&expr);

        let grid = CountingGrid::new(10, 10);
        let mut vm = Vm::with_capacity(32);
        let value = vm.eval(&program, &grid, 0, origin, &locale);

        assert_eq!(value, Value::Error(super::super::value::ErrorKind::Div0));
        assert_eq!(
            grid.reads(),
            0,
            "CHOOSE branches should not be evaluated when index is an error"
        );
    }

    #[test]
    fn vm_choose_parses_text_index_and_is_lazy() {
        let origin = CellCoord::new(0, 0);
        let locale = crate::LocaleConfig::en_us();

        // Excel parses a numeric index from text.
        // Ensure we still short-circuit the unselected branch.
        let expr = super::super::parse_formula("=CHOOSE(\"2\", A1, 7)", origin).expect("parse");
        let program = super::super::BytecodeCache::new().get_or_compile(&expr);
        let grid = CountingGrid::new(10, 10);

        let mut vm = Vm::with_capacity(32);
        let value = vm.eval(&program, &grid, 0, origin, &locale);

        assert_eq!(value, Value::Number(7.0));
        assert_eq!(
            grid.reads(),
            0,
            "unused CHOOSE branch should not be evaluated"
        );
    }

    #[test]
    fn vm_choose_range_index_returns_spill_without_evaluating_choices() {
        let origin = CellCoord::new(0, 0);
        let locale = crate::LocaleConfig::en_us();

        // Bytecode coercion treats ranges/arrays in scalar contexts as a spill attempt.
        // CHOOSE should surface that error without evaluating any branch expressions.
        let expr = super::super::parse_formula("=CHOOSE(A1:A2, A1, 7)", origin).expect("parse");
        let program = super::super::BytecodeCache::new().get_or_compile(&expr);
        let grid = CountingGrid::new(10, 10);

        let mut vm = Vm::with_capacity(32);
        let value = vm.eval(&program, &grid, 0, origin, &locale);

        assert_eq!(value, Value::Error(super::super::ErrorKind::Spill));
        assert_eq!(
            grid.reads(),
            0,
            "CHOOSE branches should not be evaluated when index is a spill error"
        );
    }

    #[test]
    fn vm_short_circuits_ifs_branches() {
        let origin = CellCoord::new(0, 0);
        let locale = crate::LocaleConfig::en_us();

        let expr = super::super::parse_formula("=IFS(FALSE, 1/0, TRUE, 9)", origin).expect("parse");
        let program = super::super::BytecodeCache::new().get_or_compile(&expr);
        let grid = super::super::grid::ColumnarGrid::new(1, 1);

        let mut vm = Vm::with_capacity(32);
        let value = vm.eval(&program, &grid, 0, origin, &locale);
        assert_eq!(value, Value::Number(9.0));
    }

    #[test]
    fn vm_ifs_missing_match_returns_na_without_evaluating_values() {
        let origin = CellCoord::new(0, 0);
        let locale = crate::LocaleConfig::en_us();

        let expr = super::super::parse_formula("=IFS(FALSE, 1/0)", origin).expect("parse");
        let program = super::super::BytecodeCache::new().get_or_compile(&expr);
        let grid = super::super::grid::ColumnarGrid::new(1, 1);

        let mut vm = Vm::with_capacity(32);
        let value = vm.eval(&program, &grid, 0, origin, &locale);
        assert_eq!(value, Value::Error(super::super::ErrorKind::NA));
    }

    #[test]
    fn vm_short_circuits_switch_branches() {
        let origin = CellCoord::new(0, 0);
        let locale = crate::LocaleConfig::en_us();

        let expr = super::super::parse_formula("=SWITCH(2, 1, 1/0, 2, 8)", origin).expect("parse");
        let program = super::super::BytecodeCache::new().get_or_compile(&expr);
        let grid = super::super::grid::ColumnarGrid::new(1, 1);

        let mut vm = Vm::with_capacity(32);
        let value = vm.eval(&program, &grid, 0, origin, &locale);
        assert_eq!(value, Value::Number(8.0));
    }

    #[test]
    fn vm_does_not_evaluate_unselected_choose_branch() {
        let origin = CellCoord::new(0, 0);
        let locale = crate::LocaleConfig::en_us();

        // CHOOSE(1, A1, B1) should not evaluate the non-selected branch.
        let expr = super::super::parse_formula("=CHOOSE(1, A1, B1)", origin).expect("parse");
        let program = super::super::BytecodeCache::new().get_or_compile(&expr);

        let grid = CountingGrid::new(10, 10);
        let mut vm = Vm::with_capacity(32);
        let value = vm.eval(&program, &grid, 0, origin, &locale);

        assert_eq!(value, Value::Empty); // A1 is empty.
        assert_eq!(
            grid.reads(),
            1,
            "unused CHOOSE branch should not be evaluated"
        );

        // CHOOSE(2, A1, 1) should not evaluate the first branch.
        let expr = super::super::parse_formula("=CHOOSE(2, A1, 1)", origin).expect("parse");
        let program = super::super::BytecodeCache::new().get_or_compile(&expr);

        let grid = CountingGrid::new(10, 10);
        let mut vm = Vm::with_capacity(32);
        let value = vm.eval(&program, &grid, 0, origin, &locale);

        assert_eq!(value, Value::Number(1.0));
        assert_eq!(
            grid.reads(),
            0,
            "unused CHOOSE branch should not be evaluated"
        );
    }
}
