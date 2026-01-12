use super::ast::{BinaryOp, UnaryOp};
use super::grid::Grid;
use super::runtime::{apply_binary, apply_implicit_intersection, apply_unary, call_function};
use super::value::{CellCoord, Value};
use crate::date::ExcelDateSystem;
use crate::locale::ValueLocaleConfig;
use chrono::{DateTime, Utc};

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
        locale: &crate::LocaleConfig,
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
                OpCode::ImplicitIntersection => {
                    let v = self.stack.pop().unwrap_or(Value::Empty);
                    self.stack
                        .push(apply_implicit_intersection(v, grid, base));
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
                    let result = call_function(func, &self.stack[start..], grid, base, locale);
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
        value_locale: ValueLocaleConfig,
    ) -> Value {
        // Preserve the existing public API while ensuring locale-aware coercion for text values
        // matches the main evaluator. This uses Excel's default 1900 date system and the current
        // wall-clock time for any date strings that omit a year.
        self.eval_with_coercion_context(
            program,
            grid,
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
        base: CellCoord,
        date_system: ExcelDateSystem,
        value_locale: ValueLocaleConfig,
        now_utc: DateTime<Utc>,
    ) -> Value {
        let _guard = super::runtime::set_thread_eval_context(date_system, value_locale, now_utc);

        // Criteria strings inside quotes should follow the workbook/value locale for numeric parsing.
        let mut locale_config = crate::LocaleConfig::en_us();
        locale_config.decimal_separator = value_locale.separators.decimal_sep;
        locale_config.thousands_separator = Some(value_locale.separators.thousands_sep);
        self.eval(program, grid, base, &locale_config)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::date::{ymd_to_serial, ExcelDate};
    use chrono::TimeZone;

    #[test]
    fn eval_with_value_locale_parses_numeric_text_using_locale() {
        let origin = CellCoord::new(0, 0);
        let expr = super::super::parse_formula("=\"1.234,56\"+1", origin).expect("parse");
        let cache = super::super::BytecodeCache::new();
        let program = cache.get_or_compile(&expr);
        let grid = super::super::grid::ColumnarGrid::new(1, 1);

        let mut vm = Vm::with_capacity(32);
        let value = vm.eval_with_value_locale(
            &program,
            &grid,
            origin,
            ValueLocaleConfig::de_de(),
        );

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
        let expected = ymd_to_serial(ExcelDate::new(2024, 1, 2), ExcelDateSystem::EXCEL_1900)
            .unwrap() as f64;

        let mut vm = Vm::with_capacity(32);
        let value = vm.eval_with_coercion_context(
            &program,
            &grid,
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
}
