use super::ast::{BinaryOp, Expr, Function, UnaryOp};
use super::grid::Grid;
use super::value::{Array as ArrayValue, CellCoord, ErrorKind, RangeRef, ResolvedRange, Value};
use crate::error::ExcelError;
use crate::value::parse_number;
use crate::simd::{self, CmpOp, NumericCriteria};
use smallvec::SmallVec;
use std::cmp::Ordering;
use std::sync::Arc;

pub fn eval_ast(expr: &Expr, grid: &dyn Grid, base: CellCoord) -> Value {
    match expr {
        Expr::Literal(v) => v.clone(),
        Expr::CellRef(r) => grid.get_value(r.resolve(base)),
        Expr::RangeRef(r) => Value::Range(*r),
        Expr::Unary { op, expr } => {
            let v = eval_ast(expr, grid, base);
            apply_unary(*op, v)
        }
        Expr::Binary { op, left, right } => {
            let l = eval_ast(left, grid, base);
            let r = eval_ast(right, grid, base);
            apply_binary(*op, l, r)
        }
        Expr::FuncCall { func, args } => {
            // Evaluate arguments first (AST evaluation).
            let mut evaluated: SmallVec<[Value; 8]> = SmallVec::with_capacity(args.len());
            for (arg_idx, arg) in args.iter().enumerate() {
                let treat_cell_as_range = match func {
                    // See `Compiler::compile_func_arg` for the rationale.
                    Function::Sum
                    | Function::Average
                    | Function::Min
                    | Function::Max
                    | Function::Count => true,
                    Function::CountIf => arg_idx == 0,
                    Function::SumProduct => true,
                    Function::Abs
                    | Function::Int
                    | Function::Round
                    | Function::RoundUp
                    | Function::RoundDown
                    | Function::Mod
                    | Function::Sign
                    | Function::Concat
                    | Function::Unknown(_) => false,
                };

                if treat_cell_as_range {
                    if let Expr::CellRef(r) = arg {
                        evaluated.push(Value::Range(RangeRef::new(*r, *r)));
                        continue;
                    }
                }

                evaluated.push(eval_ast(arg, grid, base));
            }
            call_function(func, &evaluated, grid, base)
        }
    }
}

fn parse_number_from_text(s: &str) -> Result<f64, ErrorKind> {
    parse_number(s, crate::value::NumberLocale::en_us()).map_err(|e| match e {
        ExcelError::Div0 => ErrorKind::Value,
        ExcelError::Value => ErrorKind::Value,
        ExcelError::Num => ErrorKind::Num,
    })
}

fn coerce_to_number(v: Value) -> Result<f64, ErrorKind> {
    match v {
        Value::Number(n) => Ok(n),
        Value::Bool(b) => Ok(if b { 1.0 } else { 0.0 }),
        Value::Empty => Ok(0.0),
        Value::Text(s) => parse_number_from_text(&s),
        Value::Error(e) => Err(e),
        // Dynamic arrays / range-as-scalar: treat as a spill attempt (engine semantics).
        Value::Array(_) | Value::Range(_) => Err(ErrorKind::Spill),
    }
}

fn matches_numeric_criteria(v: f64, criteria: NumericCriteria) -> bool {
    match criteria.op {
        CmpOp::Eq => v == criteria.rhs,
        CmpOp::Ne => v != criteria.rhs,
        CmpOp::Lt => v < criteria.rhs,
        CmpOp::Le => v <= criteria.rhs,
        CmpOp::Gt => v > criteria.rhs,
        CmpOp::Ge => v >= criteria.rhs,
    }
}

fn count_if_f64_blank_as_zero(values: &[f64], criteria: NumericCriteria) -> usize {
    // COUNTIF treats blank cells as zero for numeric criteria. Column slices represent blanks as
    // NaN, so normalize before comparison.
    values
        .iter()
        .filter(|v| {
            let v = if v.is_nan() { 0.0 } else { **v };
            matches_numeric_criteria(v, criteria)
        })
        .count()
}

fn coerce_countif_value_to_number(v: Value) -> Option<f64> {
    match v {
        Value::Number(n) => Some(n),
        Value::Bool(b) => Some(if b { 1.0 } else { 0.0 }),
        Value::Empty => Some(0.0),
        Value::Text(s) => parse_number_from_text(&s).ok(),
        Value::Error(_) | Value::Array(_) | Value::Range(_) => None,
    }
}

pub fn apply_unary(op: UnaryOp, v: Value) -> Value {
    let n = match coerce_to_number(v) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    match op {
        UnaryOp::Plus => Value::Number(n),
        UnaryOp::Neg => Value::Number(-n),
    }
}

pub fn apply_binary(op: BinaryOp, left: Value, right: Value) -> Value {
    use Value::*;

    match op {
        BinaryOp::Add => numeric_binop(left, right, |a, b| a + b, simd::add_f64),
        BinaryOp::Sub => numeric_binop(left, right, |a, b| a - b, simd::sub_f64),
        BinaryOp::Mul => numeric_binop(left, right, |a, b| a * b, simd::mul_f64),
        BinaryOp::Div => match (left, right) {
            (Error(e), _) | (_, Error(e)) => Error(e),
            (Array(a), Array(b)) => {
                if a.rows != b.rows || a.cols != b.cols {
                    return Error(ErrorKind::Value);
                }
                let mut out = vec![0.0; a.values.len()];
                simd::div_f64(&mut out, &a.values, &b.values);
                Value::Array(ArrayValue::new(a.rows, a.cols, out))
            }
            (Array(a), other) => {
                let denom = match coerce_to_number(other) {
                    Ok(n) => n,
                    Err(e) => return Error(e),
                };
                if denom == 0.0 {
                    return Error(ErrorKind::Div0);
                }
                let mut out = a.values.clone();
                for v in &mut out {
                    *v /= denom;
                }
                Value::Array(ArrayValue::new(a.rows, a.cols, out))
            }
            (other, Array(b)) => {
                let numer = match coerce_to_number(other) {
                    Ok(n) => n,
                    Err(e) => return Error(e),
                };
                let mut out = b.values.clone();
                for v in &mut out {
                    *v = numer / *v;
                }
                Value::Array(ArrayValue::new(b.rows, b.cols, out))
            }
            (l, r) => {
                let ln = match coerce_to_number(l) {
                    Ok(n) => n,
                    Err(e) => return Error(e),
                };
                let rn = match coerce_to_number(r) {
                    Ok(n) => n,
                    Err(e) => return Error(e),
                };
                if rn == 0.0 {
                    Error(ErrorKind::Div0)
                } else {
                    Number(ln / rn)
                }
            }
        },
        BinaryOp::Pow => {
            let a = match coerce_to_number(left) {
                Ok(n) => n,
                Err(e) => return Error(e),
            };
            let b = match coerce_to_number(right) {
                Ok(n) => n,
                Err(e) => return Error(e),
            };
            match crate::functions::math::power(a, b) {
                Ok(n) => Number(n),
                Err(e) => Error(match e {
                    ExcelError::Div0 => ErrorKind::Div0,
                    ExcelError::Value => ErrorKind::Value,
                    ExcelError::Num => ErrorKind::Num,
                }),
            }
        }
        BinaryOp::Eq | BinaryOp::Ne | BinaryOp::Lt | BinaryOp::Le | BinaryOp::Gt | BinaryOp::Ge => {
            excel_compare(left, right, op)
        }
    }
}

fn excel_compare(left: Value, right: Value, op: BinaryOp) -> Value {
    let ord = match excel_order(left, right) {
        Ok(ord) => ord,
        Err(e) => return Value::Error(e),
    };

    let result = match op {
        BinaryOp::Eq => ord == Ordering::Equal,
        BinaryOp::Ne => ord != Ordering::Equal,
        BinaryOp::Lt => ord == Ordering::Less,
        BinaryOp::Le => ord != Ordering::Greater,
        BinaryOp::Gt => ord == Ordering::Greater,
        BinaryOp::Ge => ord != Ordering::Less,
        _ => return Value::Error(ErrorKind::Value),
    };

    Value::Bool(result)
}

fn excel_order(left: Value, right: Value) -> Result<Ordering, ErrorKind> {
    if let Value::Error(e) = left {
        return Err(e);
    }
    if let Value::Error(e) = right {
        return Err(e);
    }
    if matches!(left, Value::Array(_) | Value::Range(_))
        || matches!(right, Value::Array(_) | Value::Range(_))
    {
        return Err(ErrorKind::Value);
    }

    // Blank coerces to the other type for comparisons.
    let (l, r) = match (&left, &right) {
        (Value::Empty, Value::Number(_)) => (Value::Number(0.0), right),
        (Value::Number(_), Value::Empty) => (left, Value::Number(0.0)),
        (Value::Empty, Value::Bool(_)) => (Value::Bool(false), right),
        (Value::Bool(_), Value::Empty) => (left, Value::Bool(false)),
        (Value::Empty, Value::Text(_)) => (Value::Text(Arc::from("")), right),
        (Value::Text(_), Value::Empty) => (left, Value::Text(Arc::from(""))),
        _ => (left, right),
    };

    Ok(match (l, r) {
        (Value::Number(a), Value::Number(b)) => a.partial_cmp(&b).unwrap_or(Ordering::Equal),
        (Value::Text(a), Value::Text(b)) => {
            let au = a.to_ascii_uppercase();
            let bu = b.to_ascii_uppercase();
            au.cmp(&bu)
        }
        (Value::Bool(a), Value::Bool(b)) => a.cmp(&b),
        // Type precedence (approximate Excel): numbers < text < booleans.
        (Value::Number(_), Value::Text(_) | Value::Bool(_)) => Ordering::Less,
        (Value::Text(_), Value::Bool(_)) => Ordering::Less,
        (Value::Text(_), Value::Number(_)) => Ordering::Greater,
        (Value::Bool(_), Value::Number(_) | Value::Text(_)) => Ordering::Greater,
        // Blank should have been coerced above.
        (Value::Empty, Value::Empty) => Ordering::Equal,
        (Value::Empty, _) => Ordering::Less,
        (_, Value::Empty) => Ordering::Greater,
        // Errors are handled above.
        (Value::Error(_), _) | (_, Value::Error(_)) => Ordering::Equal,
        // Arrays/ranges are rejected above.
        (Value::Array(_), _)
        | (_, Value::Array(_))
        | (Value::Range(_), _)
        | (_, Value::Range(_)) => Ordering::Equal,
    })
}

fn numeric_binop(
    left: Value,
    right: Value,
    scalar: fn(f64, f64) -> f64,
    simd_binop: fn(&mut [f64], &[f64], &[f64]),
) -> Value {
    use Value::*;
    match (left, right) {
        (Error(e), _) | (_, Error(e)) => Error(e),
        (Array(a), Array(b)) => {
            if a.rows != b.rows || a.cols != b.cols {
                return Error(ErrorKind::Value);
            }
            let mut out = vec![0.0; a.values.len()];
            simd_binop(&mut out, &a.values, &b.values);
            Value::Array(ArrayValue::new(a.rows, a.cols, out))
        }
        (Array(a), other) => {
            let b = match coerce_to_number(other) {
                Ok(n) => n,
                Err(e) => return Error(e),
            };
            let mut out = a.values.clone();
            for v in &mut out {
                *v = scalar(*v, b);
            }
            Value::Array(ArrayValue::new(a.rows, a.cols, out))
        }
        (other, Array(b)) => {
            let a = match coerce_to_number(other) {
                Ok(n) => n,
                Err(e) => return Error(e),
            };
            let mut out = b.values.clone();
            for v in &mut out {
                *v = scalar(a, *v);
            }
            Value::Array(ArrayValue::new(b.rows, b.cols, out))
        }
        (l, r) => match (coerce_to_number(l), coerce_to_number(r)) {
            (Ok(a), Ok(b)) => Number(scalar(a, b)),
            (Err(e), _) | (_, Err(e)) => Error(e),
        },
    }
}

pub fn call_function(func: &Function, args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    match func {
        Function::Sum => fn_sum(args, grid, base),
        Function::Average => fn_average(args, grid, base),
        Function::Min => fn_min(args, grid, base),
        Function::Max => fn_max(args, grid, base),
        Function::Count => fn_count(args, grid, base),
        Function::CountIf => fn_countif(args, grid, base),
        Function::SumProduct => fn_sumproduct(args, grid, base),
        Function::Abs => fn_abs(args),
        Function::Int => fn_int(args),
        Function::Round => fn_round(args),
        Function::RoundUp => fn_roundup(args),
        Function::RoundDown => fn_rounddown(args),
        Function::Mod => fn_mod(args),
        Function::Sign => fn_sign(args),
        Function::Concat => fn_concat(args),
        Function::Unknown(_) => Value::Error(ErrorKind::Name),
    }
}

fn fn_abs(args: &[Value]) -> Value {
    if args.len() != 1 {
        return Value::Error(ErrorKind::Value);
    }
    match coerce_to_number(args[0].clone()) {
        Ok(n) => Value::Number(n.abs()),
        Err(e) => Value::Error(e),
    }
}

fn fn_int(args: &[Value]) -> Value {
    if args.len() != 1 {
        return Value::Error(ErrorKind::Value);
    }
    match coerce_to_number(args[0].clone()) {
        Ok(n) => Value::Number(n.floor()),
        Err(e) => Value::Error(e),
    }
}

fn coerce_to_i64(v: Value) -> Result<i64, ErrorKind> {
    let n = coerce_to_number(v)?;
    Ok(n.trunc() as i64)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum RoundMode {
    Nearest,
    Down,
    Up,
}

fn round_with_mode(n: f64, digits: i32, mode: RoundMode) -> f64 {
    let factor = 10f64.powi(digits.saturating_abs());
    if !factor.is_finite() || factor == 0.0 {
        return n;
    }

    let scaled = if digits >= 0 { n * factor } else { n / factor };
    let rounded = match mode {
        RoundMode::Down => scaled.trunc(),
        RoundMode::Up => {
            if scaled.is_sign_negative() {
                scaled.trunc() - if scaled.fract() == 0.0 { 0.0 } else { 1.0 }
            } else {
                scaled.trunc() + if scaled.fract() == 0.0 { 0.0 } else { 1.0 }
            }
        }
        RoundMode::Nearest => {
            // Excel rounds halves away from zero.
            let frac = scaled.fract().abs();
            let base = scaled.trunc();
            if frac < 0.5 {
                base
            } else {
                base + scaled.signum()
            }
        }
    };

    if digits >= 0 {
        rounded / factor
    } else {
        rounded * factor
    }
}

fn fn_round_impl(args: &[Value], mode: RoundMode) -> Value {
    if args.len() != 2 {
        return Value::Error(ErrorKind::Value);
    }
    let number = match coerce_to_number(args[0].clone()) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let digits = match coerce_to_i64(args[1].clone()) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    Value::Number(round_with_mode(number, digits as i32, mode))
}

fn fn_round(args: &[Value]) -> Value {
    fn_round_impl(args, RoundMode::Nearest)
}

fn fn_roundup(args: &[Value]) -> Value {
    fn_round_impl(args, RoundMode::Up)
}

fn fn_rounddown(args: &[Value]) -> Value {
    fn_round_impl(args, RoundMode::Down)
}

fn fn_mod(args: &[Value]) -> Value {
    if args.len() != 2 {
        return Value::Error(ErrorKind::Value);
    }
    let n = match coerce_to_number(args[0].clone()) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let d = match coerce_to_number(args[1].clone()) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    if d == 0.0 {
        return Value::Error(ErrorKind::Div0);
    }
    Value::Number(n - d * (n / d).floor())
}

fn fn_sign(args: &[Value]) -> Value {
    if args.len() != 1 {
        return Value::Error(ErrorKind::Value);
    }
    let number = match coerce_to_number(args[0].clone()) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    if !number.is_finite() {
        return Value::Error(ErrorKind::Num);
    }
    if number > 0.0 {
        Value::Number(1.0)
    } else if number < 0.0 {
        Value::Number(-1.0)
    } else {
        Value::Number(0.0)
    }
}

fn format_number_general(n: f64) -> String {
    if n == 0.0 {
        return "0".to_string();
    }
    if n.fract() == 0.0 {
        return format!("{:.0}", n);
    }
    let s = n.to_string();
    if s == "-0" || s == "-0.0" {
        "0".to_string()
    } else {
        s
    }
}

fn coerce_to_string(v: Value) -> Result<String, ErrorKind> {
    match v {
        Value::Text(s) => Ok(s.to_string()),
        Value::Number(n) => Ok(format_number_general(n)),
        Value::Bool(b) => Ok(if b { "TRUE" } else { "FALSE" }.to_string()),
        Value::Empty => Ok(String::new()),
        Value::Error(e) => Err(e),
        Value::Array(_) | Value::Range(_) => Err(ErrorKind::Value),
    }
}

fn fn_concat(args: &[Value]) -> Value {
    if args.is_empty() {
        return Value::Error(ErrorKind::Value);
    }
    let mut out = String::new();
    for arg in args {
        match coerce_to_string(arg.clone()) {
            Ok(s) => out.push_str(&s),
            Err(e) => return Value::Error(e),
        }
    }
    Value::Text(out.into())
}

fn fn_sum(args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    let mut sum = 0.0;
    for arg in args {
        match arg {
            Value::Number(v) => sum += v,
            Value::Bool(v) => sum += if *v { 1.0 } else { 0.0 },
            Value::Array(a) => sum += simd::sum_ignore_nan_f64(&a.values),
            Value::Range(r) => match sum_range(grid, r.resolve(base)) {
                Ok(v) => sum += v,
                Err(e) => return Value::Error(e),
            },
            Value::Empty => {}
            Value::Error(e) => return Value::Error(*e),
            Value::Text(s) => match parse_number_from_text(s) {
                Ok(v) => sum += v,
                Err(e) => return Value::Error(e),
            },
        }
    }
    Value::Number(sum)
}

fn fn_average(args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    if args.is_empty() {
        return Value::Error(ErrorKind::Value);
    }
    let mut sum = 0.0;
    let mut count = 0usize;
    for arg in args {
        match arg {
            Value::Number(v) => {
                sum += v;
                count += 1;
            }
            Value::Bool(v) => {
                sum += if *v { 1.0 } else { 0.0 };
                count += 1;
            }
            Value::Array(a) => {
                let (s, c) = simd::sum_count_ignore_nan_f64(&a.values);
                sum += s;
                count += c;
            }
            Value::Range(r) => match sum_count_range(grid, r.resolve(base)) {
                Ok((s, c)) => {
                    sum += s;
                    count += c;
                }
                Err(e) => return Value::Error(e),
            },
            Value::Empty => {}
            Value::Error(e) => return Value::Error(*e),
            Value::Text(s) => match parse_number_from_text(s) {
                Ok(v) => {
                    sum += v;
                    count += 1;
                }
                Err(e) => return Value::Error(e),
            },
        }
    }
    if count == 0 {
        return Value::Error(ErrorKind::Div0);
    }
    Value::Number(sum / count as f64)
}

fn fn_min(args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    if args.is_empty() {
        return Value::Error(ErrorKind::Value);
    }
    let mut out: Option<f64> = None;
    for arg in args {
        match arg {
            Value::Number(v) => out = Some(out.map_or(*v, |prev| prev.min(*v))),
            Value::Bool(v) => {
                out = Some(out.map_or(if *v { 1.0 } else { 0.0 }, |prev| {
                    prev.min(if *v { 1.0 } else { 0.0 })
                }))
            }
            Value::Array(a) => {
                if let Some(m) = simd::min_ignore_nan_f64(&a.values) {
                    out = Some(out.map_or(m, |prev| prev.min(m)));
                }
            }
            Value::Range(r) => match min_range(grid, r.resolve(base)) {
                Ok(Some(m)) => out = Some(out.map_or(m, |prev| prev.min(m))),
                Ok(None) => {}
                Err(e) => return Value::Error(e),
            },
            Value::Empty => out = Some(out.map_or(0.0, |prev| prev.min(0.0))),
            Value::Error(e) => return Value::Error(*e),
            Value::Text(s) => match parse_number_from_text(s) {
                Ok(v) => out = Some(out.map_or(v, |prev| prev.min(v))),
                Err(e) => return Value::Error(e),
            },
        }
    }
    Value::Number(out.unwrap_or(0.0))
}

fn fn_max(args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    if args.is_empty() {
        return Value::Error(ErrorKind::Value);
    }
    let mut out: Option<f64> = None;
    for arg in args {
        match arg {
            Value::Number(v) => out = Some(out.map_or(*v, |prev| prev.max(*v))),
            Value::Bool(v) => {
                out = Some(out.map_or(if *v { 1.0 } else { 0.0 }, |prev| {
                    prev.max(if *v { 1.0 } else { 0.0 })
                }))
            }
            Value::Array(a) => {
                if let Some(m) = simd::max_ignore_nan_f64(&a.values) {
                    out = Some(out.map_or(m, |prev| prev.max(m)));
                }
            }
            Value::Range(r) => match max_range(grid, r.resolve(base)) {
                Ok(Some(m)) => out = Some(out.map_or(m, |prev| prev.max(m))),
                Ok(None) => {}
                Err(e) => return Value::Error(e),
            },
            Value::Empty => out = Some(out.map_or(0.0, |prev| prev.max(0.0))),
            Value::Error(e) => return Value::Error(*e),
            Value::Text(s) => match parse_number_from_text(s) {
                Ok(v) => out = Some(out.map_or(v, |prev| prev.max(v))),
                Err(e) => return Value::Error(e),
            },
        }
    }
    Value::Number(out.unwrap_or(0.0))
}

fn fn_count(args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    let mut count = 0usize;
    for arg in args {
        match arg {
            Value::Number(_) => count += 1,
            Value::Array(a) => count += simd::count_ignore_nan_f64(&a.values),
            Value::Range(r) => match count_range(grid, r.resolve(base)) {
                Ok(c) => count += c,
                Err(e) => return Value::Error(e),
            },
            Value::Bool(_) | Value::Empty | Value::Error(_) | Value::Text(_) => {}
        }
    }
    Value::Number(count as f64)
}

fn fn_countif(args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    if args.len() != 2 {
        return Value::Error(ErrorKind::Value);
    }
    let range = match &args[0] {
        Value::Range(r) => RangeArg::Range(*r),
        Value::Array(a) => RangeArg::Array(a),
        _ => return Value::Error(ErrorKind::Value),
    };
    let Some(criteria) = parse_numeric_criteria(&args[1]) else {
        return Value::Error(ErrorKind::Value);
    };
    let count = match range {
        RangeArg::Range(r) => match count_if_range(grid, r.resolve(base), criteria) {
            Ok(c) => c,
            Err(e) => return Value::Error(e),
        },
        RangeArg::Array(a) => count_if_f64_blank_as_zero(a.as_slice(), criteria),
    };
    Value::Number(count as f64)
}

fn fn_sumproduct(args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    if args.len() != 2 {
        return Value::Error(ErrorKind::Value);
    }
    match (&args[0], &args[1]) {
        (Value::Array(a), Value::Array(b)) => {
            if a.len() != b.len() {
                return Value::Error(ErrorKind::Value);
            }
            Value::Number(simd::sumproduct_ignore_nan_f64(&a.values, &b.values))
        }
        (Value::Range(a), Value::Range(b)) => {
            let ra = a.resolve(base);
            let rb = b.resolve(base);
            match sumproduct_range(grid, ra, rb) {
                Ok(v) => Value::Number(v),
                Err(e) => Value::Error(e),
            }
        }
        _ => Value::Error(ErrorKind::Value),
    }
}

enum RangeArg<'a> {
    Range(RangeRef),
    Array(&'a ArrayValue),
}

fn parse_numeric_criteria(v: &Value) -> Option<NumericCriteria> {
    match v {
        Value::Number(n) => Some(NumericCriteria::new(CmpOp::Eq, *n)),
        Value::Bool(b) => Some(NumericCriteria::new(CmpOp::Eq, if *b { 1.0 } else { 0.0 })),
        Value::Text(s) => parse_criteria_str(s),
        _ => None,
    }
}

fn parse_criteria_str(s: &str) -> Option<NumericCriteria> {
    let s = s.trim();
    let (op, rest) = if let Some(r) = s.strip_prefix(">=") {
        (CmpOp::Ge, r)
    } else if let Some(r) = s.strip_prefix("<=") {
        (CmpOp::Le, r)
    } else if let Some(r) = s.strip_prefix("<>") {
        (CmpOp::Ne, r)
    } else if let Some(r) = s.strip_prefix('>') {
        (CmpOp::Gt, r)
    } else if let Some(r) = s.strip_prefix('<') {
        (CmpOp::Lt, r)
    } else if let Some(r) = s.strip_prefix('=') {
        (CmpOp::Eq, r)
    } else {
        (CmpOp::Eq, s)
    };
    let rhs_str = rest.trim();
    if rhs_str.is_empty() {
        return None;
    }
    let rhs = parse_number_from_text(rhs_str).ok()?;
    Some(NumericCriteria::new(op, rhs))
}

#[inline]
fn range_in_bounds(grid: &dyn Grid, range: ResolvedRange) -> bool {
    grid.in_bounds(CellCoord {
        row: range.row_start,
        col: range.col_start,
    }) && grid.in_bounds(CellCoord {
        row: range.row_end,
        col: range.col_end,
    })
}

fn sum_range(grid: &dyn Grid, range: ResolvedRange) -> Result<f64, ErrorKind> {
    if !range_in_bounds(grid, range) {
        return Err(ErrorKind::Ref);
    }
    let mut sum = 0.0;
    for col in range.col_start..=range.col_end {
        if let Some(slice) = grid.column_slice(col, range.row_start, range.row_end) {
            sum += simd::sum_ignore_nan_f64(slice);
        } else {
            for row in range.row_start..=range.row_end {
                match grid.get_value(CellCoord { row, col }) {
                    Value::Number(v) => sum += v,
                    Value::Error(e) => return Err(e),
                    // SUM ignores text/logicals/blanks in references.
                    Value::Bool(_)
                    | Value::Text(_)
                    | Value::Empty
                    | Value::Array(_)
                    | Value::Range(_) => {}
                }
            }
        }
    }
    Ok(sum)
}

fn sum_count_range(grid: &dyn Grid, range: ResolvedRange) -> Result<(f64, usize), ErrorKind> {
    if !range_in_bounds(grid, range) {
        return Err(ErrorKind::Ref);
    }
    let mut sum = 0.0;
    let mut count = 0usize;
    for col in range.col_start..=range.col_end {
        if let Some(slice) = grid.column_slice(col, range.row_start, range.row_end) {
            let (s, c) = simd::sum_count_ignore_nan_f64(slice);
            sum += s;
            count += c;
        } else {
            for row in range.row_start..=range.row_end {
                match grid.get_value(CellCoord { row, col }) {
                    Value::Number(v) => {
                        sum += v;
                        count += 1;
                    }
                    Value::Error(e) => return Err(e),
                    // Ignore non-numeric values in references.
                    Value::Bool(_)
                    | Value::Text(_)
                    | Value::Empty
                    | Value::Array(_)
                    | Value::Range(_) => {}
                }
            }
        }
    }
    Ok((sum, count))
}

fn count_range(grid: &dyn Grid, range: ResolvedRange) -> Result<usize, ErrorKind> {
    if !range_in_bounds(grid, range) {
        return Err(ErrorKind::Ref);
    }
    let mut count = 0usize;
    for col in range.col_start..=range.col_end {
        if let Some(slice) = grid.column_slice(col, range.row_start, range.row_end) {
            count += simd::count_ignore_nan_f64(slice);
        } else {
            for row in range.row_start..=range.row_end {
                if matches!(grid.get_value(CellCoord { row, col }), Value::Number(_)) {
                    count += 1
                }
            }
        }
    }
    Ok(count)
}

fn min_range(grid: &dyn Grid, range: ResolvedRange) -> Result<Option<f64>, ErrorKind> {
    if !range_in_bounds(grid, range) {
        return Err(ErrorKind::Ref);
    }
    let mut out: Option<f64> = None;
    for col in range.col_start..=range.col_end {
        let col_min = if let Some(slice) = grid.column_slice(col, range.row_start, range.row_end) {
            simd::min_ignore_nan_f64(slice)
        } else {
            let mut m: Option<f64> = None;
            for row in range.row_start..=range.row_end {
                match grid.get_value(CellCoord { row, col }) {
                    Value::Number(v) => m = Some(m.map_or(v, |prev| prev.min(v))),
                    Value::Error(e) => return Err(e),
                    Value::Bool(_)
                    | Value::Text(_)
                    | Value::Empty
                    | Value::Array(_)
                    | Value::Range(_) => {}
                }
            }
            m
        };
        if let Some(m) = col_min {
            out = Some(out.map_or(m, |prev| prev.min(m)));
        }
    }
    Ok(out)
}

fn max_range(grid: &dyn Grid, range: ResolvedRange) -> Result<Option<f64>, ErrorKind> {
    if !range_in_bounds(grid, range) {
        return Err(ErrorKind::Ref);
    }
    let mut out: Option<f64> = None;
    for col in range.col_start..=range.col_end {
        let col_max = if let Some(slice) = grid.column_slice(col, range.row_start, range.row_end) {
            simd::max_ignore_nan_f64(slice)
        } else {
            let mut m: Option<f64> = None;
            for row in range.row_start..=range.row_end {
                match grid.get_value(CellCoord { row, col }) {
                    Value::Number(v) => m = Some(m.map_or(v, |prev| prev.max(v))),
                    Value::Error(e) => return Err(e),
                    Value::Bool(_)
                    | Value::Text(_)
                    | Value::Empty
                    | Value::Array(_)
                    | Value::Range(_) => {}
                }
            }
            m
        };
        if let Some(m) = col_max {
            out = Some(out.map_or(m, |prev| prev.max(m)));
        }
    }
    Ok(out)
}

fn count_if_range(
    grid: &dyn Grid,
    range: ResolvedRange,
    criteria: NumericCriteria,
) -> Result<usize, ErrorKind> {
    if !range_in_bounds(grid, range) {
        return Err(ErrorKind::Ref);
    }
    let mut count = 0usize;
    for col in range.col_start..=range.col_end {
        if let Some(slice) = grid.column_slice(col, range.row_start, range.row_end) {
            count += count_if_f64_blank_as_zero(slice, criteria);
        } else {
            for row in range.row_start..=range.row_end {
                if let Some(v) = coerce_countif_value_to_number(grid.get_value(CellCoord { row, col }))
                {
                    if matches_numeric_criteria(v, criteria) {
                        count += 1;
                    }
                }
            }
        }
    }
    Ok(count)
}

fn coerce_sumproduct_number(v: Value) -> Result<f64, ErrorKind> {
    match v {
        Value::Number(n) => Ok(n),
        Value::Bool(b) => Ok(if b { 1.0 } else { 0.0 }),
        Value::Text(s) => match parse_number_from_text(&s) {
            Ok(n) => Ok(n),
            Err(ErrorKind::Value) => Ok(0.0),
            Err(e) => Err(e),
        },
        Value::Empty => Ok(0.0),
        Value::Error(e) => Err(e),
        Value::Array(_) | Value::Range(_) => Err(ErrorKind::Value),
    }
}

fn sumproduct_range(grid: &dyn Grid, a: ResolvedRange, b: ResolvedRange) -> Result<f64, ErrorKind> {
    if !range_in_bounds(grid, a) || !range_in_bounds(grid, b) {
        return Err(ErrorKind::Ref);
    }
    if a.rows() != b.rows() || a.cols() != b.cols() {
        return Err(ErrorKind::Value);
    }
    let rows = a.rows();
    let cols = a.cols();
    let mut sum = 0.0;
    for col_offset in 0..cols {
        let col_a = a.col_start + col_offset;
        let col_b = b.col_start + col_offset;
        if let (Some(sa), Some(sb)) = (
            grid.column_slice(col_a, a.row_start, a.row_end),
            grid.column_slice(col_b, b.row_start, b.row_end),
        ) {
            sum += simd::sumproduct_ignore_nan_f64(sa, sb);
            continue;
        }
        for row_offset in 0..rows {
            let ra = CellCoord {
                row: a.row_start + row_offset,
                col: col_a,
            };
            let rb = CellCoord {
                row: b.row_start + row_offset,
                col: col_b,
            };
            let x = coerce_sumproduct_number(grid.get_value(ra))?;
            let y = coerce_sumproduct_number(grid.get_value(rb))?;
            sum += x * y;
        }
    }
    Ok(sum)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::bytecode::ColumnarGrid;

    #[test]
    fn range_aggregates_return_ref_for_out_of_bounds_ranges() {
        let grid = ColumnarGrid::new(10, 10);

        let range = ResolvedRange {
            row_start: 0,
            row_end: 20,
            col_start: 0,
            col_end: 0,
        };

        assert_eq!(sum_range(&grid, range), Err(ErrorKind::Ref));
        assert_eq!(sum_count_range(&grid, range), Err(ErrorKind::Ref));
        assert_eq!(count_range(&grid, range), Err(ErrorKind::Ref));

        let criteria = NumericCriteria::new(CmpOp::Gt, 0.0);
        assert_eq!(count_if_range(&grid, range, criteria), Err(ErrorKind::Ref));
        assert_eq!(min_range(&grid, range), Err(ErrorKind::Ref));
        assert_eq!(max_range(&grid, range), Err(ErrorKind::Ref));

        assert_eq!(sumproduct_range(&grid, range, range), Err(ErrorKind::Ref));
    }
}
