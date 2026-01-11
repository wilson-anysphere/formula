use super::ast::{BinaryOp, Expr, Function, UnaryOp};
use super::grid::Grid;
use super::value::{Array as ArrayValue, CellCoord, ErrorKind, RangeRef, ResolvedRange, Value};
use crate::simd::{self, CmpOp, NumericCriteria};
use smallvec::SmallVec;

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
                    Function::Unknown(_) => false,
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

fn parse_number_from_text(s: &str) -> Option<f64> {
    let trimmed = s.trim();
    if trimmed.is_empty() {
        return None;
    }
    trimmed.parse::<f64>().ok()
}

fn coerce_to_number(v: Value) -> Result<f64, ErrorKind> {
    match v {
        Value::Number(n) => Ok(n),
        Value::Bool(b) => Ok(if b { 1.0 } else { 0.0 }),
        Value::Empty => Ok(0.0),
        Value::Text(s) => parse_number_from_text(&s).ok_or(ErrorKind::Value),
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
        BinaryOp::Pow => match (coerce_to_number(left), coerce_to_number(right)) {
            (Ok(a), Ok(b)) => Number(a.powf(b)),
            (Err(e), _) | (_, Err(e)) => Error(e),
        },
        BinaryOp::Eq | BinaryOp::Ne | BinaryOp::Lt | BinaryOp::Le | BinaryOp::Gt | BinaryOp::Ge => {
            let a = match coerce_to_number(left) {
                Ok(n) => n,
                Err(e) => return Error(e),
            };
            let b = match coerce_to_number(right) {
                Ok(n) => n,
                Err(e) => return Error(e),
            };
            let res = match op {
                BinaryOp::Eq => a == b,
                BinaryOp::Ne => a != b,
                BinaryOp::Lt => a < b,
                BinaryOp::Le => a <= b,
                BinaryOp::Gt => a > b,
                BinaryOp::Ge => a >= b,
                _ => unreachable!(),
            };
            Bool(res)
        }
    }
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
        Function::Unknown(_) => Value::Error(ErrorKind::Name),
    }
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
                Some(v) => sum += v,
                None => return Value::Error(ErrorKind::Value),
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
                Some(v) => {
                    sum += v;
                    count += 1;
                }
                None => return Value::Error(ErrorKind::Value),
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
                Some(v) => out = Some(out.map_or(v, |prev| prev.min(v))),
                None => return Value::Error(ErrorKind::Value),
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
                Some(v) => out = Some(out.map_or(v, |prev| prev.max(v))),
                None => return Value::Error(ErrorKind::Value),
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
            Value::Range(r) => count += count_range(grid, r.resolve(base)),
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
        RangeArg::Range(r) => count_if_range(grid, r.resolve(base), criteria),
        RangeArg::Array(a) => simd::count_if_f64(a.as_slice(), criteria),
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
    let rhs: f64 = rest.trim().parse().ok()?;
    Some(NumericCriteria::new(op, rhs))
}

fn sum_range(grid: &dyn Grid, range: ResolvedRange) -> Result<f64, ErrorKind> {
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

fn count_range(grid: &dyn Grid, range: ResolvedRange) -> usize {
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
    count
}

fn min_range(grid: &dyn Grid, range: ResolvedRange) -> Result<Option<f64>, ErrorKind> {
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

fn count_if_range(grid: &dyn Grid, range: ResolvedRange, criteria: NumericCriteria) -> usize {
    let mut count = 0usize;
    for col in range.col_start..=range.col_end {
        if let Some(slice) = grid.column_slice(col, range.row_start, range.row_end) {
            count += simd::count_if_f64(slice, criteria);
        } else {
            for row in range.row_start..=range.row_end {
                if let Value::Number(v) = grid.get_value(CellCoord { row, col }) {
                    if matches_numeric_criteria(v, criteria) {
                        count += 1;
                    }
                }
            }
        }
    }
    count
}

fn coerce_sumproduct_number(v: Value) -> Result<f64, ErrorKind> {
    match v {
        Value::Number(n) => Ok(n),
        Value::Bool(b) => Ok(if b { 1.0 } else { 0.0 }),
        Value::Text(s) => Ok(parse_number_from_text(&s).unwrap_or(0.0)),
        Value::Empty => Ok(0.0),
        Value::Error(e) => Err(e),
        Value::Array(_) | Value::Range(_) => Err(ErrorKind::Value),
    }
}

fn sumproduct_range(grid: &dyn Grid, a: ResolvedRange, b: ResolvedRange) -> Result<f64, ErrorKind> {
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
