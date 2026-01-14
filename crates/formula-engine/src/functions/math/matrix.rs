use crate::eval::{CellAddr, CompiledExpr};
use crate::functions::{eval_scalar_arg, ArgValue, FunctionContext, Reference};
use crate::value::{Array, ErrorKind, Value};

const SINGULAR_TOL: f64 = 1.0e-12;

#[derive(Debug, Clone)]
struct Matrix {
    rows: usize,
    cols: usize,
    /// Row-major values (length = rows * cols).
    values: Vec<f64>,
}

impl Matrix {
    fn get(&self, row: usize, col: usize) -> f64 {
        self.values[row * self.cols + col]
    }
}

fn eval_matrix_arg(ctx: &dyn FunctionContext, expr: &CompiledExpr) -> Result<Matrix, ErrorKind> {
    match ctx.eval_arg(expr) {
        ArgValue::Scalar(v) => match v {
            Value::Error(e) => Err(e),
            Value::Array(arr) => coerce_array_to_matrix(ctx, &arr),
            Value::Lambda(_) | Value::Spill { .. } => Err(ErrorKind::Value),
            other => {
                let n = other.coerce_to_number_with_ctx(ctx)?;
                if !n.is_finite() {
                    return Err(ErrorKind::Num);
                }
                Ok(Matrix {
                    rows: 1,
                    cols: 1,
                    values: vec![n],
                })
            }
        },
        ArgValue::Reference(r) => coerce_reference_to_matrix(ctx, r),
        ArgValue::ReferenceUnion(_) => Err(ErrorKind::Value),
    }
}

fn coerce_reference_to_matrix(
    ctx: &dyn FunctionContext,
    reference: Reference,
) -> Result<Matrix, ErrorKind> {
    let reference = reference.normalized();
    ctx.record_reference(&reference);

    let rows = (reference.end.row - reference.start.row + 1) as usize;
    let cols = (reference.end.col - reference.start.col + 1) as usize;
    let total = rows.checked_mul(cols).ok_or(ErrorKind::Num)?;

    let mut values = Vec::new();
    if values.try_reserve_exact(total).is_err() {
        return Err(ErrorKind::Num);
    }

    for row in reference.start.row..=reference.end.row {
        for col in reference.start.col..=reference.end.col {
            let v = ctx.get_cell_value(&reference.sheet_id, CellAddr { row, col });
            let n = v.coerce_to_number_with_ctx(ctx)?;
            if !n.is_finite() {
                return Err(ErrorKind::Num);
            }
            values.push(n);
        }
    }

    Ok(Matrix { rows, cols, values })
}

fn coerce_array_to_matrix(ctx: &dyn FunctionContext, array: &Array) -> Result<Matrix, ErrorKind> {
    let total = array.rows.checked_mul(array.cols).ok_or(ErrorKind::Num)?;
    let mut values = Vec::new();
    if values.try_reserve_exact(total).is_err() {
        return Err(ErrorKind::Num);
    }
    for v in array.iter() {
        let n = v.coerce_to_number_with_ctx(ctx)?;
        if !n.is_finite() {
            return Err(ErrorKind::Num);
        }
        values.push(n);
    }
    Ok(Matrix {
        rows: array.rows,
        cols: array.cols,
        values,
    })
}

fn determinant(matrix: &Matrix) -> Result<f64, ErrorKind> {
    if matrix.rows != matrix.cols {
        return Err(ErrorKind::Value);
    }
    let n = matrix.rows;
    if n == 0 {
        return Err(ErrorKind::Value);
    }

    let mut m = Vec::new();
    if m.try_reserve_exact(matrix.values.len()).is_err() {
        return Err(ErrorKind::Num);
    }
    m.extend_from_slice(&matrix.values);
    let mut sign = 1.0;

    for i in 0..n {
        // Partial pivoting: pick the row with the largest absolute value in column i.
        let mut pivot_row = i;
        let mut pivot_abs = m[i * n + i].abs();
        for r in (i + 1)..n {
            let abs = m[r * n + i].abs();
            if abs > pivot_abs {
                pivot_abs = abs;
                pivot_row = r;
            }
        }

        if pivot_abs < SINGULAR_TOL {
            return Ok(0.0);
        }

        if pivot_row != i {
            for c in 0..n {
                m.swap(i * n + c, pivot_row * n + c);
            }
            sign = -sign;
        }

        let pivot = m[i * n + i];
        if !pivot.is_finite() {
            return Err(ErrorKind::Num);
        }

        for r in (i + 1)..n {
            let factor = m[r * n + i] / pivot;
            if !factor.is_finite() {
                return Err(ErrorKind::Num);
            }
            if factor == 0.0 {
                continue;
            }
            for c in i..n {
                let updated = m[r * n + c] - factor * m[i * n + c];
                if !updated.is_finite() {
                    return Err(ErrorKind::Num);
                }
                m[r * n + c] = updated;
            }
        }
    }

    let mut det = sign;
    for i in 0..n {
        det *= m[i * n + i];
    }
    if det.is_finite() {
        Ok(det)
    } else {
        Err(ErrorKind::Num)
    }
}

fn inverse(matrix: &Matrix) -> Result<Matrix, ErrorKind> {
    if matrix.rows != matrix.cols {
        return Err(ErrorKind::Value);
    }
    let n = matrix.rows;
    if n == 0 {
        return Err(ErrorKind::Value);
    }

    let total = n.checked_mul(n).ok_or(ErrorKind::Num)?;
    let mut a = Vec::new();
    if a.try_reserve_exact(matrix.values.len()).is_err() {
        return Err(ErrorKind::Num);
    }
    a.extend_from_slice(&matrix.values);
    let mut inv = Vec::new();
    if inv.try_reserve_exact(total).is_err() {
        return Err(ErrorKind::Num);
    }
    inv.resize(total, 0.0);
    for i in 0..n {
        inv[i * n + i] = 1.0;
    }

    for col in 0..n {
        let mut pivot_row = col;
        let mut pivot_abs = a[col * n + col].abs();
        for r in (col + 1)..n {
            let abs = a[r * n + col].abs();
            if abs > pivot_abs {
                pivot_abs = abs;
                pivot_row = r;
            }
        }

        if pivot_abs < SINGULAR_TOL {
            return Err(ErrorKind::Num);
        }

        if pivot_row != col {
            for c in 0..n {
                a.swap(col * n + c, pivot_row * n + c);
                inv.swap(col * n + c, pivot_row * n + c);
            }
        }

        let pivot = a[col * n + col];
        if !pivot.is_finite() {
            return Err(ErrorKind::Num);
        }
        if pivot.abs() < SINGULAR_TOL {
            return Err(ErrorKind::Num);
        }

        // Normalize the pivot row.
        for c in 0..n {
            a[col * n + c] /= pivot;
            inv[col * n + c] /= pivot;
            if !a[col * n + c].is_finite() || !inv[col * n + c].is_finite() {
                return Err(ErrorKind::Num);
            }
        }

        // Eliminate the pivot column from all other rows.
        for r in 0..n {
            if r == col {
                continue;
            }
            let factor = a[r * n + col];
            if factor == 0.0 {
                continue;
            }
            if !factor.is_finite() {
                return Err(ErrorKind::Num);
            }
            for c in 0..n {
                a[r * n + c] -= factor * a[col * n + c];
                inv[r * n + c] -= factor * inv[col * n + c];
                if !a[r * n + c].is_finite() || !inv[r * n + c].is_finite() {
                    return Err(ErrorKind::Num);
                }
            }
        }
    }

    Ok(Matrix {
        rows: n,
        cols: n,
        values: inv,
    })
}

fn multiply(a: &Matrix, b: &Matrix) -> Result<Matrix, ErrorKind> {
    if a.cols != b.rows {
        return Err(ErrorKind::Value);
    }
    let rows = a.rows;
    let cols = b.cols;
    let inner = a.cols;

    let total = rows.checked_mul(cols).ok_or(ErrorKind::Num)?;
    let mut out = Vec::new();
    if out.try_reserve_exact(total).is_err() {
        return Err(ErrorKind::Num);
    }

    for r in 0..rows {
        for c in 0..cols {
            let mut sum = 0.0;
            for k in 0..inner {
                sum += a.get(r, k) * b.get(k, c);
            }
            if !sum.is_finite() {
                return Err(ErrorKind::Num);
            }
            out.push(sum);
        }
    }

    Ok(Matrix {
        rows,
        cols,
        values: out,
    })
}

fn matrix_to_value_array(matrix: Matrix) -> Result<Value, ErrorKind> {
    let total = matrix.rows.checked_mul(matrix.cols).ok_or(ErrorKind::Num)?;
    let mut values = Vec::new();
    if values.try_reserve_exact(total).is_err() {
        return Err(ErrorKind::Num);
    }
    for v in matrix.values {
        if !v.is_finite() {
            return Err(ErrorKind::Num);
        }
        values.push(Value::Number(v));
    }
    Ok(Value::Array(Array::new(matrix.rows, matrix.cols, values)))
}

pub(crate) fn mdeterm(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let matrix = match eval_matrix_arg(ctx, &args[0]) {
        Ok(m) => m,
        Err(e) => return Value::Error(e),
    };
    if matrix.rows != matrix.cols {
        return Value::Error(ErrorKind::Value);
    }
    match determinant(&matrix) {
        Ok(det) => Value::Number(det),
        Err(e) => Value::Error(e),
    }
}

pub(crate) fn minverse(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let matrix = match eval_matrix_arg(ctx, &args[0]) {
        Ok(m) => m,
        Err(e) => return Value::Error(e),
    };
    if matrix.rows != matrix.cols {
        return Value::Error(ErrorKind::Value);
    }
    match inverse(&matrix) {
        Ok(out) => match matrix_to_value_array(out) {
            Ok(v) => v,
            Err(e) => Value::Error(e),
        },
        Err(e) => Value::Error(e),
    }
}

pub(crate) fn mmult(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let a = match eval_matrix_arg(ctx, &args[0]) {
        Ok(m) => m,
        Err(e) => return Value::Error(e),
    };
    let b = match eval_matrix_arg(ctx, &args[1]) {
        Ok(m) => m,
        Err(e) => return Value::Error(e),
    };

    if a.cols != b.rows {
        return Value::Error(ErrorKind::Value);
    }

    match multiply(&a, &b) {
        Ok(out) => match matrix_to_value_array(out) {
            Ok(v) => v,
            Err(e) => Value::Error(e),
        },
        Err(e) => Value::Error(e),
    }
}

pub(crate) fn munit(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let dim = match eval_scalar_arg(ctx, &args[0]).coerce_to_i64_with_ctx(ctx) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    if dim <= 0 {
        return Value::Error(ErrorKind::Value);
    }
    let dim_usize = match usize::try_from(dim) {
        Ok(v) => v,
        Err(_) => return Value::Error(ErrorKind::Num),
    };

    let total = match dim_usize.checked_mul(dim_usize) {
        Some(v) => v,
        None => return Value::Error(ErrorKind::Num),
    };

    let mut values = Vec::new();
    if values.try_reserve_exact(total).is_err() {
        return Value::Error(ErrorKind::Num);
    }

    for r in 0..dim_usize {
        for c in 0..dim_usize {
            values.push(Value::Number(if r == c { 1.0 } else { 0.0 }));
        }
    }

    Value::Array(Array::new(dim_usize, dim_usize, values))
}
