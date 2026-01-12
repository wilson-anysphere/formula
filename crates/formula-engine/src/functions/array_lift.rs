use crate::eval::CompiledExpr;
use crate::functions::{ArgValue, FunctionContext, Reference};
use crate::value::{Array, ErrorKind, Value};

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) struct Shape {
    pub(crate) rows: usize,
    pub(crate) cols: usize,
}

impl Shape {
    pub(crate) fn is_1x1(self) -> bool {
        self.rows == 1 && self.cols == 1
    }
}

pub(crate) fn value_shape(value: &Value) -> Option<Shape> {
    match value {
        Value::Array(arr) => Some(Shape {
            rows: arr.rows,
            cols: arr.cols,
        }),
        _ => None,
    }
}

pub(crate) fn dominant_shape(values: &[&Value]) -> Result<Option<Shape>, ErrorKind> {
    let mut dominant: Option<Shape> = None;
    let mut saw_array = false;

    for value in values {
        let Some(shape) = value_shape(value) else {
            continue;
        };
        saw_array = true;

        if shape.is_1x1() {
            continue;
        }

        match dominant {
            None => dominant = Some(shape),
            Some(existing) if existing == shape => {}
            Some(_) => return Err(ErrorKind::Value),
        }
    }

    if dominant.is_some() {
        return Ok(dominant);
    }

    if saw_array {
        return Ok(Some(Shape { rows: 1, cols: 1 }));
    }

    Ok(None)
}

pub(crate) fn broadcast_compatible(value: &Value, target: Shape) -> bool {
    match value {
        Value::Array(arr) => {
            (arr.rows == target.rows && arr.cols == target.cols) || (arr.rows == 1 && arr.cols == 1)
        }
        _ => true,
    }
}

pub(crate) fn element_at<'a>(value: &'a Value, target: Shape, idx: usize) -> &'a Value {
    match value {
        Value::Array(arr) => {
            if arr.rows == 1 && arr.cols == 1 {
                return &arr.values[0];
            }
            debug_assert_eq!(arr.rows, target.rows);
            debug_assert_eq!(arr.cols, target.cols);
            &arr.values[idx]
        }
        other => other,
    }
}

pub(crate) fn eval_arg(ctx: &dyn FunctionContext, expr: &CompiledExpr) -> Value {
    match ctx.eval_arg(expr) {
        ArgValue::Scalar(v) => v,
        ArgValue::Reference(r) => reference_to_value(ctx, r),
        ArgValue::ReferenceUnion(_) => Value::Error(ErrorKind::Value),
    }
}

fn reference_to_value(ctx: &dyn FunctionContext, reference: Reference) -> Value {
    let reference = reference.normalized();
    ctx.record_reference(&reference);
    if reference.is_single_cell() {
        return ctx.get_cell_value(&reference.sheet_id, reference.start);
    }

    let rows = (reference.end.row - reference.start.row + 1) as usize;
    let cols = (reference.end.col - reference.start.col + 1) as usize;
    let total = match rows.checked_mul(cols) {
        Some(v) => v,
        None => return Value::Error(ErrorKind::Spill),
    };
    if total > crate::eval::MAX_MATERIALIZED_ARRAY_CELLS {
        return Value::Error(ErrorKind::Spill);
    }
    let mut values = Vec::new();
    if values.try_reserve_exact(total).is_err() {
        return Value::Error(ErrorKind::Num);
    }
    for addr in reference.iter_cells() {
        values.push(ctx.get_cell_value(&reference.sheet_id, addr));
    }
    Value::Array(Array::new(rows, cols, values))
}

pub(crate) fn lift1(value: Value, f: impl Fn(&Value) -> Result<Value, ErrorKind>) -> Value {
    let Some(shape) = (match dominant_shape(&[&value]) {
        Ok(shape) => shape,
        Err(e) => return Value::Error(e),
    }) else {
        return match f(&value) {
            Ok(v) => v,
            Err(e) => Value::Error(e),
        };
    };

    if !broadcast_compatible(&value, shape) {
        return Value::Error(ErrorKind::Value);
    }

    let total = match shape.rows.checked_mul(shape.cols) {
        Some(v) => v,
        None => return Value::Error(ErrorKind::Spill),
    };
    if total > crate::eval::MAX_MATERIALIZED_ARRAY_CELLS {
        return Value::Error(ErrorKind::Spill);
    }
    let mut out = Vec::new();
    if out.try_reserve_exact(total).is_err() {
        return Value::Error(ErrorKind::Num);
    }
    for idx in 0..total {
        let v = element_at(&value, shape, idx);
        out.push(match f(v) {
            Ok(v) => v,
            Err(e) => Value::Error(e),
        });
    }
    Value::Array(Array::new(shape.rows, shape.cols, out))
}

pub(crate) fn lift2(
    a: Value,
    b: Value,
    f: impl Fn(&Value, &Value) -> Result<Value, ErrorKind>,
) -> Value {
    let Some(shape) = (match dominant_shape(&[&a, &b]) {
        Ok(shape) => shape,
        Err(e) => return Value::Error(e),
    }) else {
        return match f(&a, &b) {
            Ok(v) => v,
            Err(e) => Value::Error(e),
        };
    };

    if !broadcast_compatible(&a, shape) || !broadcast_compatible(&b, shape) {
        return Value::Error(ErrorKind::Value);
    }

    let total = match shape.rows.checked_mul(shape.cols) {
        Some(v) => v,
        None => return Value::Error(ErrorKind::Spill),
    };
    if total > crate::eval::MAX_MATERIALIZED_ARRAY_CELLS {
        return Value::Error(ErrorKind::Spill);
    }
    let mut out = Vec::new();
    if out.try_reserve_exact(total).is_err() {
        return Value::Error(ErrorKind::Num);
    }
    for idx in 0..total {
        let av = element_at(&a, shape, idx);
        let bv = element_at(&b, shape, idx);
        out.push(match f(av, bv) {
            Ok(v) => v,
            Err(e) => Value::Error(e),
        });
    }

    Value::Array(Array::new(shape.rows, shape.cols, out))
}

pub(crate) fn lift3(
    a: Value,
    b: Value,
    c: Value,
    f: impl Fn(&Value, &Value, &Value) -> Result<Value, ErrorKind>,
) -> Value {
    let Some(shape) = (match dominant_shape(&[&a, &b, &c]) {
        Ok(shape) => shape,
        Err(e) => return Value::Error(e),
    }) else {
        return match f(&a, &b, &c) {
            Ok(v) => v,
            Err(e) => Value::Error(e),
        };
    };

    if !broadcast_compatible(&a, shape)
        || !broadcast_compatible(&b, shape)
        || !broadcast_compatible(&c, shape)
    {
        return Value::Error(ErrorKind::Value);
    }

    let total = match shape.rows.checked_mul(shape.cols) {
        Some(v) => v,
        None => return Value::Error(ErrorKind::Spill),
    };
    if total > crate::eval::MAX_MATERIALIZED_ARRAY_CELLS {
        return Value::Error(ErrorKind::Spill);
    }
    let mut out = Vec::new();
    if out.try_reserve_exact(total).is_err() {
        return Value::Error(ErrorKind::Num);
    }
    for idx in 0..total {
        let av = element_at(&a, shape, idx);
        let bv = element_at(&b, shape, idx);
        let cv = element_at(&c, shape, idx);
        out.push(match f(av, bv, cv) {
            Ok(v) => v,
            Err(e) => Value::Error(e),
        });
    }

    Value::Array(Array::new(shape.rows, shape.cols, out))
}

pub(crate) fn lift4(
    a: Value,
    b: Value,
    c: Value,
    d: Value,
    f: impl Fn(&Value, &Value, &Value, &Value) -> Result<Value, ErrorKind>,
) -> Value {
    let Some(shape) = (match dominant_shape(&[&a, &b, &c, &d]) {
        Ok(shape) => shape,
        Err(e) => return Value::Error(e),
    }) else {
        return match f(&a, &b, &c, &d) {
            Ok(v) => v,
            Err(e) => Value::Error(e),
        };
    };

    if !broadcast_compatible(&a, shape)
        || !broadcast_compatible(&b, shape)
        || !broadcast_compatible(&c, shape)
        || !broadcast_compatible(&d, shape)
    {
        return Value::Error(ErrorKind::Value);
    }

    let total = match shape.rows.checked_mul(shape.cols) {
        Some(v) => v,
        None => return Value::Error(ErrorKind::Spill),
    };
    if total > crate::eval::MAX_MATERIALIZED_ARRAY_CELLS {
        return Value::Error(ErrorKind::Spill);
    }
    let mut out = Vec::new();
    if out.try_reserve_exact(total).is_err() {
        return Value::Error(ErrorKind::Num);
    }
    for idx in 0..total {
        let av = element_at(&a, shape, idx);
        let bv = element_at(&b, shape, idx);
        let cv = element_at(&c, shape, idx);
        let dv = element_at(&d, shape, idx);
        out.push(match f(av, bv, cv, dv) {
            Ok(v) => v,
            Err(e) => Value::Error(e),
        });
    }

    Value::Array(Array::new(shape.rows, shape.cols, out))
}

pub(crate) fn lift5(
    a: Value,
    b: Value,
    c: Value,
    d: Value,
    e: Value,
    f: impl Fn(&Value, &Value, &Value, &Value, &Value) -> Result<Value, ErrorKind>,
) -> Value {
    let Some(shape) = (match dominant_shape(&[&a, &b, &c, &d, &e]) {
        Ok(shape) => shape,
        Err(err) => return Value::Error(err),
    }) else {
        return match f(&a, &b, &c, &d, &e) {
            Ok(v) => v,
            Err(err) => Value::Error(err),
        };
    };

    if !broadcast_compatible(&a, shape)
        || !broadcast_compatible(&b, shape)
        || !broadcast_compatible(&c, shape)
        || !broadcast_compatible(&d, shape)
        || !broadcast_compatible(&e, shape)
    {
        return Value::Error(ErrorKind::Value);
    }

    let total = match shape.rows.checked_mul(shape.cols) {
        Some(v) => v,
        None => return Value::Error(ErrorKind::Spill),
    };
    if total > crate::eval::MAX_MATERIALIZED_ARRAY_CELLS {
        return Value::Error(ErrorKind::Spill);
    }
    let mut out = Vec::new();
    if out.try_reserve_exact(total).is_err() {
        return Value::Error(ErrorKind::Num);
    }
    for idx in 0..total {
        let av = element_at(&a, shape, idx);
        let bv = element_at(&b, shape, idx);
        let cv = element_at(&c, shape, idx);
        let dv = element_at(&d, shape, idx);
        let ev = element_at(&e, shape, idx);
        out.push(match f(av, bv, cv, dv, ev) {
            Ok(v) => v,
            Err(err) => Value::Error(err),
        });
    }

    Value::Array(Array::new(shape.rows, shape.cols, out))
}
