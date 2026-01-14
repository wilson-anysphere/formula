use crate::eval::CompiledExpr;
use crate::functions::{eval_scalar_arg, ArgValue, ArraySupport, FunctionContext, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::{Array, ErrorKind, Value};

#[derive(Debug, Clone)]
struct MatrixArg {
    rows: usize,
    cols: usize,
    /// Row-major order.
    values: Vec<Option<f64>>,
}

fn scalar_to_number(ctx: &dyn FunctionContext, value: Value) -> Result<Option<f64>, ErrorKind> {
    match value {
        Value::Error(e) => Err(e),
        Value::Number(n) => Ok(Some(n)),
        Value::Bool(b) => Ok(Some(if b { 1.0 } else { 0.0 })),
        Value::Blank => Ok(None),
        Value::Text(s) => Ok(Some(Value::Text(s).coerce_to_number_with_ctx(ctx)?)),
        Value::Entity(_) | Value::Record(_) => Err(ErrorKind::Value),
        Value::Array(arr) => Ok(Some(arr.top_left().coerce_to_number_with_ctx(ctx)?)),
        Value::Lambda(_) | Value::Reference(_) | Value::ReferenceUnion(_) | Value::Spill { .. } => {
            Err(ErrorKind::Value)
        }
    }
}

fn value_to_optional_number_in_range(value: Value) -> Result<Option<f64>, ErrorKind> {
    match value {
        Value::Error(e) => Err(e),
        Value::Number(n) => Ok(Some(n)),
        Value::Lambda(_) => Err(ErrorKind::Value),
        Value::Bool(_)
        | Value::Text(_)
        | Value::Entity(_)
        | Value::Record(_)
        | Value::Blank
        | Value::Array(_)
        | Value::Spill { .. }
        | Value::Reference(_)
        | Value::ReferenceUnion(_) => Ok(None),
    }
}

fn arg_to_matrix(ctx: &dyn FunctionContext, arg: ArgValue) -> Result<MatrixArg, ErrorKind> {
    match arg {
        ArgValue::Scalar(v) => match v {
            Value::Array(arr) => {
                let mut out = Vec::with_capacity(arr.values.len());
                for el in arr.iter() {
                    match el {
                        Value::Error(e) => return Err(*e),
                        Value::Number(n) => out.push(Some(*n)),
                        Value::Lambda(_) => return Err(ErrorKind::Value),
                        Value::Bool(_)
                        | Value::Text(_)
                        | Value::Entity(_)
                        | Value::Record(_)
                        | Value::Blank
                        | Value::Array(_)
                        | Value::Spill { .. }
                        | Value::Reference(_)
                        | Value::ReferenceUnion(_) => out.push(None),
                    }
                }
                Ok(MatrixArg {
                    rows: arr.rows,
                    cols: arr.cols,
                    values: out,
                })
            }
            other => Ok(MatrixArg {
                rows: 1,
                cols: 1,
                values: vec![scalar_to_number(ctx, other)?],
            }),
        },
        ArgValue::Reference(r) => {
            let r = r.normalized();
            ctx.record_reference(&r);
            let rows = (r.end.row - r.start.row + 1) as usize;
            let cols = (r.end.col - r.start.col + 1) as usize;
            let mut out = Vec::with_capacity(rows.saturating_mul(cols));
            for addr in r.iter_cells() {
                let v = ctx.get_cell_value(&r.sheet_id, addr);
                out.push(value_to_optional_number_in_range(v)?);
            }
            Ok(MatrixArg {
                rows,
                cols,
                values: out,
            })
        }
        ArgValue::ReferenceUnion(_) => Err(ErrorKind::Value),
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum XOrientation {
    /// Single predictor vector.
    Vector,
    /// Observations are in rows, predictors in columns (n x p).
    ColumnsArePredictors,
    /// Observations are in columns, predictors in rows (p x n).
    RowsArePredictors,
}

#[derive(Debug, Clone)]
struct ParsedXY {
    /// Original y shape (rows, cols).
    y_shape: (usize, usize),
    /// Filtered y values (length n).
    y: Vec<f64>,
    /// Filtered x matrix row-major (n * p).
    x: Vec<f64>,
    /// Indices in the original vector that were kept in `y` / `x`.
    kept_obs_indices: Vec<usize>,
    /// Predictor count.
    p: usize,
    /// How X was interpreted.
    x_orientation: XOrientation,
}

fn parse_y_vector(y: &MatrixArg) -> Result<(usize, usize, Vec<Option<f64>>), ErrorKind> {
    if !(y.rows == 1 || y.cols == 1) {
        // Excel supports multiple dependent variables, but we currently require a single y vector.
        return Err(ErrorKind::Ref);
    }
    Ok((y.rows, y.cols, y.values.clone()))
}

fn parse_x_matrix(
    x: &MatrixArg,
    n_obs: usize,
) -> Result<(XOrientation, usize /*p*/, Vec<Option<f64>>), ErrorKind> {
    if x.rows == 0 || x.cols == 0 {
        return Err(ErrorKind::Ref);
    }

    // Vector is always treated as a single predictor regardless of orientation.
    if x.rows == 1 || x.cols == 1 {
        if x.values.len() != n_obs {
            return Err(ErrorKind::Ref);
        }
        return Ok((XOrientation::Vector, 1, x.values.clone()));
    }

    if x.rows == n_obs {
        Ok((XOrientation::ColumnsArePredictors, x.cols, x.values.clone()))
    } else if x.cols == n_obs {
        Ok((XOrientation::RowsArePredictors, x.rows, x.values.clone()))
    } else {
        Err(ErrorKind::Ref)
    }
}

fn build_filtered_xy(
    y: &[Option<f64>],
    x: &[Option<f64>],
    n_obs: usize,
    p: usize,
    x_orientation: XOrientation,
    x_cols: usize,
) -> Result<(Vec<f64>, Vec<f64>, Vec<usize>), ErrorKind> {
    debug_assert_eq!(y.len(), n_obs);

    let mut out_y = Vec::new();
    let mut out_x = Vec::new();
    let mut kept = Vec::new();

    out_y.try_reserve_exact(n_obs).map_err(|_| ErrorKind::Num)?;
    out_x
        .try_reserve_exact(n_obs.saturating_mul(p))
        .map_err(|_| ErrorKind::Num)?;
    kept.try_reserve_exact(n_obs).map_err(|_| ErrorKind::Num)?;

    for obs in 0..n_obs {
        let Some(yv) = y[obs] else {
            continue;
        };
        if !yv.is_finite() {
            return Err(ErrorKind::Num);
        }

        let mut row = Vec::with_capacity(p);
        let mut row_ok = true;
        match x_orientation {
            XOrientation::Vector => match x.get(obs).copied().flatten() {
                Some(xv) => {
                    if !xv.is_finite() {
                        return Err(ErrorKind::Num);
                    }
                    row.push(xv);
                }
                None => row_ok = false,
            },
            XOrientation::ColumnsArePredictors => {
                // x is n x p in row-major: (obs, pred)
                for pred in 0..p {
                    let idx = obs * x_cols + pred;
                    let Some(xv) = x.get(idx).copied().flatten() else {
                        row_ok = false;
                        break;
                    };
                    if !xv.is_finite() {
                        return Err(ErrorKind::Num);
                    }
                    row.push(xv);
                }
            }
            XOrientation::RowsArePredictors => {
                // x is p x n in row-major: (pred, obs). Here x_cols == n_obs.
                for pred in 0..p {
                    let idx = pred * x_cols + obs;
                    let Some(xv) = x.get(idx).copied().flatten() else {
                        row_ok = false;
                        break;
                    };
                    if !xv.is_finite() {
                        return Err(ErrorKind::Num);
                    }
                    row.push(xv);
                }
            }
        }

        if !row_ok {
            continue;
        }

        debug_assert_eq!(row.len(), p);
        out_y.push(yv);
        out_x.extend_from_slice(&row);
        kept.push(obs);
    }

    Ok((out_y, out_x, kept))
}

fn parse_known_xy(
    ctx: &dyn FunctionContext,
    known_y_expr: &CompiledExpr,
    known_x_expr: Option<&CompiledExpr>,
) -> Result<ParsedXY, ErrorKind> {
    let known_y = arg_to_matrix(ctx, ctx.eval_arg(known_y_expr))?;
    let (y_rows, y_cols, y_values) = parse_y_vector(&known_y)?;
    let n_obs = y_values.len();

    let (x_orientation, p, x_values, x_cols) = if let Some(expr) = known_x_expr {
        let arg = ctx.eval_arg(expr);
        // Treat a literal blank (omitted arg) as missing and default to 1..n.
        if matches!(arg, ArgValue::Scalar(Value::Blank)) {
            let mut xs = Vec::with_capacity(n_obs);
            for i in 0..n_obs {
                xs.push(Some((i + 1) as f64));
            }
            (XOrientation::Vector, 1, xs, 1)
        } else {
            let x_mat = arg_to_matrix(ctx, arg)?;
            let (orient, p, vals) = parse_x_matrix(&x_mat, n_obs)?;
            (orient, p, vals, x_mat.cols)
        }
    } else {
        let mut xs = Vec::with_capacity(n_obs);
        for i in 0..n_obs {
            xs.push(Some((i + 1) as f64));
        }
        (XOrientation::Vector, 1, xs, 1)
    };

    if p == 0 {
        return Err(ErrorKind::Ref);
    }

    let (y, x, kept_obs_indices) =
        build_filtered_xy(&y_values, &x_values, n_obs, p, x_orientation, x_cols)?;

    Ok(ParsedXY {
        y_shape: (y_rows, y_cols),
        y,
        x,
        kept_obs_indices,
        p,
        x_orientation,
    })
}

fn parse_const_arg(
    ctx: &dyn FunctionContext,
    expr: Option<&CompiledExpr>,
) -> Result<bool, ErrorKind> {
    let Some(expr) = expr else {
        return Ok(true);
    };
    let v = eval_scalar_arg(ctx, expr);
    if matches!(v, Value::Blank) {
        // Excel default.
        return Ok(true);
    }
    v.coerce_to_bool_with_ctx(ctx)
}

fn parse_stats_arg(
    ctx: &dyn FunctionContext,
    expr: Option<&CompiledExpr>,
) -> Result<bool, ErrorKind> {
    let Some(expr) = expr else {
        return Ok(false);
    };
    eval_scalar_arg(ctx, expr).coerce_to_bool_with_ctx(ctx)
}

fn linest_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let stats = match parse_stats_arg(ctx, args.get(3)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let include_intercept = match parse_const_arg(ctx, args.get(2)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let known_x_expr = args.get(1);
    let parsed = match parse_known_xy(ctx, &args[0], known_x_expr) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let fit = match crate::functions::statistical::regression::linest(
        &parsed.y,
        &parsed.x,
        parsed.y.len(),
        parsed.p,
        include_intercept,
        stats,
    ) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let cols = parsed.p + 1;
    let rows: usize = if stats { 5 } else { 1 };
    let total = rows.saturating_mul(cols);
    let mut out = Vec::new();
    if out.try_reserve_exact(total).is_err() {
        return Value::Error(ErrorKind::Num);
    }

    // Row 1: coefficients in reverse X column order, then intercept.
    for pred in (0..parsed.p).rev() {
        out.push(Value::Number(fit.slopes[pred]));
    }
    out.push(Value::Number(fit.intercept));

    if stats {
        // Row 2: standard errors.
        match (&fit.slope_standard_errors, fit.intercept_standard_error) {
            (Some(se_slopes), Some(se_intercept)) => {
                for pred in (0..parsed.p).rev() {
                    out.push(Value::Number(se_slopes[pred]));
                }
                out.push(Value::Number(se_intercept));
            }
            _ => {
                for _ in 0..cols {
                    out.push(Value::Error(ErrorKind::Div0));
                }
            }
        }

        // Row 3: R^2, se_y; rest #N/A.
        out.push(Value::Number(fit.r_squared));
        out.push(match fit.standard_error_y {
            Some(v) => Value::Number(v),
            None => Value::Error(ErrorKind::Div0),
        });
        for _ in 2..cols {
            out.push(Value::Error(ErrorKind::NA));
        }

        // Row 4: F, df; rest #N/A.
        out.push(match fit.f_statistic {
            Some(v) => Value::Number(v),
            None => Value::Error(ErrorKind::Div0),
        });
        out.push(Value::Number(fit.df_resid));
        for _ in 2..cols {
            out.push(Value::Error(ErrorKind::NA));
        }

        // Row 5: ssreg, ssresid; rest #N/A.
        out.push(Value::Number(fit.ss_regression));
        out.push(Value::Number(fit.ss_resid));
        for _ in 2..cols {
            out.push(Value::Error(ErrorKind::NA));
        }
    }

    Value::Array(Array::new(rows, cols, out))
}

fn logest_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let stats = match parse_stats_arg(ctx, args.get(3)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let include_intercept = match parse_const_arg(ctx, args.get(2)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let known_x_expr = args.get(1);
    let parsed = match parse_known_xy(ctx, &args[0], known_x_expr) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let fit = match crate::functions::statistical::regression::logest(
        &parsed.y,
        &parsed.x,
        parsed.y.len(),
        parsed.p,
        include_intercept,
        stats,
    ) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let cols = parsed.p + 1;
    let rows: usize = if stats { 5 } else { 1 };
    let total = rows.saturating_mul(cols);
    let mut out = Vec::new();
    if out.try_reserve_exact(total).is_err() {
        return Value::Error(ErrorKind::Num);
    }

    // Row 1: coefficients (m values) in reverse order, then b.
    for pred in (0..parsed.p).rev() {
        out.push(Value::Number(fit.bases[pred]));
    }
    out.push(Value::Number(fit.intercept));

    if stats {
        match (&fit.base_standard_errors, fit.intercept_standard_error) {
            (Some(se_bases), Some(se_intercept)) => {
                for pred in (0..parsed.p).rev() {
                    out.push(Value::Number(se_bases[pred]));
                }
                out.push(Value::Number(se_intercept));
            }
            _ => {
                for _ in 0..cols {
                    out.push(Value::Error(ErrorKind::Div0));
                }
            }
        }

        out.push(Value::Number(fit.r_squared));
        out.push(match fit.standard_error_y {
            Some(v) => Value::Number(v),
            None => Value::Error(ErrorKind::Div0),
        });
        for _ in 2..cols {
            out.push(Value::Error(ErrorKind::NA));
        }

        out.push(match fit.f_statistic {
            Some(v) => Value::Number(v),
            None => Value::Error(ErrorKind::Div0),
        });
        out.push(Value::Number(fit.df_resid));
        for _ in 2..cols {
            out.push(Value::Error(ErrorKind::NA));
        }

        out.push(Value::Number(fit.ss_regression));
        out.push(Value::Number(fit.ss_resid));
        for _ in 2..cols {
            out.push(Value::Error(ErrorKind::NA));
        }
    }

    Value::Array(Array::new(rows, cols, out))
}

fn predict_for_known_xy(parsed: &ParsedXY, predict_row: impl Fn(&[f64]) -> f64) -> Value {
    let (out_rows, out_cols) = parsed.y_shape;
    let total = out_rows.saturating_mul(out_cols);
    let mut out = vec![Value::Error(ErrorKind::NA); total];

    for (row_idx, &obs_idx) in parsed.kept_obs_indices.iter().enumerate() {
        let base = row_idx * parsed.p;
        let row = &parsed.x[base..base + parsed.p];
        let yhat = predict_row(row);
        out[obs_idx] = if yhat.is_finite() {
            Value::Number(yhat)
        } else {
            Value::Error(ErrorKind::Num)
        };
    }

    Value::Array(Array::new(out_rows, out_cols, out))
}

fn trend_predict_with_arg(
    ctx: &dyn FunctionContext,
    parsed: &ParsedXY,
    new_x_arg: Option<ArgValue>,
    predict_row: impl Fn(&[f64]) -> f64,
    p: usize,
) -> Value {
    let Some(new_x_arg) = new_x_arg else {
        return predict_for_known_xy(parsed, predict_row);
    };
    if matches!(new_x_arg, ArgValue::Scalar(Value::Blank)) {
        return predict_for_known_xy(parsed, predict_row);
    }

    let new_x = match arg_to_matrix(ctx, new_x_arg) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    if p == 1 {
        // Single predictor: preserve shape.
        let total = new_x.rows.saturating_mul(new_x.cols);
        let mut out = Vec::with_capacity(total);
        for el in new_x.values {
            let Some(xv) = el else {
                out.push(Value::Error(ErrorKind::Value));
                continue;
            };
            if !xv.is_finite() {
                out.push(Value::Error(ErrorKind::Num));
                continue;
            }
            let yhat = predict_row(&[xv]);
            out.push(if yhat.is_finite() {
                Value::Number(yhat)
            } else {
                Value::Error(ErrorKind::Num)
            });
        }
        return Value::Array(Array::new(new_x.rows, new_x.cols, out));
    }

    // Multi predictor: require consistent orientation with known_x.
    match parsed.x_orientation {
        XOrientation::ColumnsArePredictors => {
            if new_x.cols != p {
                return Value::Error(ErrorKind::Ref);
            }
            let n_new = new_x.rows;
            let mut out = Vec::with_capacity(n_new);
            for obs in 0..n_new {
                let mut row = Vec::with_capacity(p);
                let mut ok = true;
                for pred in 0..p {
                    let idx = obs * new_x.cols + pred;
                    let Some(xv) = new_x.values[idx] else {
                        ok = false;
                        break;
                    };
                    if !xv.is_finite() {
                        ok = false;
                        break;
                    }
                    row.push(xv);
                }
                if !ok {
                    out.push(Value::Error(ErrorKind::Value));
                    continue;
                }
                let yhat = predict_row(&row);
                out.push(if yhat.is_finite() {
                    Value::Number(yhat)
                } else {
                    Value::Error(ErrorKind::Num)
                });
            }
            Value::Array(Array::new(n_new, 1, out))
        }
        XOrientation::RowsArePredictors => {
            if new_x.rows != p {
                return Value::Error(ErrorKind::Ref);
            }
            let n_new = new_x.cols;
            let mut out = Vec::with_capacity(n_new);
            for obs in 0..n_new {
                let mut row = Vec::with_capacity(p);
                let mut ok = true;
                for pred in 0..p {
                    let idx = pred * new_x.cols + obs;
                    let Some(xv) = new_x.values[idx] else {
                        ok = false;
                        break;
                    };
                    if !xv.is_finite() {
                        ok = false;
                        break;
                    }
                    row.push(xv);
                }
                if !ok {
                    out.push(Value::Error(ErrorKind::Value));
                    continue;
                }
                let yhat = predict_row(&row);
                out.push(if yhat.is_finite() {
                    Value::Number(yhat)
                } else {
                    Value::Error(ErrorKind::Num)
                });
            }
            Value::Array(Array::new(1, n_new, out))
        }
        XOrientation::Vector => Value::Error(ErrorKind::Ref),
    }
}

fn trend_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    // Parse args the same way as TREND.
    let include_intercept = match parse_const_arg(ctx, args.get(3)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let parsed = match parse_known_xy(ctx, &args[0], args.get(1)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let fit = match crate::functions::statistical::regression::linest(
        &parsed.y,
        &parsed.x,
        parsed.y.len(),
        parsed.p,
        include_intercept,
        false,
    ) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let new_x_arg = args.get(2).map(|expr| ctx.eval_arg(expr));

    trend_predict_with_arg(
        ctx,
        &parsed,
        new_x_arg,
        |row| {
            let mut acc = fit.intercept;
            for (m, x) in fit.slopes.iter().zip(row.iter()) {
                acc += m * *x;
            }
            acc
        },
        fit.slopes.len(),
    )
}

fn growth_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let include_intercept = match parse_const_arg(ctx, args.get(3)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let parsed = match parse_known_xy(ctx, &args[0], args.get(1)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let fit = match crate::functions::statistical::regression::logest(
        &parsed.y,
        &parsed.x,
        parsed.y.len(),
        parsed.p,
        include_intercept,
        false,
    ) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let new_x_arg = args.get(2).map(|expr| ctx.eval_arg(expr));

    trend_predict_with_arg(
        ctx,
        &parsed,
        new_x_arg,
        |row| {
            let mut acc = fit.intercept;
            for (m, x) in fit.bases.iter().zip(row.iter()) {
                acc *= m.powf(*x);
            }
            acc
        },
        fit.bases.len(),
    )
}

inventory::submit! {
    FunctionSpec {
        name: "LINEST",
        min_args: 1,
        max_args: 4,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Bool, ValueType::Bool],
        implementation: linest_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "LOGEST",
        min_args: 1,
        max_args: 4,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Bool, ValueType::Bool],
        implementation: logest_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "TREND",
        min_args: 1,
        max_args: 4,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Any, ValueType::Bool],
        implementation: trend_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "GROWTH",
        min_args: 1,
        max_args: 4,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Any, ValueType::Bool],
        implementation: growth_fn,
    }
}

// On wasm targets, `inventory` registrations can be dropped by the linker if the object file
// contains no otherwise-referenced symbols. Referencing this function from a `#[used]` table in
// `functions/mod.rs` ensures the module (and its `inventory::submit!` entries) are retained.
#[cfg(target_arch = "wasm32")]
pub(super) fn __force_link() {}
