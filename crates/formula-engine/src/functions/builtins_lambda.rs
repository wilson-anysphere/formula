use std::sync::Arc;

use crate::eval::{CompiledExpr, Expr, SheetReference, LAMBDA_OMITTED_PREFIX};
use crate::functions::{
    ArraySupport, FunctionContext, FunctionSpec, ThreadSafety, ValueType, Volatility,
};
use crate::value::{try_casefold, with_casefolded_key, ErrorKind, Lambda, Value};

const VAR_ARGS: usize = 255;

inventory::submit! {
    FunctionSpec {
        name: "LET",
        min_args: 3,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Any],
        implementation: let_fn,
    }
}

fn let_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    if args.len() < 3 || args.len() % 2 == 0 {
        return Value::Error(ErrorKind::Value);
    }

    let last = args.len() - 1;

    ctx.push_local_scope();
    struct ScopeGuard<'a>(&'a dyn FunctionContext);
    impl Drop for ScopeGuard<'_> {
        fn drop(&mut self) {
            self.0.pop_local_scope();
        }
    }
    let _guard = ScopeGuard(ctx);

    for pair in args[..last].chunks_exact(2) {
        let Some(name) = bare_identifier(&pair[0]) else {
            return Value::Error(ErrorKind::Value);
        };
        let name_key = match try_casefold(name.trim()) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };

        let value = ctx.eval_arg(&pair[1]);
        ctx.set_local_key(name_key, value);
    }

    ctx.eval_formula(&args[last])
}

inventory::submit! {
    FunctionSpec {
        name: "LAMBDA",
        min_args: 1,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Any],
        implementation: lambda_fn,
    }
}

fn lambda_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    if args.is_empty() {
        return Value::Error(ErrorKind::Value);
    }

    let param_count = args.len().saturating_sub(1);
    let mut params: Vec<String> = Vec::new();
    if params.try_reserve_exact(param_count).is_err() {
        debug_assert!(false, "allocation failed (lambda params, len={param_count})");
        return Value::Error(ErrorKind::Num);
    }

    for param_expr in &args[..args.len() - 1] {
        let Some(name) = bare_identifier(param_expr) else {
            return Value::Error(ErrorKind::Value);
        };
        let name_key = match try_casefold(name.trim()) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        if params.iter().any(|p| p == &name_key) {
            return Value::Error(ErrorKind::Value);
        }
        params.push(name_key);
    }

    let Some(body) = args.last() else {
        debug_assert!(false, "checked args is non-empty");
        return Value::Error(ErrorKind::Value);
    };
    let mut env = ctx.capture_lexical_env();
    env.retain(|k, _| !k.starts_with(LAMBDA_OMITTED_PREFIX));

    Value::Lambda(Lambda {
        params: params.into(),
        body: Arc::new(body.clone()),
        env: Arc::new(env),
    })
}

inventory::submit! {
    FunctionSpec {
        name: "ISOMITTED",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Bool,
        arg_types: &[ValueType::Any],
        implementation: isomitted_fn,
    }
}

fn isomitted_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let Some(name) = bare_identifier(&args[0]) else {
        return Value::Error(ErrorKind::Value);
    };

    let Some(key) = with_casefolded_key(name.trim(), |folded| {
        let mut key = String::new();
        let len = LAMBDA_OMITTED_PREFIX.len() + folded.len();
        if key.try_reserve_exact(len).is_err() {
            debug_assert!(false, "allocation failed (isomitted key, len={len})");
            return None;
        }
        key.push_str(LAMBDA_OMITTED_PREFIX);
        key.push_str(folded);
        Some(key)
    }) else {
        return Value::Error(ErrorKind::Num);
    };
    let env = ctx.capture_lexical_env();
    Value::Bool(matches!(env.get(&key), Some(Value::Bool(true))))
}

fn bare_identifier(expr: &CompiledExpr) -> Option<&str> {
    match expr {
        Expr::NameRef(nref) if matches!(nref.sheet, SheetReference::Current) => Some(&nref.name),
        _ => None,
    }
}

// On wasm targets, `inventory` registrations can be dropped by the linker if the object file
// contains no otherwise-referenced symbols. Referencing this function from a `#[used]` table in
// `functions/mod.rs` ensures the module (and its `inventory::submit!` entries) are retained.
#[cfg(target_arch = "wasm32")]
pub(super) fn __force_link() {}
