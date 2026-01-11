use std::collections::HashSet;
use std::sync::Arc;

use crate::eval::{CompiledExpr, Expr, SheetReference, LAMBDA_OMITTED_PREFIX};
use crate::functions::{
    ArgValue, ArraySupport, FunctionContext, FunctionSpec, ThreadSafety, ValueType, Volatility,
};
use crate::value::{ErrorKind, Lambda, Value};

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
        let name_key = name.trim().to_ascii_uppercase();

        let value = ctx.eval_arg(&pair[1]);
        if let ArgValue::Scalar(Value::Error(e)) = &value {
            return Value::Error(*e);
        }

        ctx.set_local(&name_key, value);
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

    let mut params: Vec<String> = Vec::with_capacity(args.len().saturating_sub(1));
    let mut seen: HashSet<String> = HashSet::new();

    for param_expr in &args[..args.len() - 1] {
        let Some(name) = bare_identifier(param_expr) else {
            return Value::Error(ErrorKind::Value);
        };
        let name_key = name.trim().to_ascii_uppercase();
        if !seen.insert(name_key.clone()) {
            return Value::Error(ErrorKind::Value);
        }
        params.push(name_key);
    }

    let body = args.last().expect("checked args is non-empty");
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

    let key = format!(
        "{LAMBDA_OMITTED_PREFIX}{}",
        name.trim().to_ascii_uppercase()
    );
    let env = ctx.capture_lexical_env();
    Value::Bool(matches!(env.get(&key), Some(Value::Bool(true))))
}

fn bare_identifier(expr: &CompiledExpr) -> Option<&str> {
    match expr {
        Expr::NameRef(nref) if matches!(nref.sheet, SheetReference::Current) => Some(&nref.name),
        _ => None,
    }
}
