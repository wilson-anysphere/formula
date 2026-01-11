use crate::eval::address::CellAddr;
use crate::eval::ast::{BinaryOp, CompiledExpr, CompareOp, Expr, PostfixOp, SheetReference, UnaryOp};
use crate::date::ExcelDateSystem;
use crate::error::ExcelError;
use crate::functions::{ArgValue as FnArgValue, FunctionContext, Reference as FnReference};
use crate::value::{Array, ErrorKind, Value};
use std::cmp::Ordering;
use std::cell::RefCell;
use std::rc::Rc;

#[derive(Debug, Clone, Copy)]
pub struct EvalContext {
    pub current_sheet: usize,
    pub current_cell: CellAddr,
}

pub trait ValueResolver {
    fn sheet_exists(&self, sheet_id: usize) -> bool;
    fn get_cell_value(&self, sheet_id: usize, addr: CellAddr) -> Value;
    fn resolve_structured_ref(
        &self,
        ctx: EvalContext,
        sref: &crate::structured_refs::StructuredRef,
    ) -> Option<(usize, CellAddr, CellAddr)>;
    fn resolve_name(&self, _sheet_id: usize, _name: &str) -> Option<ResolvedName> {
        None
    }
    /// If `addr` is part of a spilled array, returns the spill origin cell.
    fn spill_origin(&self, _sheet_id: usize, _addr: CellAddr) -> Option<CellAddr> {
        None
    }
    /// If `origin` is the origin of a spilled array, returns the full spill range (inclusive).
    fn spill_range(&self, _sheet_id: usize, _origin: CellAddr) -> Option<(CellAddr, CellAddr)> {
        None
    }
}

#[derive(Debug, Clone)]
pub enum ResolvedName {
    Constant(Value),
    Expr(CompiledExpr),
}

#[derive(Debug, Clone, Copy)]
struct ResolvedRange {
    sheet_id: usize,
    start: CellAddr,
    end: CellAddr,
}

impl ResolvedRange {
    fn normalized(self) -> Self {
        let (r1, r2) = if self.start.row <= self.end.row {
            (self.start.row, self.end.row)
        } else {
            (self.end.row, self.start.row)
        };
        let (c1, c2) = if self.start.col <= self.end.col {
            (self.start.col, self.end.col)
        } else {
            (self.end.col, self.start.col)
        };
        Self {
            sheet_id: self.sheet_id,
            start: CellAddr { row: r1, col: c1 },
            end: CellAddr { row: r2, col: c2 },
        }
    }

    fn is_single_cell(self) -> bool {
        self.start == self.end
    }
}

#[derive(Debug, Clone)]
enum EvalValue {
    Scalar(Value),
    Reference(Vec<ResolvedRange>),
}

pub struct Evaluator<'a, R: ValueResolver> {
    resolver: &'a R,
    ctx: EvalContext,
    name_stack: Rc<RefCell<Vec<(usize, String)>>>,
    date_system: ExcelDateSystem,
}

impl<'a, R: ValueResolver> Evaluator<'a, R> {
    pub fn new(resolver: &'a R, ctx: EvalContext) -> Self {
        Self::new_with_date_system(resolver, ctx, ExcelDateSystem::EXCEL_1900)
    }

    pub fn new_with_date_system(
        resolver: &'a R,
        ctx: EvalContext,
        date_system: ExcelDateSystem,
    ) -> Self {
        Self {
            resolver,
            ctx,
            name_stack: Rc::new(RefCell::new(Vec::new())),
            date_system,
        }
    }

    fn with_ctx(&self, ctx: EvalContext) -> Self {
        Self {
            resolver: self.resolver,
            ctx,
            name_stack: Rc::clone(&self.name_stack),
            date_system: self.date_system,
        }
    }

    /// Evaluate a compiled AST as a scalar formula result.
    pub fn eval_formula(&self, expr: &CompiledExpr) -> Value {
        match self.eval_value(expr) {
            EvalValue::Scalar(v) => v,
            EvalValue::Reference(range) => self.deref_reference_dynamic(range),
        }
    }

    fn eval_value(&self, expr: &CompiledExpr) -> EvalValue {
        match expr {
            Expr::Number(n) => EvalValue::Scalar(Value::Number(*n)),
            Expr::Text(s) => EvalValue::Scalar(Value::Text(s.clone())),
            Expr::Bool(b) => EvalValue::Scalar(Value::Bool(*b)),
            Expr::Blank => EvalValue::Scalar(Value::Blank),
            Expr::Error(e) => EvalValue::Scalar(Value::Error(*e)),
            Expr::CellRef(r) => match self.resolve_sheet_id(&r.sheet) {
                Some(sheet_id) if self.resolver.sheet_exists(sheet_id) => {
                    EvalValue::Reference(vec![ResolvedRange {
                        sheet_id,
                        start: r.addr,
                        end: r.addr,
                    }])
                }
                _ => EvalValue::Scalar(Value::Error(ErrorKind::Ref)),
            },
            Expr::RangeRef(r) => match self.resolve_sheet_id(&r.sheet) {
                Some(sheet_id) if self.resolver.sheet_exists(sheet_id) => {
                    EvalValue::Reference(vec![ResolvedRange {
                        sheet_id,
                        start: r.start,
                        end: r.end,
                    }])
                }
                _ => EvalValue::Scalar(Value::Error(ErrorKind::Ref)),
            },
            Expr::StructuredRef(sref) => match self.resolver.resolve_structured_ref(self.ctx, sref) {
                Some((sheet_id, start, end)) if self.resolver.sheet_exists(sheet_id) => {
                    EvalValue::Reference(vec![ResolvedRange { sheet_id, start, end }])
                }
                _ => EvalValue::Scalar(Value::Error(ErrorKind::Name)),
            },
            Expr::NameRef(nref) => self.eval_name_ref(nref),
            Expr::SpillRange(inner) => {
                let v = self.eval_value(inner);
                match v {
                    EvalValue::Scalar(_) => EvalValue::Scalar(Value::Error(ErrorKind::Ref)),
                    EvalValue::Reference(mut ranges) => {
                        // Spill-range references are only well-defined for a single-cell reference.
                        if ranges.len() != 1 {
                            return EvalValue::Reference(ranges);
                        }
                        let range = ranges
                            .pop()
                            .expect("checked len() above");

                        // If `#` is applied to a multi-cell reference, treat it as a no-op.
                        if !range.is_single_cell() {
                            return EvalValue::Reference(vec![range]);
                        }

                        let sheet_id = range.sheet_id;
                        let addr = range.start;
                        let origin = self.resolver.spill_origin(sheet_id, addr).unwrap_or(addr);

                        match self.resolver.spill_range(sheet_id, origin) {
                            Some((start, end)) => {
                                EvalValue::Reference(vec![ResolvedRange { sheet_id, start, end }])
                            }
                            None => EvalValue::Reference(vec![ResolvedRange { sheet_id, start: origin, end: origin }]),
                        }
                    }
                }
            }
            Expr::Unary { op, expr } => {
                let v = self.eval_value(expr);
                let v = self.deref_eval_value_dynamic(v);
                EvalValue::Scalar(elementwise_unary(&v, |elem| numeric_unary(*op, elem)))
            }
            Expr::Postfix { op, expr } => match op {
                PostfixOp::Percent => {
                    let v = self.deref_eval_value_dynamic(self.eval_value(expr));
                    EvalValue::Scalar(elementwise_unary(&v, numeric_percent))
                }
            },
            Expr::Binary { op, left, right } => match *op {
                BinaryOp::Range | BinaryOp::Union | BinaryOp::Intersect => {
                    self.eval_reference_binary(*op, left, right)
                }
                BinaryOp::Concat => {
                    let l = self.deref_eval_value_dynamic(self.eval_value(left));
                    let r = self.deref_eval_value_dynamic(self.eval_value(right));
                    let out = elementwise_binary(&l, &r, concat_binary);
                    EvalValue::Scalar(out)
                }
                BinaryOp::Pow | BinaryOp::Add | BinaryOp::Sub | BinaryOp::Mul | BinaryOp::Div => {
                    let l = self.deref_eval_value_dynamic(self.eval_value(left));
                    let r = self.deref_eval_value_dynamic(self.eval_value(right));
                    let out = elementwise_binary(&l, &r, |a, b| numeric_binary(*op, a, b));
                    EvalValue::Scalar(out)
                }
            },
            Expr::Compare { op, left, right } => {
                let l = self.deref_eval_value_dynamic(self.eval_value(left));
                let r = self.deref_eval_value_dynamic(self.eval_value(right));
                let out = elementwise_binary(&l, &r, |a, b| excel_compare(a, b, *op));
                EvalValue::Scalar(out)
            }
            Expr::FunctionCall { name, args, .. } => {
                EvalValue::Scalar(crate::functions::call_function(self, name, args))
            }
            Expr::ImplicitIntersection(inner) => {
                let v = self.eval_value(inner);
                match v {
                    EvalValue::Scalar(v) => EvalValue::Scalar(v),
                    EvalValue::Reference(ranges) => {
                        EvalValue::Scalar(self.apply_implicit_intersection(&ranges))
                    }
                }
            }
        }
    }

    fn deref_eval_value_dynamic(&self, value: EvalValue) -> Value {
        match value {
            EvalValue::Scalar(v) => v,
            EvalValue::Reference(range) => self.deref_reference_dynamic(range),
        }
    }

    fn eval_name_ref(&self, nref: &crate::eval::NameRef<usize>) -> EvalValue {
        let Some(sheet_id) = self.resolve_sheet_id(&nref.sheet) else {
            return EvalValue::Scalar(Value::Error(ErrorKind::Ref));
        };
        if !self.resolver.sheet_exists(sheet_id) {
            return EvalValue::Scalar(Value::Error(ErrorKind::Ref));
        }

        let Some(def) = self.resolver.resolve_name(sheet_id, &nref.name) else {
            return EvalValue::Scalar(Value::Error(ErrorKind::Name));
        };

        // Prevent infinite recursion from self-referential name chains.
        let key = (sheet_id, nref.name.to_ascii_uppercase());
        {
            let mut stack = self.name_stack.borrow_mut();
            if stack.contains(&key) {
                return EvalValue::Scalar(Value::Error(ErrorKind::Name));
            }
            stack.push(key.clone());
        }

        struct NameGuard {
            stack: Rc<RefCell<Vec<(usize, String)>>>,
            key: (usize, String),
        }

        impl Drop for NameGuard {
            fn drop(&mut self) {
                let mut stack = self.stack.borrow_mut();
                let popped = stack.pop();
                debug_assert_eq!(popped.as_ref(), Some(&self.key));
            }
        }

        let _guard = NameGuard {
            stack: Rc::clone(&self.name_stack),
            key,
        };

        match def {
            ResolvedName::Constant(v) => EvalValue::Scalar(v),
            ResolvedName::Expr(expr) => {
                let evaluator = self.with_ctx(EvalContext {
                    current_sheet: sheet_id,
                    current_cell: self.ctx.current_cell,
                });
                evaluator.eval_value(&expr)
            }
        }
    }

    fn eval_scalar(&self, expr: &CompiledExpr) -> Value {
        match self.eval_value(expr) {
            EvalValue::Scalar(v) => v,
            EvalValue::Reference(ranges) => self.deref_reference_scalar(&ranges),
        }
    }

    fn resolve_sheet_id(&self, sheet: &SheetReference<usize>) -> Option<usize> {
        match sheet {
            SheetReference::Current => Some(self.ctx.current_sheet),
            SheetReference::Sheet(id) => Some(*id),
            SheetReference::External(_) => None,
        }
    }

    fn deref_reference_scalar(&self, ranges: &[ResolvedRange]) -> Value {
        match ranges {
            [only] if only.is_single_cell() => self.resolver.get_cell_value(only.sheet_id, only.start),
            _ => {
                // Multi-cell references used as scalars behave like a spill attempt.
                Value::Error(ErrorKind::Spill)
            }
        }
    }

    fn deref_reference_dynamic(&self, ranges: Vec<ResolvedRange>) -> Value {
        match ranges.as_slice() {
            [] => Value::Error(ErrorKind::Ref),
            [only] => self.deref_reference_dynamic_single(*only),
            // Discontiguous unions cannot be represented as a single rectangular spill.
            _ => Value::Error(ErrorKind::Value),
        }
    }

    fn deref_reference_dynamic_single(&self, range: ResolvedRange) -> Value {
        if range.is_single_cell() {
            return self.resolver.get_cell_value(range.sheet_id, range.start);
        }
        let range = range.normalized();
        let rows = (range.end.row - range.start.row + 1) as usize;
        let cols = (range.end.col - range.start.col + 1) as usize;
        let mut values = Vec::with_capacity(rows.saturating_mul(cols));
        for row in range.start.row..=range.end.row {
            for col in range.start.col..=range.end.col {
                values.push(self.resolver.get_cell_value(range.sheet_id, CellAddr { row, col }));
            }
        }
        Value::Array(Array::new(rows, cols, values))
    }

    fn apply_implicit_intersection(&self, ranges: &[ResolvedRange]) -> Value {
        match ranges {
            [] => Value::Error(ErrorKind::Value),
            [only] => self.apply_implicit_intersection_single(*only),
            many => {
                // If multiple areas intersect, Excel's implicit intersection is ambiguous. We
                // approximate by succeeding only when exactly one area intersects.
                let mut hits = Vec::new();
                for r in many {
                    let v = self.apply_implicit_intersection_single(*r);
                    if !matches!(v, Value::Error(ErrorKind::Value)) {
                        hits.push(v);
                    }
                }
                match hits.as_slice() {
                    [only] => only.clone(),
                    _ => Value::Error(ErrorKind::Value),
                }
            }
        }
    }

    fn apply_implicit_intersection_single(&self, range: ResolvedRange) -> Value {
        if range.is_single_cell() {
            return self.resolver.get_cell_value(range.sheet_id, range.start);
        }

        let range = range.normalized();
        let cur = self.ctx.current_cell;

        // 1D ranges intersect on the matching row/column.
        if range.start.col == range.end.col {
            if cur.row >= range.start.row && cur.row <= range.end.row {
                return self.resolver.get_cell_value(
                    range.sheet_id,
                    CellAddr {
                        row: cur.row,
                        col: range.start.col,
                    },
                );
            }
            return Value::Error(ErrorKind::Value);
        }
        if range.start.row == range.end.row {
            if cur.col >= range.start.col && cur.col <= range.end.col {
                return self.resolver.get_cell_value(
                    range.sheet_id,
                    CellAddr {
                        row: range.start.row,
                        col: cur.col,
                    },
                );
            }
            return Value::Error(ErrorKind::Value);
        }

        // 2D ranges intersect only if the current cell is within the rectangle.
        if cur.row >= range.start.row
            && cur.row <= range.end.row
            && cur.col >= range.start.col
            && cur.col <= range.end.col
        {
            return self.resolver.get_cell_value(range.sheet_id, cur);
        }

        Value::Error(ErrorKind::Value)
    }

    fn eval_reference_binary(&self, op: BinaryOp, left: &CompiledExpr, right: &CompiledExpr) -> EvalValue {
        let left = match self.eval_reference_operand(left) {
            Ok(r) => r,
            Err(v) => return EvalValue::Scalar(v),
        };
        let right = match self.eval_reference_operand(right) {
            Ok(r) => r,
            Err(v) => return EvalValue::Scalar(v),
        };

        match op {
            BinaryOp::Union => {
                let Some(sheet_id) = left.first().map(|r| r.sheet_id) else {
                    return EvalValue::Scalar(Value::Error(ErrorKind::Ref));
                };
                if left.iter().any(|r| r.sheet_id != sheet_id) || right.iter().any(|r| r.sheet_id != sheet_id) {
                    return EvalValue::Scalar(Value::Error(ErrorKind::Ref));
                }

                let mut out = left;
                out.extend(right);
                EvalValue::Reference(out)
            }
            BinaryOp::Intersect => {
                let mut out = Vec::new();
                for a in &left {
                    for b in &right {
                        if a.sheet_id != b.sheet_id {
                            return EvalValue::Scalar(Value::Error(ErrorKind::Ref));
                        }
                        if let Some(r) = intersect_ranges(*a, *b) {
                            out.push(r);
                        }
                    }
                }
                if out.is_empty() {
                    return EvalValue::Scalar(Value::Error(ErrorKind::Null));
                }
                EvalValue::Reference(out)
            }
            BinaryOp::Range => {
                let (Some(a), Some(b)) = (left.first().copied(), right.first().copied()) else {
                    return EvalValue::Scalar(Value::Error(ErrorKind::Ref));
                };
                if left.len() != 1 || right.len() != 1 {
                    return EvalValue::Scalar(Value::Error(ErrorKind::Value));
                }
                if a.sheet_id != b.sheet_id {
                    return EvalValue::Scalar(Value::Error(ErrorKind::Ref));
                }

                let a = a.normalized();
                let b = b.normalized();

                let start = CellAddr {
                    row: a.start.row.min(b.start.row),
                    col: a.start.col.min(b.start.col),
                };
                let end = CellAddr {
                    row: a.end.row.max(b.end.row),
                    col: a.end.col.max(b.end.col),
                };

                EvalValue::Reference(vec![ResolvedRange {
                    sheet_id: a.sheet_id,
                    start,
                    end,
                }])
            }
            _ => EvalValue::Scalar(Value::Error(ErrorKind::Value)),
        }
    }

    fn eval_reference_operand(&self, expr: &CompiledExpr) -> Result<Vec<ResolvedRange>, Value> {
        match self.eval_value(expr) {
            EvalValue::Reference(r) => Ok(r),
            EvalValue::Scalar(Value::Error(e)) => Err(Value::Error(e)),
            EvalValue::Scalar(_) => Err(Value::Error(ErrorKind::Value)),
        }
    }

    // Built-in functions are implemented in `crate::functions` and dispatched via
    // `crate::functions::call_function`.
}

fn intersect_ranges(a: ResolvedRange, b: ResolvedRange) -> Option<ResolvedRange> {
    if a.sheet_id != b.sheet_id {
        return None;
    }
    let a = a.normalized();
    let b = b.normalized();

    let start_row = a.start.row.max(b.start.row);
    let end_row = a.end.row.min(b.end.row);
    if start_row > end_row {
        return None;
    }
    let start_col = a.start.col.max(b.start.col);
    let end_col = a.end.col.min(b.end.col);
    if start_col > end_col {
        return None;
    }

    Some(ResolvedRange {
        sheet_id: a.sheet_id,
        start: CellAddr {
            row: start_row,
            col: start_col,
        },
        end: CellAddr {
            row: end_row,
            col: end_col,
        },
    })
}

impl<'a, R: ValueResolver> FunctionContext for Evaluator<'a, R> {
    fn eval_arg(&self, expr: &CompiledExpr) -> FnArgValue {
        match self.eval_value(expr) {
            EvalValue::Scalar(v) => FnArgValue::Scalar(v),
            EvalValue::Reference(mut ranges) => {
                // Ensure a stable order for deterministic function behavior (e.g. COUNT over a
                // multi-area union).
                ranges.sort_by_key(|r| (r.sheet_id, r.start.row, r.start.col, r.end.row, r.end.col));
                match ranges.as_slice() {
                    [only] => FnArgValue::Reference(FnReference {
                        sheet_id: only.sheet_id,
                        start: only.start,
                        end: only.end,
                    }),
                    _ => FnArgValue::ReferenceUnion(
                        ranges
                            .into_iter()
                            .map(|r| FnReference {
                                sheet_id: r.sheet_id,
                                start: r.start,
                                end: r.end,
                            })
                            .collect(),
                    ),
                }
            }
        }
    }

    fn eval_scalar(&self, expr: &CompiledExpr) -> Value {
        Evaluator::eval_scalar(self, expr)
    }

    fn apply_implicit_intersection(&self, reference: FnReference) -> Value {
        Evaluator::apply_implicit_intersection(self, &[ResolvedRange {
            sheet_id: reference.sheet_id,
            start: reference.start,
            end: reference.end,
        }])
    }

    fn get_cell_value(&self, sheet_id: usize, addr: CellAddr) -> Value {
        self.resolver.get_cell_value(sheet_id, addr)
    }

    fn now_utc(&self) -> chrono::DateTime<chrono::Utc> {
        chrono::Utc::now()
    }

    fn date_system(&self) -> ExcelDateSystem {
        self.date_system
    }
}

fn excel_compare(left: &Value, right: &Value, op: CompareOp) -> Value {
    let ord = match excel_order(left, right) {
        Ok(ord) => ord,
        Err(e) => return Value::Error(e),
    };

    let result = match op {
        CompareOp::Eq => ord == Ordering::Equal,
        CompareOp::Ne => ord != Ordering::Equal,
        CompareOp::Lt => ord == Ordering::Less,
        CompareOp::Le => ord != Ordering::Greater,
        CompareOp::Gt => ord == Ordering::Greater,
        CompareOp::Ge => ord != Ordering::Less,
    };

    Value::Bool(result)
}

fn excel_order(left: &Value, right: &Value) -> Result<Ordering, ErrorKind> {
    if let Value::Error(e) = left {
        return Err(*e);
    }
    if let Value::Error(e) = right {
        return Err(*e);
    }
    if matches!(left, Value::Array(_) | Value::Spill { .. })
        || matches!(right, Value::Array(_) | Value::Spill { .. })
    {
        return Err(ErrorKind::Value);
    }

    // Blank coerces to the other type for comparisons.
    let (l, r) = match (left, right) {
        (Value::Blank, Value::Number(_)) => (Value::Number(0.0), right.clone()),
        (Value::Number(_), Value::Blank) => (left.clone(), Value::Number(0.0)),
        (Value::Blank, Value::Bool(_)) => (Value::Bool(false), right.clone()),
        (Value::Bool(_), Value::Blank) => (left.clone(), Value::Bool(false)),
        (Value::Blank, Value::Text(_)) => (Value::Text(String::new()), right.clone()),
        (Value::Text(_), Value::Blank) => (left.clone(), Value::Text(String::new())),
        _ => (left.clone(), right.clone()),
    };

    Ok(match (&l, &r) {
        (Value::Number(a), Value::Number(b)) => a.partial_cmp(b).unwrap_or(Ordering::Equal),
        (Value::Text(a), Value::Text(b)) => {
            let au = a.to_ascii_uppercase();
            let bu = b.to_ascii_uppercase();
            au.cmp(&bu)
        }
        (Value::Bool(a), Value::Bool(b)) => a.cmp(b),
        // Type precedence (approximate Excel): numbers < text < booleans.
        (Value::Number(_), Value::Text(_) | Value::Bool(_)) => Ordering::Less,
        (Value::Text(_), Value::Bool(_)) => Ordering::Less,
        (Value::Text(_), Value::Number(_)) => Ordering::Greater,
        (Value::Bool(_), Value::Number(_) | Value::Text(_)) => Ordering::Greater,
        // Blank should have been coerced above.
        (Value::Blank, Value::Blank) => Ordering::Equal,
        (Value::Blank, _) => Ordering::Less,
        (_, Value::Blank) => Ordering::Greater,
        // Errors are handled above.
        (Value::Error(_), _) | (_, Value::Error(_)) => Ordering::Equal,
        // Arrays/spill markers are rejected above.
        (Value::Array(_), _)
        | (_, Value::Array(_))
        | (Value::Spill { .. }, _)
        | (_, Value::Spill { .. }) => Ordering::Equal,
    })
}

fn numeric_unary(op: UnaryOp, value: &Value) -> Value {
    match value {
        Value::Error(e) => return Value::Error(*e),
        other => {
            let n = match other.coerce_to_number() {
                Ok(n) => n,
                Err(e) => return Value::Error(e),
            };
            let out = match op {
                UnaryOp::Plus => n,
                UnaryOp::Minus => -n,
            };
            Value::Number(out)
        }
    }
}

fn numeric_percent(value: &Value) -> Value {
    match value {
        Value::Error(e) => return Value::Error(*e),
        other => {
            let n = match other.coerce_to_number() {
                Ok(n) => n,
                Err(e) => return Value::Error(e),
            };
            Value::Number(n / 100.0)
        }
    }
}

fn concat_binary(left: &Value, right: &Value) -> Value {
    if let Value::Error(e) = left {
        return Value::Error(*e);
    }
    if let Value::Error(e) = right {
        return Value::Error(*e);
    }

    let ls = match left.coerce_to_string() {
        Ok(s) => s,
        Err(e) => return Value::Error(e),
    };
    let rs = match right.coerce_to_string() {
        Ok(s) => s,
        Err(e) => return Value::Error(e),
    };

    Value::Text(format!("{ls}{rs}"))
}

fn numeric_binary(op: BinaryOp, left: &Value, right: &Value) -> Value {
    if let Value::Error(e) = left {
        return Value::Error(*e);
    }
    if let Value::Error(e) = right {
        return Value::Error(*e);
    }

    let ln = match left.coerce_to_number() {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let rn = match right.coerce_to_number() {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };

    match op {
        BinaryOp::Add => Value::Number(ln + rn),
        BinaryOp::Sub => Value::Number(ln - rn),
        BinaryOp::Mul => Value::Number(ln * rn),
        BinaryOp::Div => {
            if rn == 0.0 {
                Value::Error(ErrorKind::Div0)
            } else {
                Value::Number(ln / rn)
            }
        }
        BinaryOp::Pow => match crate::functions::math::power(ln, rn) {
            Ok(n) => Value::Number(n),
            Err(e) => Value::Error(match e {
                ExcelError::Div0 => ErrorKind::Div0,
                ExcelError::Value => ErrorKind::Value,
                ExcelError::Num => ErrorKind::Num,
            }),
        },
        _ => Value::Error(ErrorKind::Value),
    }
}

fn elementwise_unary(value: &Value, f: impl Fn(&Value) -> Value) -> Value {
    match value {
        Value::Array(arr) => Value::Array(Array::new(
            arr.rows,
            arr.cols,
            arr.iter().map(f).collect(),
        )),
        other => f(other),
    }
}

fn elementwise_binary(left: &Value, right: &Value, f: impl Fn(&Value, &Value) -> Value) -> Value {
    match (left, right) {
        (Value::Array(left_arr), Value::Array(right_arr)) => {
            if left_arr.rows == right_arr.rows && left_arr.cols == right_arr.cols {
                return Value::Array(Array::new(
                    left_arr.rows,
                    left_arr.cols,
                    left_arr
                        .values
                        .iter()
                        .zip(right_arr.values.iter())
                        .map(|(a, b)| f(a, b))
                        .collect(),
                ));
            }

            if left_arr.rows == 1 && left_arr.cols == 1 {
                let scalar = left_arr.values.get(0).unwrap_or(&Value::Blank);
                return Value::Array(Array::new(
                    right_arr.rows,
                    right_arr.cols,
                    right_arr.values.iter().map(|b| f(scalar, b)).collect(),
                ));
            }

            if right_arr.rows == 1 && right_arr.cols == 1 {
                let scalar = right_arr.values.get(0).unwrap_or(&Value::Blank);
                return Value::Array(Array::new(
                    left_arr.rows,
                    left_arr.cols,
                    left_arr.values.iter().map(|a| f(a, scalar)).collect(),
                ));
            }

            Value::Error(ErrorKind::Value)
        }
        (Value::Array(left_arr), right_scalar) => Value::Array(Array::new(
            left_arr.rows,
            left_arr.cols,
            left_arr.values.iter().map(|a| f(a, right_scalar)).collect(),
        )),
        (left_scalar, Value::Array(right_arr)) => Value::Array(Array::new(
            right_arr.rows,
            right_arr.cols,
            right_arr.values.iter().map(|b| f(left_scalar, b)).collect(),
        )),
        (left_scalar, right_scalar) => f(left_scalar, right_scalar),
    }
}
