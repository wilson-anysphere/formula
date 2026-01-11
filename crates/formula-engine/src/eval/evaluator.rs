use crate::eval::address::CellAddr;
use crate::eval::ast::{BinaryOp, CompiledExpr, CompareOp, Expr, SheetReference, UnaryOp};
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
    Reference(ResolvedRange),
}

pub struct Evaluator<'a, R: ValueResolver> {
    resolver: &'a R,
    ctx: EvalContext,
    name_stack: Rc<RefCell<Vec<(usize, String)>>>,
}

impl<'a, R: ValueResolver> Evaluator<'a, R> {
    pub fn new(resolver: &'a R, ctx: EvalContext) -> Self {
        Self {
            resolver,
            ctx,
            name_stack: Rc::new(RefCell::new(Vec::new())),
        }
    }

    fn with_ctx(&self, ctx: EvalContext) -> Self {
        Self {
            resolver: self.resolver,
            ctx,
            name_stack: Rc::clone(&self.name_stack),
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
                    EvalValue::Reference(ResolvedRange {
                        sheet_id,
                        start: r.addr,
                        end: r.addr,
                    })
                }
                _ => EvalValue::Scalar(Value::Error(ErrorKind::Ref)),
            },
            Expr::RangeRef(r) => match self.resolve_sheet_id(&r.sheet) {
                Some(sheet_id) if self.resolver.sheet_exists(sheet_id) => {
                    EvalValue::Reference(ResolvedRange {
                        sheet_id,
                        start: r.start,
                        end: r.end,
                    })
                }
                _ => EvalValue::Scalar(Value::Error(ErrorKind::Ref)),
            },
            Expr::StructuredRef(sref) => match self.resolver.resolve_structured_ref(self.ctx, sref) {
                Some((sheet_id, start, end)) if self.resolver.sheet_exists(sheet_id) => {
                    EvalValue::Reference(ResolvedRange { sheet_id, start, end })
                }
                _ => EvalValue::Scalar(Value::Error(ErrorKind::Name)),
            },
            Expr::NameRef(nref) => self.eval_name_ref(nref),
            Expr::Unary { op, expr } => {
                let v = self.eval_scalar(expr);
                match v {
                    Value::Error(e) => EvalValue::Scalar(Value::Error(e)),
                    other => {
                        let n = match other.coerce_to_number() {
                            Ok(n) => n,
                            Err(e) => return EvalValue::Scalar(Value::Error(e)),
                        };
                        let out = match op {
                            UnaryOp::Plus => n,
                            UnaryOp::Minus => -n,
                        };
                        EvalValue::Scalar(Value::Number(out))
                    }
                }
            }
            Expr::Binary { op, left, right } => {
                let l = self.eval_scalar(left);
                if let Value::Error(e) = l {
                    return EvalValue::Scalar(Value::Error(e));
                }
                let r = self.eval_scalar(right);
                if let Value::Error(e) = r {
                    return EvalValue::Scalar(Value::Error(e));
                }
                let ln = match l.coerce_to_number() {
                    Ok(n) => n,
                    Err(e) => return EvalValue::Scalar(Value::Error(e)),
                };
                let rn = match r.coerce_to_number() {
                    Ok(n) => n,
                    Err(e) => return EvalValue::Scalar(Value::Error(e)),
                };
                let out = match op {
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
                };
                EvalValue::Scalar(out)
            }
            Expr::Compare { op, left, right } => {
                let l = self.eval_scalar(left);
                if let Value::Error(e) = l {
                    return EvalValue::Scalar(Value::Error(e));
                }
                let r = self.eval_scalar(right);
                if let Value::Error(e) = r {
                    return EvalValue::Scalar(Value::Error(e));
                }

                let b = excel_compare(&l, &r, *op);
                EvalValue::Scalar(b)
            }
            Expr::FunctionCall { name, args, .. } => {
                EvalValue::Scalar(crate::functions::call_function(self, name, args))
            }
            Expr::ImplicitIntersection(inner) => {
                let v = self.eval_value(inner);
                match v {
                    EvalValue::Scalar(v) => EvalValue::Scalar(v),
                    EvalValue::Reference(range) => {
                        EvalValue::Scalar(self.apply_implicit_intersection(range))
                    }
                }
            }
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
            EvalValue::Reference(range) => self.deref_reference_scalar(range),
        }
    }

    fn resolve_sheet_id(&self, sheet: &SheetReference<usize>) -> Option<usize> {
        match sheet {
            SheetReference::Current => Some(self.ctx.current_sheet),
            SheetReference::Sheet(id) => Some(*id),
            SheetReference::External(_) => None,
        }
    }

    fn deref_reference_scalar(&self, range: ResolvedRange) -> Value {
        if range.is_single_cell() {
            self.resolver.get_cell_value(range.sheet_id, range.start)
        } else {
            // Dynamic array spilling is not implemented yet; multi-cell references used as
            // scalars behave like a spill attempt.
            Value::Error(ErrorKind::Spill)
        }
    }

    fn deref_reference_dynamic(&self, range: ResolvedRange) -> Value {
        if range.is_single_cell() {
            return self.resolver.get_cell_value(range.sheet_id, range.start);
        }
        let range = range.normalized();
        let rows = (range.end.row - range.start.row + 1) as usize;
        let cols = (range.end.col - range.start.col + 1) as usize;
        let mut values = Vec::with_capacity(rows.saturating_mul(cols));
        for row in range.start.row..=range.end.row {
            for col in range.start.col..=range.end.col {
                values.push(
                    self.resolver
                        .get_cell_value(range.sheet_id, CellAddr { row, col }),
                );
            }
        }
        Value::Array(Array::new(rows, cols, values))
    }

    fn apply_implicit_intersection(&self, range: ResolvedRange) -> Value {
        if range.is_single_cell() {
            return self.resolver.get_cell_value(range.sheet_id, range.start);
        }

        let range = range.normalized();
        let cur = self.ctx.current_cell;

        // 1D ranges intersect on the matching row/column.
        if range.start.col == range.end.col {
            if cur.row >= range.start.row && cur.row <= range.end.row {
                return self
                    .resolver
                    .get_cell_value(range.sheet_id, CellAddr { row: cur.row, col: range.start.col });
            }
            return Value::Error(ErrorKind::Value);
        }
        if range.start.row == range.end.row {
            if cur.col >= range.start.col && cur.col <= range.end.col {
                return self
                    .resolver
                    .get_cell_value(range.sheet_id, CellAddr { row: range.start.row, col: cur.col });
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

    // Built-in functions are implemented in `crate::functions` and dispatched via
    // `crate::functions::call_function`.
}

impl<'a, R: ValueResolver> FunctionContext for Evaluator<'a, R> {
    fn eval_arg(&self, expr: &CompiledExpr) -> FnArgValue {
        match self.eval_value(expr) {
            EvalValue::Scalar(v) => FnArgValue::Scalar(v),
            EvalValue::Reference(r) => FnArgValue::Reference(FnReference {
                sheet_id: r.sheet_id,
                start: r.start,
                end: r.end,
            }),
        }
    }

    fn eval_scalar(&self, expr: &CompiledExpr) -> Value {
        Evaluator::eval_scalar(self, expr)
    }

    fn apply_implicit_intersection(&self, reference: FnReference) -> Value {
        let range = ResolvedRange {
            sheet_id: reference.sheet_id,
            start: reference.start,
            end: reference.end,
        };
        Evaluator::apply_implicit_intersection(self, range)
    }

    fn get_cell_value(&self, sheet_id: usize, addr: CellAddr) -> Value {
        self.resolver.get_cell_value(sheet_id, addr)
    }

    fn now_utc(&self) -> chrono::DateTime<chrono::Utc> {
        chrono::Utc::now()
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
