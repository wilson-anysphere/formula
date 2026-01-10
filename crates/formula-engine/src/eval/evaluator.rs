use crate::error::ExcelError;
use crate::eval::address::CellAddr;
use crate::eval::ast::{BinaryOp, CompareOp, CompiledExpr, Expr, SheetReference, UnaryOp};
use crate::functions::financial;
use crate::value::{ErrorKind, Value};
use std::cmp::Ordering;

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

    fn iter_cells(self) -> impl Iterator<Item = CellAddr> {
        let norm = self.normalized();
        let rows = norm.start.row..=norm.end.row;
        let cols = norm.start.col..=norm.end.col;
        rows.flat_map(move |row| cols.clone().map(move |col| CellAddr { row, col }))
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
}

impl<'a, R: ValueResolver> Evaluator<'a, R> {
    pub fn new(resolver: &'a R, ctx: EvalContext) -> Self {
        Self { resolver, ctx }
    }

    /// Evaluate a compiled AST as a scalar formula result.
    pub fn eval_formula(&self, expr: &CompiledExpr) -> Value {
        self.eval_scalar(expr)
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
            Expr::Unary { op, expr } => {
                let v = self.eval_scalar(expr);
                match v {
                    Value::Error(e) => EvalValue::Scalar(Value::Error(e)),
                    other => {
                        let n = match coerce_to_number(&other) {
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
                let ln = match coerce_to_number(&l) {
                    Ok(n) => n,
                    Err(e) => return EvalValue::Scalar(Value::Error(e)),
                };
                let rn = match coerce_to_number(&r) {
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
            Expr::FunctionCall { name, args } => EvalValue::Scalar(self.eval_function(name, args)),
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

    fn apply_implicit_intersection(&self, range: ResolvedRange) -> Value {
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

    fn eval_function(&self, name: &str, args: &[CompiledExpr]) -> Value {
        match name {
            "IF" => self.fn_if(args),
            "IFERROR" => self.fn_iferror(args),
            "ISERROR" => self.fn_iserror(args),
            "SUM" => self.fn_sum(args),
            "PV" => self.fn_pv(args),
            "FV" => self.fn_fv(args),
            "PMT" => self.fn_pmt(args),
            "NPER" => self.fn_nper(args),
            "RATE" => self.fn_rate(args),
            "IPMT" => self.fn_ipmt(args),
            "PPMT" => self.fn_ppmt(args),
            "SLN" => self.fn_sln(args),
            "SYD" => self.fn_syd(args),
            "DDB" => self.fn_ddb(args),
            "NPV" => self.fn_npv(args),
            "IRR" => self.fn_irr(args),
            "XNPV" => self.fn_xnpv(args),
            "XIRR" => self.fn_xirr(args),
            _ => Value::Error(ErrorKind::Name),
        }
    }

    fn fn_if(&self, args: &[CompiledExpr]) -> Value {
        if args.is_empty() {
            return Value::Error(ErrorKind::Value);
        }
        let cond_val = self.eval_scalar(&args[0]);
        if let Value::Error(e) = cond_val {
            return Value::Error(e);
        }
        let cond = match coerce_to_bool(&cond_val) {
            Ok(b) => b,
            Err(e) => return Value::Error(e),
        };

        if cond {
            if args.len() >= 2 {
                self.eval_scalar(&args[1])
            } else {
                Value::Bool(true)
            }
        } else if args.len() >= 3 {
            self.eval_scalar(&args[2])
        } else {
            Value::Bool(false)
        }
    }

    fn fn_iferror(&self, args: &[CompiledExpr]) -> Value {
        if args.len() < 2 {
            return Value::Error(ErrorKind::Value);
        }
        let first = self.eval_scalar(&args[0]);
        match first {
            Value::Error(_) => self.eval_scalar(&args[1]),
            other => other,
        }
    }

    fn fn_iserror(&self, args: &[CompiledExpr]) -> Value {
        if args.len() != 1 {
            return Value::Error(ErrorKind::Value);
        }
        let v = self.eval_scalar(&args[0]);
        Value::Bool(matches!(v, Value::Error(_)))
    }

    fn fn_sum(&self, args: &[CompiledExpr]) -> Value {
        let mut acc = 0.0;

        for arg in args {
            let ev = self.eval_value(arg);
            match ev {
                EvalValue::Scalar(v) => match v {
                    Value::Error(e) => return Value::Error(e),
                    Value::Number(n) => acc += n,
                    Value::Bool(b) => acc += if b { 1.0 } else { 0.0 },
                    Value::Blank => {}
                    Value::Text(s) => {
                        if let Some(n) = parse_number_from_text(&s) {
                            acc += n;
                        }
                    }
                },
                EvalValue::Reference(range) => {
                    for addr in range.iter_cells() {
                        let v = self.resolver.get_cell_value(range.sheet_id, addr);
                        match v {
                            Value::Error(e) => return Value::Error(e),
                            Value::Number(n) => acc += n,
                            // Excel quirk: logicals/text in references are ignored by SUM.
                            Value::Bool(_) | Value::Text(_) | Value::Blank => {}
                        }
                    }
                }
            }
        }

        Value::Number(acc)
    }

    fn eval_number_arg(&self, expr: &CompiledExpr) -> Result<f64, ErrorKind> {
        let v = self.eval_scalar(expr);
        match v {
            Value::Error(e) => Err(e),
            other => coerce_to_number(&other),
        }
    }

    fn eval_optional_number_arg(
        &self,
        expr: Option<&CompiledExpr>,
    ) -> Result<Option<f64>, ErrorKind> {
        match expr {
            Some(e) => Ok(Some(self.eval_number_arg(e)?)),
            None => Ok(None),
        }
    }

    fn excel_result_number(res: Result<f64, ExcelError>) -> Value {
        match res {
            Ok(n) => Value::Number(n),
            Err(e) => Value::Error(match e {
                ExcelError::Div0 => ErrorKind::Div0,
                ExcelError::Value => ErrorKind::Value,
                ExcelError::Num => ErrorKind::Num,
            }),
        }
    }

    fn collect_numbers_strict_from_arg(&self, arg: &CompiledExpr) -> Result<Vec<f64>, ErrorKind> {
        let ev = self.eval_value(arg);
        match ev {
            EvalValue::Scalar(v) => Ok(vec![coerce_to_number(&v)?]),
            EvalValue::Reference(range) => {
                let mut out = Vec::new();
                for addr in range.iter_cells() {
                    let v = self.resolver.get_cell_value(range.sheet_id, addr);
                    out.push(coerce_to_number(&v)?);
                }
                Ok(out)
            }
        }
    }

    fn collect_npv_values_from_arg(&self, arg: &CompiledExpr) -> Result<Vec<f64>, ErrorKind> {
        // NPV's discounting depends on the position of each cashflow. When values are
        // supplied via references, Excel ignores non-numeric cells. For NPV we treat
        // them as 0 so the period index is preserved (i.e. a blank cell represents a
        // period with no cashflow rather than "removing" a period).
        let ev = self.eval_value(arg);
        match ev {
            EvalValue::Scalar(v) => match v {
                Value::Error(e) => Err(e),
                Value::Number(n) => Ok(vec![n]),
                Value::Bool(b) => Ok(vec![if b { 1.0 } else { 0.0 }]),
                Value::Blank => Ok(vec![0.0]),
                Value::Text(s) => Ok(vec![parse_number_from_text(&s).unwrap_or(0.0)]),
            },
            EvalValue::Reference(range) => {
                let mut out = Vec::new();
                for addr in range.iter_cells() {
                    let v = self.resolver.get_cell_value(range.sheet_id, addr);
                    match v {
                        Value::Error(e) => return Err(e),
                        Value::Number(n) => out.push(n),
                        Value::Bool(_) | Value::Text(_) | Value::Blank => out.push(0.0),
                    }
                }
                Ok(out)
            }
        }
    }

    fn collect_irr_values_from_arg(&self, arg: &CompiledExpr) -> Result<Vec<f64>, ErrorKind> {
        // Excel IRR accepts references that contain non-numeric cells; text/logical/blank
        // entries are ignored (i.e. contribute 0) but their position still counts as a period.
        // Errors propagate.
        let ev = self.eval_value(arg);
        match ev {
            EvalValue::Scalar(v) => match v {
                Value::Error(e) => Err(e),
                Value::Number(n) => Ok(vec![n]),
                Value::Bool(_) | Value::Text(_) | Value::Blank => Ok(vec![0.0]),
            },
            EvalValue::Reference(range) => {
                let mut out = Vec::new();
                for addr in range.iter_cells() {
                    let v = self.resolver.get_cell_value(range.sheet_id, addr);
                    match v {
                        Value::Error(e) => return Err(e),
                        Value::Number(n) => out.push(n),
                        Value::Bool(_) | Value::Text(_) | Value::Blank => out.push(0.0),
                    }
                }
                Ok(out)
            }
        }
    }

    fn fn_pv(&self, args: &[CompiledExpr]) -> Value {
        if args.len() < 3 || args.len() > 5 {
            return Value::Error(ErrorKind::Value);
        }
        let rate = match self.eval_number_arg(&args[0]) {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        let nper = match self.eval_number_arg(&args[1]) {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        let pmt = match self.eval_number_arg(&args[2]) {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        let fv = match self.eval_optional_number_arg(args.get(3)) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        let typ = match self.eval_optional_number_arg(args.get(4)) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        Self::excel_result_number(financial::pv(rate, nper, pmt, fv, typ))
    }

    fn fn_fv(&self, args: &[CompiledExpr]) -> Value {
        if args.len() < 3 || args.len() > 5 {
            return Value::Error(ErrorKind::Value);
        }
        let rate = match self.eval_number_arg(&args[0]) {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        let nper = match self.eval_number_arg(&args[1]) {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        let pmt = match self.eval_number_arg(&args[2]) {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        let pv = match self.eval_optional_number_arg(args.get(3)) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        let typ = match self.eval_optional_number_arg(args.get(4)) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        Self::excel_result_number(financial::fv(rate, nper, pmt, pv, typ))
    }

    fn fn_pmt(&self, args: &[CompiledExpr]) -> Value {
        if args.len() < 3 || args.len() > 5 {
            return Value::Error(ErrorKind::Value);
        }
        let rate = match self.eval_number_arg(&args[0]) {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        let nper = match self.eval_number_arg(&args[1]) {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        let pv = match self.eval_number_arg(&args[2]) {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        let fv = match self.eval_optional_number_arg(args.get(3)) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        let typ = match self.eval_optional_number_arg(args.get(4)) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        Self::excel_result_number(financial::pmt(rate, nper, pv, fv, typ))
    }

    fn fn_nper(&self, args: &[CompiledExpr]) -> Value {
        if args.len() < 3 || args.len() > 5 {
            return Value::Error(ErrorKind::Value);
        }
        let rate = match self.eval_number_arg(&args[0]) {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        let pmt = match self.eval_number_arg(&args[1]) {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        let pv = match self.eval_number_arg(&args[2]) {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        let fv = match self.eval_optional_number_arg(args.get(3)) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        let typ = match self.eval_optional_number_arg(args.get(4)) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        Self::excel_result_number(financial::nper(rate, pmt, pv, fv, typ))
    }

    fn fn_rate(&self, args: &[CompiledExpr]) -> Value {
        if args.len() < 3 || args.len() > 6 {
            return Value::Error(ErrorKind::Value);
        }
        let nper = match self.eval_number_arg(&args[0]) {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        let pmt = match self.eval_number_arg(&args[1]) {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        let pv = match self.eval_number_arg(&args[2]) {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        let fv = match self.eval_optional_number_arg(args.get(3)) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        let typ = match self.eval_optional_number_arg(args.get(4)) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        let guess = match self.eval_optional_number_arg(args.get(5)) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        Self::excel_result_number(financial::rate(nper, pmt, pv, fv, typ, guess))
    }

    fn fn_ipmt(&self, args: &[CompiledExpr]) -> Value {
        if args.len() < 4 || args.len() > 6 {
            return Value::Error(ErrorKind::Value);
        }
        let rate = match self.eval_number_arg(&args[0]) {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        let per = match self.eval_number_arg(&args[1]) {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        let nper = match self.eval_number_arg(&args[2]) {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        let pv = match self.eval_number_arg(&args[3]) {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        let fv = match self.eval_optional_number_arg(args.get(4)) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        let typ = match self.eval_optional_number_arg(args.get(5)) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        Self::excel_result_number(financial::ipmt(rate, per, nper, pv, fv, typ))
    }

    fn fn_ppmt(&self, args: &[CompiledExpr]) -> Value {
        if args.len() < 4 || args.len() > 6 {
            return Value::Error(ErrorKind::Value);
        }
        let rate = match self.eval_number_arg(&args[0]) {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        let per = match self.eval_number_arg(&args[1]) {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        let nper = match self.eval_number_arg(&args[2]) {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        let pv = match self.eval_number_arg(&args[3]) {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        let fv = match self.eval_optional_number_arg(args.get(4)) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        let typ = match self.eval_optional_number_arg(args.get(5)) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        Self::excel_result_number(financial::ppmt(rate, per, nper, pv, fv, typ))
    }

    fn fn_sln(&self, args: &[CompiledExpr]) -> Value {
        if args.len() != 3 {
            return Value::Error(ErrorKind::Value);
        }
        let cost = match self.eval_number_arg(&args[0]) {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        let salvage = match self.eval_number_arg(&args[1]) {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        let life = match self.eval_number_arg(&args[2]) {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        Self::excel_result_number(financial::sln(cost, salvage, life))
    }

    fn fn_syd(&self, args: &[CompiledExpr]) -> Value {
        if args.len() != 4 {
            return Value::Error(ErrorKind::Value);
        }
        let cost = match self.eval_number_arg(&args[0]) {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        let salvage = match self.eval_number_arg(&args[1]) {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        let life = match self.eval_number_arg(&args[2]) {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        let per = match self.eval_number_arg(&args[3]) {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        Self::excel_result_number(financial::syd(cost, salvage, life, per))
    }

    fn fn_ddb(&self, args: &[CompiledExpr]) -> Value {
        if args.len() < 4 || args.len() > 5 {
            return Value::Error(ErrorKind::Value);
        }
        let cost = match self.eval_number_arg(&args[0]) {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        let salvage = match self.eval_number_arg(&args[1]) {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        let life = match self.eval_number_arg(&args[2]) {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        let period = match self.eval_number_arg(&args[3]) {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        let factor = match self.eval_optional_number_arg(args.get(4)) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        Self::excel_result_number(financial::ddb(cost, salvage, life, period, factor))
    }

    fn fn_npv(&self, args: &[CompiledExpr]) -> Value {
        if args.len() < 2 {
            return Value::Error(ErrorKind::Value);
        }
        let rate = match self.eval_number_arg(&args[0]) {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };

        let mut values = Vec::new();
        for arg in &args[1..] {
            match self.collect_npv_values_from_arg(arg) {
                Ok(mut nums) => values.append(&mut nums),
                Err(e) => return Value::Error(e),
            }
        }

        Self::excel_result_number(financial::npv(rate, &values))
    }

    fn fn_irr(&self, args: &[CompiledExpr]) -> Value {
        if args.is_empty() || args.len() > 2 {
            return Value::Error(ErrorKind::Value);
        }

        let values = match self.collect_irr_values_from_arg(&args[0]) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };

        let guess = match self.eval_optional_number_arg(args.get(1)) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };

        Self::excel_result_number(financial::irr(&values, guess))
    }

    fn fn_xnpv(&self, args: &[CompiledExpr]) -> Value {
        if args.len() != 3 {
            return Value::Error(ErrorKind::Value);
        }
        let rate = match self.eval_number_arg(&args[0]) {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };

        let values = match self.collect_numbers_strict_from_arg(&args[1]) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        let dates = match self.collect_numbers_strict_from_arg(&args[2]) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };

        Self::excel_result_number(financial::xnpv(rate, &values, &dates))
    }

    fn fn_xirr(&self, args: &[CompiledExpr]) -> Value {
        if args.len() < 2 || args.len() > 3 {
            return Value::Error(ErrorKind::Value);
        }
        let values = match self.collect_numbers_strict_from_arg(&args[0]) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        let dates = match self.collect_numbers_strict_from_arg(&args[1]) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        let guess = match self.eval_optional_number_arg(args.get(2)) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };

        Self::excel_result_number(financial::xirr(&values, &dates, guess))
    }
}

fn parse_number_from_text(s: &str) -> Option<f64> {
    let trimmed = s.trim();
    if trimmed.is_empty() {
        return None;
    }
    trimmed.parse::<f64>().ok()
}

fn coerce_to_number(v: &Value) -> Result<f64, ErrorKind> {
    match v {
        Value::Number(n) => Ok(*n),
        Value::Bool(b) => Ok(if *b { 1.0 } else { 0.0 }),
        Value::Blank => Ok(0.0),
        Value::Text(s) => parse_number_from_text(s).ok_or(ErrorKind::Value),
        Value::Error(e) => Err(*e),
    }
}

fn coerce_to_bool(v: &Value) -> Result<bool, ErrorKind> {
    match v {
        Value::Bool(b) => Ok(*b),
        Value::Number(n) => Ok(*n != 0.0),
        Value::Blank => Ok(false),
        Value::Text(s) => {
            let t = s.trim();
            if t.eq_ignore_ascii_case("TRUE") {
                return Ok(true);
            }
            if t.eq_ignore_ascii_case("FALSE") {
                return Ok(false);
            }
            if let Some(n) = parse_number_from_text(t) {
                return Ok(n != 0.0);
            }
            Err(ErrorKind::Value)
        }
        Value::Error(e) => Err(*e),
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
    })
}
