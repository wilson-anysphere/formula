use std::collections::HashMap;
use std::time::Instant;

use thiserror::Error;

use crate::ast::{BinOp, Expr, ProcedureDef, ProcedureKind, Stmt, UnOp, VbaProgram};
use crate::object_model::{
    range_on_active_sheet, Spreadsheet, VbaObject, VbaObjectRef, VbaRangeRef,
};
use crate::sandbox::{Permission, PermissionChecker, VbaSandboxPolicy};
use crate::value::VbaValue;

#[derive(Debug, Error)]
pub enum VbaError {
    #[error("Parse error: {0}")]
    Parse(String),
    #[error("Runtime error: {0}")]
    Runtime(String),
    #[error("Sandbox violation: {0}")]
    Sandbox(String),
    #[error("Execution timed out")]
    Timeout,
    #[error("Execution exceeded step limit")]
    StepLimit,
}

#[derive(Debug, Default)]
pub struct ExecutionResult {
    pub returned: Option<VbaValue>,
}

pub struct VbaRuntime {
    program: VbaProgram,
    sandbox: VbaSandboxPolicy,
    permission_checker: Option<Box<dyn PermissionChecker>>,
}

impl VbaRuntime {
    pub fn new(program: VbaProgram) -> Self {
        Self {
            program,
            sandbox: VbaSandboxPolicy::default(),
            permission_checker: None,
        }
    }

    pub fn with_sandbox_policy(mut self, policy: VbaSandboxPolicy) -> Self {
        self.sandbox = policy;
        self
    }

    pub fn with_permission_checker(mut self, checker: Box<dyn PermissionChecker>) -> Self {
        self.permission_checker = Some(checker);
        self
    }

    pub fn execute(
        &self,
        spreadsheet: &mut dyn Spreadsheet,
        entry: &str,
        args: &[VbaValue],
    ) -> Result<ExecutionResult, VbaError> {
        let proc = self
            .program
            .get(entry)
            .ok_or_else(|| VbaError::Runtime(format!("Unknown procedure `{entry}`")))?;
        let mut exec = Executor::new(
            &self.program,
            spreadsheet,
            &self.sandbox,
            self.permission_checker.as_deref(),
        );
        exec.call_procedure(proc, args)
    }

    /// Convenience: execute `Workbook_Open` if present.
    pub fn fire_workbook_open(&self, spreadsheet: &mut dyn Spreadsheet) -> Result<(), VbaError> {
        if self.program.get("workbook_open").is_some() {
            self.execute(spreadsheet, "Workbook_Open", &[])?;
        }
        Ok(())
    }

    /// Fire `Worksheet_Change` if present.
    pub fn fire_worksheet_change(
        &self,
        spreadsheet: &mut dyn Spreadsheet,
        target: VbaRangeRef,
    ) -> Result<(), VbaError> {
        if self.program.get("worksheet_change").is_some() {
            self.execute(
                spreadsheet,
                "Worksheet_Change",
                &[VbaValue::Object(VbaObjectRef::new(VbaObject::Range(
                    target,
                )))],
            )?;
        }
        Ok(())
    }

    /// Fire `Worksheet_SelectionChange` if present.
    pub fn fire_worksheet_selection_change(
        &self,
        spreadsheet: &mut dyn Spreadsheet,
        target: VbaRangeRef,
    ) -> Result<(), VbaError> {
        if self.program.get("worksheet_selectionchange").is_some() {
            self.execute(
                spreadsheet,
                "Worksheet_SelectionChange",
                &[VbaValue::Object(VbaObjectRef::new(VbaObject::Range(
                    target,
                )))],
            )?;
        }
        Ok(())
    }
}

#[derive(Debug, Clone)]
enum ErrorMode {
    Default,
    ResumeNext,
    GotoLabel(String),
}

struct Frame {
    locals: HashMap<String, VbaValue>,
    error_mode: ErrorMode,
}

struct Executor<'a> {
    program: &'a VbaProgram,
    sheet: &'a mut dyn Spreadsheet,
    sandbox: &'a VbaSandboxPolicy,
    permission_checker: Option<&'a dyn PermissionChecker>,
    start: Instant,
    steps: u64,
}

impl<'a> Executor<'a> {
    fn new(
        program: &'a VbaProgram,
        sheet: &'a mut dyn Spreadsheet,
        sandbox: &'a VbaSandboxPolicy,
        permission_checker: Option<&'a dyn PermissionChecker>,
    ) -> Self {
        Self {
            program,
            sheet,
            sandbox,
            permission_checker,
            start: Instant::now(),
            steps: 0,
        }
    }

    fn tick(&mut self) -> Result<(), VbaError> {
        self.steps = self.steps.saturating_add(1);
        if self.steps > self.sandbox.max_steps {
            return Err(VbaError::StepLimit);
        }
        // Only check wall clock every 256 steps.
        if (self.steps & 0xFF) == 0 && self.start.elapsed() > self.sandbox.max_execution_time {
            return Err(VbaError::Timeout);
        }
        Ok(())
    }

    fn call_procedure(
        &mut self,
        proc: &'a ProcedureDef,
        args: &[VbaValue],
    ) -> Result<ExecutionResult, VbaError> {
        let mut frame = Frame {
            locals: HashMap::new(),
            error_mode: ErrorMode::Default,
        };

        // VBA Functions return by assigning to the function name.
        if proc.kind == ProcedureKind::Function {
            frame
                .locals
                .insert(proc.name.to_ascii_lowercase(), VbaValue::Empty);
        }

        for (idx, param) in proc.params.iter().enumerate() {
            let value = args.get(idx).cloned().unwrap_or(VbaValue::Empty);
            frame.locals.insert(param.name.to_ascii_lowercase(), value);
        }

        // Provide global-ish built-ins via locals for simplicity.
        frame.locals.insert(
            "application".to_string(),
            VbaValue::Object(VbaObjectRef::new(VbaObject::Application)),
        );
        frame.locals.insert(
            "thisworkbook".to_string(),
            VbaValue::Object(VbaObjectRef::new(VbaObject::Workbook)),
        );

        let flow = self.exec_block(&mut frame, &proc.body)?;
        let mut result = ExecutionResult::default();
        match flow {
            ControlFlow::Continue | ControlFlow::ExitSub | ControlFlow::ExitFor => {}
            ControlFlow::ExitFunction => {}
            ControlFlow::Goto(label) => {
                return Err(VbaError::Runtime(format!(
                    "GoTo `{label}` reached outside of its block"
                )));
            }
        }

        if proc.kind == ProcedureKind::Function {
            result.returned = Some(
                frame
                    .locals
                    .get(&proc.name.to_ascii_lowercase())
                    .cloned()
                    .unwrap_or(VbaValue::Empty),
            );
        }
        Ok(result)
    }

    fn exec_block(&mut self, frame: &mut Frame, body: &[Stmt]) -> Result<ControlFlow, VbaError> {
        // Precompute label map for GoTo/On Error GoTo.
        let label_map = collect_labels(body);
        let mut pc: usize = 0;
        while pc < body.len() {
            self.tick()?;
            let stmt = &body[pc];
            match self.exec_stmt(frame, stmt, &label_map)? {
                ControlFlow::Continue => pc += 1,
                ControlFlow::ExitSub => return Ok(ControlFlow::ExitSub),
                ControlFlow::ExitFunction => return Ok(ControlFlow::ExitFunction),
                ControlFlow::ExitFor => return Ok(ControlFlow::ExitFor),
                ControlFlow::Goto(label) => {
                    if let Some(dest) = label_map.get(&label) {
                        pc = *dest;
                    } else {
                        // The label may live outside of this nested block; propagate upwards so the
                        // outer block can resolve it.
                        return Ok(ControlFlow::Goto(label));
                    }
                }
            }
        }
        Ok(ControlFlow::Continue)
    }

    fn exec_stmt(
        &mut self,
        frame: &mut Frame,
        stmt: &Stmt,
        labels: &HashMap<String, usize>,
    ) -> Result<ControlFlow, VbaError> {
        let res = (|| -> Result<ControlFlow, VbaError> {
            match stmt {
                Stmt::Dim(vars) => {
                    for v in vars {
                        frame.locals.insert(v.to_ascii_lowercase(), VbaValue::Empty);
                    }
                    Ok(ControlFlow::Continue)
                }
                Stmt::Assign { target, value } => {
                    let rhs = self.eval_expr(frame, value)?;
                    self.assign(frame, target, rhs)?;
                    Ok(ControlFlow::Continue)
                }
                Stmt::Set { target, value } => {
                    let rhs = self.eval_expr(frame, value)?;
                    self.assign(frame, target, rhs)?;
                    Ok(ControlFlow::Continue)
                }
                Stmt::ExprStmt(expr) => {
                    // VBA allows calling Subs/methods without parentheses. We model that by treating a
                    // bare member access as a zero-arg method invocation when possible, falling back
                    // to a property read otherwise.
                    match expr {
                        Expr::Member { object, member } => {
                            let obj = self.eval_expr(frame, object)?;
                            let obj = obj.as_object().ok_or_else(|| {
                                VbaError::Runtime("Call on non-object".to_string())
                            })?;
                            match self.call_object_method(frame, obj.clone(), member, &[]) {
                                Ok(_) => Ok(ControlFlow::Continue),
                                Err(VbaError::Runtime(_)) => {
                                    let _ = self.eval_expr(frame, expr)?;
                                    Ok(ControlFlow::Continue)
                                }
                                Err(e) => Err(e),
                            }
                        }
                        Expr::Var(name) => {
                            let name_lc = name.to_ascii_lowercase();
                            let is_callable = matches!(
                                name_lc.as_str(),
                                "range"
                                    | "cells"
                                    | "msgbox"
                                    | "debugprint"
                                    | "array"
                                    | "worksheets"
                                    | "createobject"
                            ) || self.program.get(&name_lc).is_some();

                            if is_callable {
                                let _ = self.call_global(frame, name, &[])?;
                            } else {
                                let _ = self.eval_expr(frame, expr)?;
                            }
                            Ok(ControlFlow::Continue)
                        }
                        _ => {
                            let _ = self.eval_expr(frame, expr)?;
                            Ok(ControlFlow::Continue)
                        }
                    }
                }
                Stmt::If {
                    cond,
                    then_body,
                    elseifs,
                    else_body,
                } => {
                    if self.eval_expr(frame, cond)?.is_truthy() {
                        return self.exec_block(frame, then_body);
                    }
                    for (c, b) in elseifs {
                        if self.eval_expr(frame, c)?.is_truthy() {
                            return self.exec_block(frame, b);
                        }
                    }
                    self.exec_block(frame, else_body)
                }
                Stmt::For {
                    var,
                    start,
                    end,
                    step,
                    body,
                } => {
                    let start_v = self.eval_expr(frame, start)?.to_f64().unwrap_or(0.0);
                    let end_v = self.eval_expr(frame, end)?.to_f64().unwrap_or(0.0);
                    let step_v = if let Some(step) = step {
                        self.eval_expr(frame, step)?.to_f64().unwrap_or(1.0)
                    } else {
                        1.0
                    };
                    if step_v == 0.0 {
                        return Err(VbaError::Runtime("For loop Step cannot be 0".to_string()));
                    }

                    let mut i = start_v;
                    let cmp = |i: f64, end: f64, step: f64| -> bool {
                        if step > 0.0 {
                            i <= end
                        } else {
                            i >= end
                        }
                    };

                    while cmp(i, end_v, step_v) {
                        frame
                            .locals
                            .insert(var.to_ascii_lowercase(), VbaValue::Double(i));
                        match self.exec_block(frame, body)? {
                            ControlFlow::Continue => {}
                            ControlFlow::ExitFor => break,
                            ControlFlow::ExitSub => return Ok(ControlFlow::ExitSub),
                            ControlFlow::ExitFunction => return Ok(ControlFlow::ExitFunction),
                            ControlFlow::Goto(label) => return Ok(ControlFlow::Goto(label)),
                        }
                        i += step_v;
                    }
                    Ok(ControlFlow::Continue)
                }
                Stmt::DoWhile { cond, body } => {
                    while self.eval_expr(frame, cond)?.is_truthy() {
                        match self.exec_block(frame, body)? {
                            ControlFlow::Continue => {}
                            ControlFlow::ExitFor => {
                                // Treat Exit For inside Do While as error; ignore for now.
                                break;
                            }
                            ControlFlow::ExitSub => return Ok(ControlFlow::ExitSub),
                            ControlFlow::ExitFunction => return Ok(ControlFlow::ExitFunction),
                            ControlFlow::Goto(label) => return Ok(ControlFlow::Goto(label)),
                        }
                    }
                    Ok(ControlFlow::Continue)
                }
                Stmt::ExitSub => Ok(ControlFlow::ExitSub),
                Stmt::ExitFunction => Ok(ControlFlow::ExitFunction),
                Stmt::ExitFor => Ok(ControlFlow::ExitFor),
                Stmt::OnErrorResumeNext => {
                    frame.error_mode = ErrorMode::ResumeNext;
                    Ok(ControlFlow::Continue)
                }
                Stmt::OnErrorGoto0 => {
                    frame.error_mode = ErrorMode::Default;
                    Ok(ControlFlow::Continue)
                }
                Stmt::OnErrorGotoLabel(label) => {
                    frame.error_mode = ErrorMode::GotoLabel(label.to_ascii_lowercase());
                    Ok(ControlFlow::Continue)
                }
                Stmt::Label(_) => Ok(ControlFlow::Continue),
                Stmt::Goto(label) => Ok(ControlFlow::Goto(label.to_ascii_lowercase())),
            }
        })();

        res.or_else(|err| self.handle_stmt_error(frame, err, labels))
    }

    fn handle_stmt_error(
        &mut self,
        frame: &mut Frame,
        err: VbaError,
        _labels: &HashMap<String, usize>,
    ) -> Result<ControlFlow, VbaError> {
        match &frame.error_mode {
            ErrorMode::Default => Err(err),
            ErrorMode::ResumeNext => Ok(ControlFlow::Continue),
            ErrorMode::GotoLabel(label) => Ok(ControlFlow::Goto(label.clone())),
        }
    }

    fn assign(
        &mut self,
        frame: &mut Frame,
        target: &Expr,
        value: VbaValue,
    ) -> Result<(), VbaError> {
        match target {
            Expr::Var(name) => {
                frame.locals.insert(name.to_ascii_lowercase(), value);
                Ok(())
            }
            Expr::Member { object, member } => {
                let obj = self.eval_expr(frame, object)?;
                let obj = obj.as_object().ok_or_else(|| {
                    VbaError::Runtime("Assignment target is not an object".to_string())
                })?;
                self.set_object_member(obj, member, value)
            }
            _ => Err(VbaError::Runtime(
                "Unsupported assignment target".to_string(),
            )),
        }
    }

    fn set_object_member(
        &mut self,
        obj: VbaObjectRef,
        member: &str,
        value: VbaValue,
    ) -> Result<(), VbaError> {
        let member_lc = member.to_ascii_lowercase();
        match &mut *obj.borrow_mut() {
            VbaObject::Range(range) => match member_lc.as_str() {
                "value" => {
                    self.sheet.set_cell_value(
                        range.sheet,
                        range.start_row,
                        range.start_col,
                        value,
                    )?;
                    Ok(())
                }
                "formula" => {
                    let formula = value.to_string_lossy();
                    self.sheet.set_cell_formula(
                        range.sheet,
                        range.start_row,
                        range.start_col,
                        formula,
                    )?;
                    Ok(())
                }
                _ => Err(VbaError::Runtime(format!(
                    "Unknown Range member `{member}`"
                ))),
            },
            VbaObject::Worksheet { .. } => Err(VbaError::Runtime(format!(
                "Cannot assign to Worksheet member `{member}`"
            ))),
            VbaObject::Application => Err(VbaError::Runtime(format!(
                "Cannot assign to Application member `{member}`"
            ))),
            VbaObject::Workbook => Err(VbaError::Runtime(format!(
                "Cannot assign to Workbook member `{member}`"
            ))),
            VbaObject::Collection { .. } => Err(VbaError::Runtime(format!(
                "Cannot assign to Collection member `{member}`"
            ))),
        }
    }

    fn eval_expr(&mut self, frame: &mut Frame, expr: &Expr) -> Result<VbaValue, VbaError> {
        self.tick()?;
        match expr {
            Expr::Literal(v) => Ok(v.clone()),
            Expr::Var(name) => {
                let name_lc = name.to_ascii_lowercase();
                if let Some(v) = frame.locals.get(&name_lc).cloned() {
                    return Ok(v);
                }
                // Built-in global functions/properties without explicit object.
                match name_lc.as_str() {
                    "activesheet" => {
                        Ok(VbaValue::Object(VbaObjectRef::new(VbaObject::Worksheet {
                            sheet: self.sheet.active_sheet(),
                        })))
                    }
                    "activecell" => {
                        let (r, c) = self.sheet.active_cell();
                        Ok(VbaValue::Object(VbaObjectRef::new(VbaObject::Range(
                            VbaRangeRef {
                                sheet: self.sheet.active_sheet(),
                                start_row: r,
                                start_col: c,
                                end_row: r,
                                end_col: c,
                            },
                        ))))
                    }
                    _ => Ok(VbaValue::Empty),
                }
            }
            Expr::Unary { op, expr } => {
                let v = self.eval_expr(frame, expr)?;
                match op {
                    UnOp::Neg => Ok(VbaValue::Double(-v.to_f64().unwrap_or(0.0))),
                    UnOp::Not => Ok(VbaValue::Boolean(!v.is_truthy())),
                }
            }
            Expr::Binary { op, left, right } => {
                let l = self.eval_expr(frame, left)?;
                let r = self.eval_expr(frame, right)?;
                self.eval_binop(*op, l, r)
            }
            Expr::Member { object, member } => {
                let obj = self.eval_expr(frame, object)?;
                let obj = obj
                    .as_object()
                    .ok_or_else(|| VbaError::Runtime("Member access on non-object".to_string()))?;
                self.get_object_member(obj, member)
            }
            Expr::Call { callee, args } => {
                // First evaluate callee expression. This could be:
                // - Var("Range") => built-in Range function.
                // - Member(obj, "Range") => worksheet method.
                // - Member(obj, "Cells") => worksheet method.
                // - Var(procName) => procedure call.
                //
                // We evaluate as special cases before generic eval of callee.
                if let Expr::Var(name) = &**callee {
                    // `arr(i)` is an index operation when `arr` is an Array value.
                    let name_lc = name.to_ascii_lowercase();
                    if let Some(VbaValue::Array(arr)) = frame.locals.get(&name_lc).cloned() {
                        let idx = self
                            .eval_expr(
                                frame,
                                args.get(0).ok_or_else(|| {
                                    VbaError::Runtime("Array index missing".to_string())
                                })?,
                            )?
                            .to_f64()
                            .unwrap_or(0.0) as isize;
                        if idx < 0 || (idx as usize) >= arr.len() {
                            return Ok(VbaValue::Empty);
                        }
                        return Ok(arr[idx as usize].clone());
                    }
                    return self.call_global(frame, name, args);
                }
                if let Expr::Member { object, member } = &**callee {
                    let obj = self.eval_expr(frame, object)?;
                    let obj = obj
                        .as_object()
                        .ok_or_else(|| VbaError::Runtime("Call on non-object".to_string()))?;
                    return self.call_object_method(frame, obj, member, args);
                }
                // Otherwise evaluate callee and attempt to call/index.
                let callee_val = self.eval_expr(frame, callee)?;
                if let VbaValue::Array(arr) = callee_val {
                    let idx = self
                        .eval_expr(
                            frame,
                            args.get(0).ok_or_else(|| {
                                VbaError::Runtime("Array index missing".to_string())
                            })?,
                        )?
                        .to_f64()
                        .unwrap_or(0.0) as isize;
                    if idx < 0 || (idx as usize) >= arr.len() {
                        return Ok(VbaValue::Empty);
                    }
                    return Ok(arr[idx as usize].clone());
                }
                Err(VbaError::Runtime("Unsupported call target".to_string()))
            }
            Expr::Index { .. } => Err(VbaError::Runtime("Indexing not implemented".to_string())),
        }
    }

    fn eval_binop(&self, op: BinOp, l: VbaValue, r: VbaValue) -> Result<VbaValue, VbaError> {
        match op {
            BinOp::Add => Ok(VbaValue::Double(
                l.to_f64().unwrap_or(0.0) + r.to_f64().unwrap_or(0.0),
            )),
            BinOp::Sub => Ok(VbaValue::Double(
                l.to_f64().unwrap_or(0.0) - r.to_f64().unwrap_or(0.0),
            )),
            BinOp::Mul => Ok(VbaValue::Double(
                l.to_f64().unwrap_or(0.0) * r.to_f64().unwrap_or(0.0),
            )),
            BinOp::Div => Ok(VbaValue::Double(
                l.to_f64().unwrap_or(0.0) / r.to_f64().unwrap_or(0.0),
            )),
            BinOp::Concat => Ok(VbaValue::String(format!(
                "{}{}",
                l.to_string_lossy(),
                r.to_string_lossy()
            ))),
            BinOp::Eq => Ok(VbaValue::Boolean(l == r)),
            BinOp::Ne => Ok(VbaValue::Boolean(l != r)),
            BinOp::Lt => Ok(VbaValue::Boolean(
                l.to_f64().unwrap_or(0.0) < r.to_f64().unwrap_or(0.0),
            )),
            BinOp::Le => Ok(VbaValue::Boolean(
                l.to_f64().unwrap_or(0.0) <= r.to_f64().unwrap_or(0.0),
            )),
            BinOp::Gt => Ok(VbaValue::Boolean(
                l.to_f64().unwrap_or(0.0) > r.to_f64().unwrap_or(0.0),
            )),
            BinOp::Ge => Ok(VbaValue::Boolean(
                l.to_f64().unwrap_or(0.0) >= r.to_f64().unwrap_or(0.0),
            )),
            BinOp::And => Ok(VbaValue::Boolean(l.is_truthy() && r.is_truthy())),
            BinOp::Or => Ok(VbaValue::Boolean(l.is_truthy() || r.is_truthy())),
        }
    }

    fn call_global(
        &mut self,
        frame: &mut Frame,
        name: &str,
        args: &[Expr],
    ) -> Result<VbaValue, VbaError> {
        let name_lc = name.to_ascii_lowercase();
        match name_lc.as_str() {
            "range" => {
                let arg0 = args
                    .get(0)
                    .ok_or_else(|| VbaError::Runtime("Range() missing argument".to_string()))?;
                let a1 = self.eval_expr(frame, arg0)?.to_string_lossy();
                Ok(VbaValue::Object(range_on_active_sheet(self.sheet, &a1)?))
            }
            "cells" => {
                let row = self
                    .eval_expr(
                        frame,
                        args.get(0)
                            .ok_or_else(|| VbaError::Runtime("Cells() missing row".to_string()))?,
                    )?
                    .to_f64()
                    .unwrap_or(1.0) as u32;
                let col = self
                    .eval_expr(
                        frame,
                        args.get(1)
                            .ok_or_else(|| VbaError::Runtime("Cells() missing col".to_string()))?,
                    )?
                    .to_f64()
                    .unwrap_or(1.0) as u32;
                Ok(VbaValue::Object(VbaObjectRef::new(VbaObject::Range(
                    VbaRangeRef {
                        sheet: self.sheet.active_sheet(),
                        start_row: row,
                        start_col: col,
                        end_row: row,
                        end_col: col,
                    },
                ))))
            }
            "msgbox" => {
                let msg = self
                    .eval_expr(
                        frame,
                        args.get(0).ok_or_else(|| {
                            VbaError::Runtime("MsgBox() missing prompt".to_string())
                        })?,
                    )?
                    .to_string_lossy();
                self.sheet.log(format!("MsgBox: {msg}"));
                Ok(VbaValue::Double(0.0))
            }
            "debugprint" => {
                let mut parts = Vec::new();
                for arg in args {
                    parts.push(self.eval_expr(frame, arg)?.to_string_lossy());
                }
                self.sheet.log(format!("Debug.Print {}", parts.join(" ")));
                Ok(VbaValue::Empty)
            }
            "array" => {
                let mut values = Vec::new();
                for arg in args {
                    values.push(self.eval_expr(frame, arg)?);
                }
                Ok(VbaValue::Array(std::rc::Rc::new(values)))
            }
            "worksheets" => {
                let name = self
                    .eval_expr(
                        frame,
                        args.get(0).ok_or_else(|| {
                            VbaError::Runtime("Worksheets() missing name".to_string())
                        })?,
                    )?
                    .to_string_lossy();
                let idx = self
                    .sheet
                    .sheet_index(&name)
                    .ok_or_else(|| VbaError::Runtime(format!("Unknown worksheet `{name}`")))?;
                Ok(VbaValue::Object(VbaObjectRef::new(VbaObject::Worksheet {
                    sheet: idx,
                })))
            }
            "__new" => {
                let class = self
                    .eval_expr(
                        frame,
                        args.get(0).ok_or_else(|| {
                            VbaError::Runtime("__new() missing class".to_string())
                        })?,
                    )?
                    .to_string_lossy();

                match class.to_ascii_lowercase().as_str() {
                    "collection" => {
                        Ok(VbaValue::Object(VbaObjectRef::new(VbaObject::Collection {
                            items: Vec::new(),
                        })))
                    }
                    other => Err(VbaError::Runtime(format!(
                        "Unsupported class in New: {other}"
                    ))),
                }
            }
            "createobject" => {
                self.require_permission(Permission::Network, "CreateObject")?;
                Err(VbaError::Sandbox(
                    "CreateObject is not supported".to_string(),
                ))
            }
            _ => {
                // Procedure call.
                if let Some(proc) = self.program.get(&name_lc) {
                    let mut arg_vals = Vec::new();
                    for arg in args {
                        arg_vals.push(self.eval_expr(frame, arg)?);
                    }
                    let res = self.call_procedure(proc, &arg_vals)?;
                    Ok(res.returned.unwrap_or(VbaValue::Empty))
                } else {
                    Err(VbaError::Runtime(format!(
                        "Unknown function/procedure `{name}`"
                    )))
                }
            }
        }
    }

    fn call_object_method(
        &mut self,
        frame: &mut Frame,
        obj: VbaObjectRef,
        member: &str,
        args: &[Expr],
    ) -> Result<VbaValue, VbaError> {
        let member_lc = member.to_ascii_lowercase();
        let snapshot = obj.borrow().clone();
        match snapshot {
            VbaObject::Worksheet { sheet } => match member_lc.as_str() {
                "range" => {
                    let a1 = self
                        .eval_expr(
                            frame,
                            args.get(0).ok_or_else(|| {
                                VbaError::Runtime("Range() missing argument".to_string())
                            })?,
                        )?
                        .to_string_lossy();
                    let range_ref = self.sheet_range(sheet, &a1)?;
                    Ok(VbaValue::Object(VbaObjectRef::new(VbaObject::Range(
                        range_ref,
                    ))))
                }
                "cells" => {
                    let row = self
                        .eval_expr(
                            frame,
                            args.get(0).ok_or_else(|| {
                                VbaError::Runtime("Cells() missing row".to_string())
                            })?,
                        )?
                        .to_f64()
                        .unwrap_or(1.0) as u32;
                    let col = self
                        .eval_expr(
                            frame,
                            args.get(1).ok_or_else(|| {
                                VbaError::Runtime("Cells() missing col".to_string())
                            })?,
                        )?
                        .to_f64()
                        .unwrap_or(1.0) as u32;
                    Ok(VbaValue::Object(VbaObjectRef::new(VbaObject::Range(
                        VbaRangeRef {
                            sheet,
                            start_row: row,
                            start_col: col,
                            end_row: row,
                            end_col: col,
                        },
                    ))))
                }
                "activate" => {
                    self.sheet.set_active_sheet(sheet)?;
                    Ok(VbaValue::Empty)
                }
                _ => Err(VbaError::Runtime(format!(
                    "Unknown Worksheet method `{member}`"
                ))),
            },
            VbaObject::Range(range) => match member_lc.as_str() {
                "select" => {
                    self.sheet.set_active_sheet(range.sheet)?;
                    self.sheet
                        .set_active_cell(range.start_row, range.start_col)?;
                    Ok(VbaValue::Empty)
                }
                "copy" => {
                    if args.is_empty() {
                        // Clipboard not modelled; treat as no-op.
                        return Ok(VbaValue::Empty);
                    }
                    let dest = self.eval_expr(frame, args.get(0).unwrap())?;
                    let dest = dest.as_object().ok_or_else(|| {
                        VbaError::Runtime("Copy destination must be a Range".to_string())
                    })?;
                    let dest_range = match &*dest.borrow() {
                        VbaObject::Range(r) => *r,
                        _ => {
                            return Err(VbaError::Runtime(
                                "Copy destination must be a Range".to_string(),
                            ))
                        }
                    };
                    self.copy_range(range, dest_range)?;
                    Ok(VbaValue::Empty)
                }
                "autofill" => {
                    // `Range.AutoFill Destination, [Type]`
                    let dest_expr = args.get(0).ok_or_else(|| {
                        VbaError::Runtime("AutoFill() missing destination".to_string())
                    })?;
                    let dest = self.eval_expr(frame, dest_expr)?;
                    let dest = dest.as_object().ok_or_else(|| {
                        VbaError::Runtime("AutoFill destination must be a Range".to_string())
                    })?;
                    let dest_range = match &*dest.borrow() {
                        VbaObject::Range(r) => *r,
                        _ => {
                            return Err(VbaError::Runtime(
                                "AutoFill destination must be a Range".to_string(),
                            ))
                        }
                    };
                    self.copy_range(range, dest_range)?;
                    Ok(VbaValue::Empty)
                }
                _ => Err(VbaError::Runtime(format!(
                    "Unknown Range method `{member}`"
                ))),
            },
            VbaObject::Application => match member_lc.as_str() {
                "range" => {
                    let a1 = self
                        .eval_expr(
                            frame,
                            args.get(0).ok_or_else(|| {
                                VbaError::Runtime("Range() missing argument".to_string())
                            })?,
                        )?
                        .to_string_lossy();
                    Ok(VbaValue::Object(range_on_active_sheet(self.sheet, &a1)?))
                }
                _ => Err(VbaError::Runtime(format!(
                    "Unknown Application method `{member}`"
                ))),
            },
            VbaObject::Workbook => match member_lc.as_str() {
                "worksheets" => {
                    let name = self
                        .eval_expr(
                            frame,
                            args.get(0).ok_or_else(|| {
                                VbaError::Runtime("Worksheets() missing name".to_string())
                            })?,
                        )?
                        .to_string_lossy();
                    let idx = self
                        .sheet
                        .sheet_index(&name)
                        .ok_or_else(|| VbaError::Runtime(format!("Unknown worksheet `{name}`")))?;
                    Ok(VbaValue::Object(VbaObjectRef::new(VbaObject::Worksheet {
                        sheet: idx,
                    })))
                }
                _ => Err(VbaError::Runtime(format!(
                    "Unknown Workbook method `{member}`"
                ))),
            },
            VbaObject::Collection { .. } => match member_lc.as_str() {
                "add" => {
                    let item = self.eval_expr(
                        frame,
                        args.get(0)
                            .ok_or_else(|| VbaError::Runtime("Add() missing item".to_string()))?,
                    )?;
                    if let VbaObject::Collection { items } = &mut *obj.borrow_mut() {
                        items.push(item);
                    }
                    Ok(VbaValue::Empty)
                }
                "item" => {
                    let index = self
                        .eval_expr(
                            frame,
                            args.get(0).ok_or_else(|| {
                                VbaError::Runtime("Item() missing index".to_string())
                            })?,
                        )?
                        .to_f64()
                        .unwrap_or(0.0) as isize;

                    if index <= 0 {
                        return Ok(VbaValue::Empty);
                    }
                    if let VbaObject::Collection { items } = &*obj.borrow() {
                        let idx = (index - 1) as usize;
                        Ok(items.get(idx).cloned().unwrap_or(VbaValue::Empty))
                    } else {
                        Ok(VbaValue::Empty)
                    }
                }
                _ => Err(VbaError::Runtime(format!(
                    "Unknown Collection method `{member}`"
                ))),
            },
        }
    }

    fn sheet_range(&self, sheet: usize, a1: &str) -> Result<VbaRangeRef, VbaError> {
        let (r1, c1, r2, c2) = crate::object_model::a1_to_row_col_range(a1)?;
        Ok(VbaRangeRef {
            sheet,
            start_row: r1,
            start_col: c1,
            end_row: r2,
            end_col: c2,
        })
    }

    fn get_object_member(&mut self, obj: VbaObjectRef, member: &str) -> Result<VbaValue, VbaError> {
        let member_lc = member.to_ascii_lowercase();
        match &*obj.borrow() {
            VbaObject::Application => match member_lc.as_str() {
                "activesheet" => Ok(VbaValue::Object(VbaObjectRef::new(VbaObject::Worksheet {
                    sheet: self.sheet.active_sheet(),
                }))),
                "activecell" => {
                    let (r, c) = self.sheet.active_cell();
                    Ok(VbaValue::Object(VbaObjectRef::new(VbaObject::Range(
                        VbaRangeRef {
                            sheet: self.sheet.active_sheet(),
                            start_row: r,
                            start_col: c,
                            end_row: r,
                            end_col: c,
                        },
                    ))))
                }
                _ => Err(VbaError::Runtime(format!(
                    "Unknown Application member `{member}`"
                ))),
            },
            VbaObject::Workbook => match member_lc.as_str() {
                "activesheet" => Ok(VbaValue::Object(VbaObjectRef::new(VbaObject::Worksheet {
                    sheet: self.sheet.active_sheet(),
                }))),
                _ => Err(VbaError::Runtime(format!(
                    "Unknown Workbook member `{member}`"
                ))),
            },
            VbaObject::Worksheet { sheet } => match member_lc.as_str() {
                "name" => Ok(VbaValue::String(
                    self.sheet.sheet_name(*sheet).unwrap_or("").to_string(),
                )),
                _ => Err(VbaError::Runtime(format!(
                    "Unknown Worksheet member `{member}`"
                ))),
            },
            VbaObject::Range(range) => match member_lc.as_str() {
                "value" => self
                    .sheet
                    .get_cell_value(range.sheet, range.start_row, range.start_col),
                "formula" => Ok(self
                    .sheet
                    .get_cell_formula(range.sheet, range.start_row, range.start_col)?
                    .map(VbaValue::String)
                    .unwrap_or(VbaValue::Empty)),
                _ => Err(VbaError::Runtime(format!(
                    "Unknown Range member `{member}`"
                ))),
            },
            VbaObject::Collection { items } => match member_lc.as_str() {
                "count" => Ok(VbaValue::Double(items.len() as f64)),
                _ => Err(VbaError::Runtime(format!(
                    "Unknown Collection member `{member}`"
                ))),
            },
        }
    }

    fn require_permission(&self, permission: Permission, feature: &str) -> Result<(), VbaError> {
        if self.sandbox.can(permission, self.permission_checker) {
            Ok(())
        } else {
            Err(VbaError::Sandbox(format!(
                "{feature} requires permission: {permission:?}"
            )))
        }
    }

    fn copy_range(&mut self, src: VbaRangeRef, dest: VbaRangeRef) -> Result<(), VbaError> {
        let src_rows = src.end_row.saturating_sub(src.start_row) + 1;
        let src_cols = src.end_col.saturating_sub(src.start_col) + 1;
        let dest_rows = dest.end_row.saturating_sub(dest.start_row) + 1;
        let dest_cols = dest.end_col.saturating_sub(dest.start_col) + 1;

        for dr in 0..dest_rows {
            for dc in 0..dest_cols {
                let sr = src.start_row + (dr % src_rows);
                let sc = src.start_col + (dc % src_cols);
                let value = self.sheet.get_cell_value(src.sheet, sr, sc)?;
                let formula = self.sheet.get_cell_formula(src.sheet, sr, sc)?;

                let tr = dest.start_row + dr;
                let tc = dest.start_col + dc;
                self.sheet.set_cell_value(dest.sheet, tr, tc, value)?;
                if let Some(formula) = formula {
                    self.sheet.set_cell_formula(dest.sheet, tr, tc, formula)?;
                }
            }
        }

        Ok(())
    }
}

#[derive(Debug)]
enum ControlFlow {
    Continue,
    ExitSub,
    ExitFunction,
    ExitFor,
    Goto(String),
}

fn collect_labels(body: &[Stmt]) -> HashMap<String, usize> {
    let mut map = HashMap::new();
    for (idx, stmt) in body.iter().enumerate() {
        if let Stmt::Label(label) = stmt {
            map.insert(label.to_ascii_lowercase(), idx);
        }
    }
    map
}
