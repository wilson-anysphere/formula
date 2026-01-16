use std::cell::RefCell;
use std::collections::{HashMap, HashSet};
use std::time::Instant;

use chrono::{Datelike, Duration as ChronoDuration, Local, NaiveDate, NaiveDateTime, Timelike};
use formula_model::column_label_to_index;
use thiserror::Error;

use crate::ast::{
    BinOp, CaseComparisonOp, CaseCondition, Expr, LoopConditionKind, ProcedureDef, ProcedureKind,
    SelectCaseArm, Stmt, UnOp, VarDecl, VbaProgram, VbaType,
};
use crate::object_model::{
    row_col_to_a1, Spreadsheet, VbaErrObject, VbaObject, VbaObjectRef, VbaRangeRef,
};
use crate::sandbox::{Permission, PermissionChecker, VbaSandboxPolicy};
use crate::value::{VbaArray, VbaArrayRef, VbaValue};

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
    pub selection: Option<VbaRangeRef>,
}

#[derive(Debug, Default)]
struct GlobalState {
    values: HashMap<String, VbaValue>,
    types: HashMap<String, VbaType>,
    consts: HashSet<String>,
    const_exprs: HashMap<String, Expr>,
}

pub struct VbaRuntime {
    program: VbaProgram,
    sandbox: VbaSandboxPolicy,
    permission_checker: Option<Box<dyn PermissionChecker>>,
    globals: RefCell<GlobalState>,
}

impl VbaRuntime {
    pub fn new(program: VbaProgram) -> Self {
        let globals = RefCell::new(init_globals(&program));
        Self {
            program,
            sandbox: VbaSandboxPolicy::default(),
            permission_checker: None,
            globals,
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
        self.execute_with_selection(spreadsheet, entry, args, None)
    }

    pub fn execute_with_selection(
        &self,
        spreadsheet: &mut dyn Spreadsheet,
        entry: &str,
        args: &[VbaValue],
        selection: Option<VbaRangeRef>,
    ) -> Result<ExecutionResult, VbaError> {
        let proc = self
            .program
            .get(entry)
            .ok_or_else(|| VbaError::Runtime(format!("Unknown procedure `{entry}`")))?;

        let mut exec = Executor::new(
            &self.program,
            &self.globals,
            spreadsheet,
            &self.sandbox,
            self.permission_checker.as_deref(),
            selection,
        );
        exec.call_procedure(proc, args)
    }

    /// Convenience: execute `Workbook_Open` if present.
    pub fn fire_workbook_open(&self, spreadsheet: &mut dyn Spreadsheet) -> Result<(), VbaError> {
        self.fire_workbook_open_with_selection(spreadsheet, None)
            .map(|_| ())
    }

    pub fn fire_workbook_open_with_selection(
        &self,
        spreadsheet: &mut dyn Spreadsheet,
        selection: Option<VbaRangeRef>,
    ) -> Result<ExecutionResult, VbaError> {
        if self.program.get("workbook_open").is_some() {
            return self.execute_with_selection(spreadsheet, "Workbook_Open", &[], selection);
        }
        Ok(ExecutionResult {
            returned: None,
            selection,
        })
    }

    /// Convenience: execute `Workbook_BeforeClose` if present.
    pub fn fire_workbook_before_close(
        &self,
        spreadsheet: &mut dyn Spreadsheet,
    ) -> Result<(), VbaError> {
        self.fire_workbook_before_close_with_selection(spreadsheet, None)
            .map(|_| ())
    }

    pub fn fire_workbook_before_close_with_selection(
        &self,
        spreadsheet: &mut dyn Spreadsheet,
        selection: Option<VbaRangeRef>,
    ) -> Result<ExecutionResult, VbaError> {
        if self.program.get("workbook_beforeclose").is_some() {
            return self.execute_with_selection(
                spreadsheet,
                "Workbook_BeforeClose",
                &[],
                selection,
            );
        }
        Ok(ExecutionResult {
            returned: None,
            selection,
        })
    }

    /// Fire `Worksheet_Change` if present.
    pub fn fire_worksheet_change(
        &self,
        spreadsheet: &mut dyn Spreadsheet,
        target: VbaRangeRef,
    ) -> Result<(), VbaError> {
        self.fire_worksheet_change_with_selection(spreadsheet, target, None)
            .map(|_| ())
    }

    pub fn fire_worksheet_change_with_selection(
        &self,
        spreadsheet: &mut dyn Spreadsheet,
        target: VbaRangeRef,
        selection: Option<VbaRangeRef>,
    ) -> Result<ExecutionResult, VbaError> {
        if self.program.get("worksheet_change").is_some() {
            return self.execute_with_selection(
                spreadsheet,
                "Worksheet_Change",
                &[VbaValue::Object(VbaObjectRef::new(VbaObject::Range(target)))],
                selection,
            );
        }
        Ok(ExecutionResult {
            returned: None,
            selection,
        })
    }

    /// Fire `Worksheet_SelectionChange` if present.
    pub fn fire_worksheet_selection_change(
        &self,
        spreadsheet: &mut dyn Spreadsheet,
        target: VbaRangeRef,
    ) -> Result<(), VbaError> {
        self.fire_worksheet_selection_change_with_selection(spreadsheet, target, None)
            .map(|_| ())
    }

    pub fn fire_worksheet_selection_change_with_selection(
        &self,
        spreadsheet: &mut dyn Spreadsheet,
        target: VbaRangeRef,
        selection: Option<VbaRangeRef>,
    ) -> Result<ExecutionResult, VbaError> {
        if self.program.get("worksheet_selectionchange").is_some() {
            return self.execute_with_selection(
                spreadsheet,
                "Worksheet_SelectionChange",
                &[VbaValue::Object(VbaObjectRef::new(VbaObject::Range(target)))],
                selection,
            );
        }
        Ok(ExecutionResult {
            returned: None,
            selection,
        })
    }
}

fn init_globals(program: &VbaProgram) -> GlobalState {
    let mut state = GlobalState::default();

    for decl in &program.module_vars {
        let name = decl.name.to_ascii_lowercase();
        state.types.insert(name.clone(), decl.ty);
        state
            .values
            .insert(name, default_value_for_decl(decl).unwrap_or(VbaValue::Empty));
    }

    for decl in &program.module_consts {
        let name = decl.name.to_ascii_lowercase();
        state.consts.insert(name.clone());
        if let Some(ty) = decl.ty {
            state.types.insert(name.clone(), ty);
        }
        if let Some(v) = eval_const_expr(&decl.value) {
            let v = if let Some(ty) = decl.ty {
                coerce_value_to_type(v, ty)
            } else {
                v
            };
            state.values.insert(name.clone(), v);
        } else {
            state.const_exprs.insert(name.clone(), decl.value.clone());
        }
    }

    insert_builtin_excel_constants(&mut state);
    state
}

fn insert_builtin_excel_constants(state: &mut GlobalState) {
    // Common Excel/VBA constants referenced by recorded macros. These are provided by the Excel
    // object library in real VBA and should remain available even when `Option Explicit` is on.
    //
    // We intentionally keep this list small and grow it as we discover new real-world macros.
    const BUILTINS: &[(&str, f64)] = &[
        // PasteSpecial
        ("xlpasteall", -4104.0),
        ("xlpastevalues", -4163.0),
        ("xlpasteformulas", -4123.0),
        ("xlpasteformats", -4122.0),
        // Misc
        ("xlnone", -4142.0),
        // Direction (used by Range.End)
        ("xldown", -4121.0),
        ("xlup", -4162.0),
        ("xltoleft", -4159.0),
        ("xltoright", -4161.0),
    ];

    for (name, value) in BUILTINS {
        if state.values.contains_key(*name) || state.const_exprs.contains_key(*name) {
            continue;
        }
        state.values.insert((*name).to_string(), VbaValue::Double(*value));
        state.consts.insert((*name).to_string());
    }
}

fn default_value_for_decl(decl: &VarDecl) -> Result<VbaValue, VbaError> {
    if decl.dims.is_empty() {
        return Ok(default_value_for_type(decl.ty));
    }

    if decl.dims.len() != 1 {
        return Err(VbaError::Runtime(format!(
            "Only 1D arrays are supported (got {} dims) for `{}`",
            decl.dims.len(),
            decl.name
        )));
    }

    let dim = &decl.dims[0];
    let lower = dim
        .lower
        .as_ref()
        .and_then(eval_const_expr)
        .and_then(|v| v.to_f64())
        .unwrap_or(0.0) as i32;
    let upper = eval_const_expr(&dim.upper)
        .and_then(|v| v.to_f64())
        .ok_or_else(|| {
            VbaError::Runtime(format!(
                "Array bound must be a constant number for `{}`",
                decl.name
            ))
        })? as i32;

    let len = (upper - lower + 1).max(0) as usize;
    let mut values = Vec::with_capacity(len);
    for _ in 0..len {
        values.push(default_value_for_type(decl.ty));
    }

    Ok(VbaValue::Array(std::rc::Rc::new(RefCell::new(VbaArray::new(
        lower, values,
    )))))
}

fn default_value_for_type(ty: VbaType) -> VbaValue {
    match ty {
        VbaType::Variant => VbaValue::Empty,
        VbaType::Integer | VbaType::Long | VbaType::Double => VbaValue::Double(0.0),
        VbaType::String => VbaValue::String(String::new()),
        VbaType::Boolean => VbaValue::Boolean(false),
        VbaType::Date => VbaValue::Date(0.0),
    }
}

fn eval_const_expr(expr: &Expr) -> Option<VbaValue> {
    match expr {
        Expr::Literal(v) => Some(v.clone()),
        Expr::Missing => Some(VbaValue::Empty),
        Expr::Unary { op, expr } => {
            let v = eval_const_expr(expr)?;
            match op {
                UnOp::Neg => Some(VbaValue::Double(-v.to_f64()?)),
                UnOp::Not => Some(VbaValue::Boolean(!v.is_truthy())),
            }
        }
        Expr::Binary { op, left, right } => {
            let l = eval_const_expr(left)?;
            let r = eval_const_expr(right)?;
            match op {
                BinOp::Add => Some(VbaValue::Double(l.to_f64()? + r.to_f64()?)),
                BinOp::Sub => Some(VbaValue::Double(l.to_f64()? - r.to_f64()?)),
                BinOp::Mul => Some(VbaValue::Double(l.to_f64()? * r.to_f64()?)),
                BinOp::Div => Some(VbaValue::Double(l.to_f64()? / r.to_f64()?)),
                BinOp::IntDiv => Some(VbaValue::Double((l.to_f64()? / r.to_f64()?).floor())),
                BinOp::Mod => {
                    let a = l.to_f64()?;
                    let b = r.to_f64()?;
                    Some(VbaValue::Double(a - b * (a / b).floor()))
                }
                BinOp::Pow => Some(VbaValue::Double(l.to_f64()?.powf(r.to_f64()?))),
                BinOp::Concat => Some(VbaValue::String(format!(
                    "{}{}",
                    l.to_string_lossy(),
                    r.to_string_lossy()
                ))),
                BinOp::Eq => Some(VbaValue::Boolean(l == r)),
                BinOp::Ne => Some(VbaValue::Boolean(l != r)),
                BinOp::Lt => Some(VbaValue::Boolean(l.to_f64()? < r.to_f64()?)),
                BinOp::Le => Some(VbaValue::Boolean(l.to_f64()? <= r.to_f64()?)),
                BinOp::Gt => Some(VbaValue::Boolean(l.to_f64()? > r.to_f64()?)),
                BinOp::Ge => Some(VbaValue::Boolean(l.to_f64()? >= r.to_f64()?)),
                BinOp::And => Some(VbaValue::Boolean(l.is_truthy() && r.is_truthy())),
                BinOp::Or => Some(VbaValue::Boolean(l.is_truthy() || r.is_truthy())),
            }
        }
        _ => None,
    }
}

fn coerce_value_to_type(value: VbaValue, ty: VbaType) -> VbaValue {
    match ty {
        VbaType::Variant => value,
        VbaType::String => VbaValue::String(value.to_string_lossy()),
        VbaType::Boolean => VbaValue::Boolean(value.is_truthy()),
        VbaType::Double => VbaValue::Double(value.to_f64().unwrap_or(0.0)),
        VbaType::Integer | VbaType::Long => {
            VbaValue::Double(vba_round_bankers(value.to_f64().unwrap_or(0.0)) as f64)
        }
        VbaType::Date => match value {
            VbaValue::Date(d) => VbaValue::Date(d),
            other => VbaValue::Date(other.to_f64().unwrap_or(0.0)),
        },
    }
}

#[derive(Debug, Clone)]
enum ErrorMode {
    Default,
    ResumeNext,
    GotoLabel(String),
}

#[derive(Debug, Default, Clone, Copy)]
struct ResumeState {
    pc: Option<usize>,
    next_pc: Option<usize>,
}

struct Frame {
    locals: HashMap<String, VbaValue>,
    types: HashMap<String, VbaType>,
    consts: HashSet<String>,
    error_mode: ErrorMode,
    resume: ResumeState,
}

#[derive(Debug, Clone)]
struct Clipboard {
    rows: u32,
    cols: u32,
    values: Vec<VbaValue>,
    formulas: Vec<Option<String>>,
}

struct Executor<'a> {
    program: &'a VbaProgram,
    globals: &'a RefCell<GlobalState>,
    sheet: &'a mut dyn Spreadsheet,
    sandbox: &'a VbaSandboxPolicy,
    permission_checker: Option<&'a dyn PermissionChecker>,
    err_obj: VbaObjectRef,
    with_stack: Vec<VbaObjectRef>,
    clipboard: Option<Clipboard>,
    selection: Option<VbaRangeRef>,
    start: Instant,
    steps: u64,
}

impl<'a> Executor<'a> {
    fn new(
        program: &'a VbaProgram,
        globals: &'a RefCell<GlobalState>,
        sheet: &'a mut dyn Spreadsheet,
        sandbox: &'a VbaSandboxPolicy,
        permission_checker: Option<&'a dyn PermissionChecker>,
        selection: Option<VbaRangeRef>,
    ) -> Self {
        Self {
            program,
            globals,
            sheet,
            sandbox,
            permission_checker,
            err_obj: VbaObjectRef::new(VbaObject::Err(VbaErrObject::default())),
            with_stack: Vec::new(),
            clipboard: None,
            selection,
            start: Instant::now(),
            steps: 0,
        }
    }

    fn tick(&mut self) -> Result<(), VbaError> {
        self.steps = self.steps.saturating_add(1);
        if self.steps > self.sandbox.max_steps {
            return Err(VbaError::StepLimit);
        }
        if (self.steps & 0xFF) == 0 && self.start.elapsed() > self.sandbox.max_execution_time {
            return Err(VbaError::Timeout);
        }
        Ok(())
    }

    fn set_err(&mut self, err: &VbaError) {
        let (number, description) = match err {
            VbaError::Runtime(msg) => (1, msg.clone()),
            VbaError::Sandbox(msg) => (70, msg.clone()),
            VbaError::Timeout => (1, "Execution timed out".to_string()),
            VbaError::StepLimit => (1, "Execution exceeded step limit".to_string()),
            VbaError::Parse(msg) => (1, msg.clone()),
        };

        if let VbaObject::Err(obj) = &mut *self.err_obj.borrow_mut() {
            obj.number = number;
            obj.description = description;
        }
    }

    fn clear_err(&mut self) {
        if let VbaObject::Err(obj) = &mut *self.err_obj.borrow_mut() {
            obj.number = 0;
            obj.description.clear();
        }
    }

    fn call_procedure(
        &mut self,
        proc: &'a ProcedureDef,
        args: &[VbaValue],
    ) -> Result<ExecutionResult, VbaError> {
        let mut frame = Frame {
            locals: HashMap::new(),
            types: HashMap::new(),
            consts: HashSet::new(),
            error_mode: ErrorMode::Default,
            resume: ResumeState::default(),
        };

        let mut function_return_key: Option<String> = None;

        // VBA Functions return by assigning to the function name.
        if proc.kind == ProcedureKind::Function {
            let name = proc.name.to_ascii_lowercase();
            function_return_key = Some(name.clone());
            let default = proc
                .return_type
                .map(default_value_for_type)
                .unwrap_or(VbaValue::Empty);
            frame.locals.insert(name.clone(), default);
            if let Some(ty) = proc.return_type {
                frame.types.insert(name, ty);
            }
        }

        for (idx, param) in proc.params.iter().enumerate() {
            let value = args.get(idx).cloned().unwrap_or(VbaValue::Empty);
            let name = param.name.to_ascii_lowercase();
            frame.locals.insert(name.clone(), value);
            if let Some(ty) = param.ty {
                frame.types.insert(name, ty);
            }
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
        frame
            .locals
            .insert("err".to_string(), VbaValue::Object(self.err_obj.clone()));

        let flow = self.exec_block(&mut frame, &proc.body)?;
        let mut result = ExecutionResult::default();
        match flow {
            ControlFlow::Continue
            | ControlFlow::ExitSub
            | ControlFlow::ExitFor
            | ControlFlow::ExitDo => {}
            ControlFlow::ExitFunction => {}
            ControlFlow::Goto(label) | ControlFlow::ErrorGoto(label) => {
                return Err(VbaError::Runtime(format!(
                    "GoTo `{label}` reached outside of its block"
                )));
            }
            ControlFlow::Resume(_) => {
                return Err(VbaError::Runtime(
                    "Resume reached outside of error handler".to_string(),
                ));
            }
        }

        if proc.kind == ProcedureKind::Function {
            result.returned = Some(
                function_return_key
                    .as_ref()
                    .and_then(|key| frame.locals.get(key))
                    .cloned()
                    .unwrap_or(VbaValue::Empty),
            );
        }
        result.selection = self.selection;
        Ok(result)
    }

    fn exec_block(&mut self, frame: &mut Frame, body: &[Stmt]) -> Result<ControlFlow, VbaError> {
        let label_map = collect_labels(body);
        let mut pc: usize = 0;
        while pc < body.len() {
            self.tick()?;

            let stmt = &body[pc];
            let stmt_res = self.exec_stmt(frame, stmt);

            let mut flow = match stmt_res {
                Ok(flow) => flow,
                Err(err) => match &frame.error_mode {
                    ErrorMode::Default => return Err(err),
                    ErrorMode::ResumeNext => {
                        self.set_err(&err);
                        pc += 1;
                        continue;
                    }
                    ErrorMode::GotoLabel(label) => {
                        self.set_err(&err);
                        frame.resume.pc = Some(pc);
                        frame.resume.next_pc = Some(pc.saturating_add(1));
                        ControlFlow::ErrorGoto(label.clone())
                    }
                },
            };

            match &mut flow {
                ControlFlow::Continue => pc += 1,
                ControlFlow::ExitSub => return Ok(ControlFlow::ExitSub),
                ControlFlow::ExitFunction => return Ok(ControlFlow::ExitFunction),
                ControlFlow::ExitFor => return Ok(ControlFlow::ExitFor),
                ControlFlow::ExitDo => return Ok(ControlFlow::ExitDo),
                ControlFlow::Goto(label) => {
                    if let Some(dest) = label_map.get(label) {
                        pc = *dest;
                    } else {
                        return Ok(ControlFlow::Goto(label.clone()));
                    }
                }
                ControlFlow::ErrorGoto(label) => {
                    if let Some(dest) = label_map.get(label) {
                        pc = *dest;
                    } else {
                        // We're unwinding an error handler to an outer block; adjust resume to
                        // resume after the statement that contains this nested block.
                        frame.resume.pc = Some(pc);
                        frame.resume.next_pc = Some(pc.saturating_add(1));
                        return Ok(ControlFlow::ErrorGoto(label.clone()));
                    }
                }
                ControlFlow::Resume(kind) => {
                    let (target_pc, clear_resume) = match kind {
                        ResumeKind::Next => (frame.resume.next_pc, true),
                        ResumeKind::Same => (frame.resume.pc, true),
                        ResumeKind::Label(label) => {
                            if let Some(dest) = label_map.get(label) {
                                (Some(*dest), true)
                            } else {
                                return Ok(ControlFlow::Resume(kind.clone()));
                            }
                        }
                    };

                    let Some(target_pc) = target_pc else {
                        return Err(VbaError::Runtime("Resume without active error".to_string()));
                    };

                    self.clear_err();
                    if clear_resume {
                        frame.resume = ResumeState::default();
                    }
                    pc = target_pc;
                }
            }
        }
        Ok(ControlFlow::Continue)
    }

    fn exec_stmt(&mut self, frame: &mut Frame, stmt: &Stmt) -> Result<ControlFlow, VbaError> {
        match stmt {
            Stmt::Dim(vars) => {
                for decl in vars {
                    let name = decl.name.to_ascii_lowercase();
                    if frame.locals.contains_key(&name) {
                        continue;
                    }
                    frame.types.insert(name.clone(), decl.ty);
                    frame.locals.insert(name, default_value_for_decl(decl)?);
                }
                Ok(ControlFlow::Continue)
            }
            Stmt::Const(decls) => {
                for decl in decls {
                    let name = decl.name.to_ascii_lowercase();
                    if frame.locals.contains_key(&name) {
                        continue;
                    }
                    let mut value = self.eval_expr(frame, &decl.value)?;
                    if let Some(ty) = decl.ty {
                        value = self.coerce_to_type(frame, value, ty)?;
                        frame.types.insert(name.clone(), ty);
                    }
                    frame.locals.insert(name.clone(), value);
                    frame.consts.insert(name.clone());
                }
                Ok(ControlFlow::Continue)
            }
            Stmt::Assign { target, value } => {
                let rhs = self.eval_expr(frame, value)?;
                // Default member semantics: `x = Range("A1")` reads `Range("A1").Value`, while
                // `Set x = Range("A1")` assigns the object reference.
                let rhs = self.coerce_default_member(frame, rhs)?;
                self.assign(frame, target, rhs)?;
                Ok(ControlFlow::Continue)
            }
            Stmt::Set { target, value } => {
                let rhs = self.eval_expr(frame, value)?;
                self.assign(frame, target, rhs)?;
                Ok(ControlFlow::Continue)
            }
            Stmt::ExprStmt(expr) => {
                match expr {
                    Expr::Member { object, member } => {
                        let obj = self.eval_expr(frame, object)?;
                        let obj = obj
                            .as_object()
                            .ok_or_else(|| VbaError::Runtime("Call on non-object".to_string()))?;
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
                                | "cstr"
                                | "clng"
                                | "cint"
                                | "cdbl"
                                | "cbool"
                                | "cdate"
                                | "format"
                                | "dateadd"
                                | "datediff"
                                | "now"
                                | "date"
                                | "time"
                                | "ucase"
                                | "lcase"
                                | "trim"
                                | "left"
                                | "right"
                                | "mid"
                                | "len"
                                | "replace"
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
                    self.tick()?;
                    self.assign_var(frame, var, VbaValue::Double(i))?;
                    match self.exec_block(frame, body)? {
                        ControlFlow::Continue => {}
                        ControlFlow::ExitFor => break,
                        ControlFlow::ExitDo => break,
                        ControlFlow::ExitSub => return Ok(ControlFlow::ExitSub),
                        ControlFlow::ExitFunction => return Ok(ControlFlow::ExitFunction),
                        ControlFlow::Goto(label) => return Ok(ControlFlow::Goto(label)),
                        ControlFlow::ErrorGoto(label) => return Ok(ControlFlow::ErrorGoto(label)),
                        ControlFlow::Resume(kind) => return Ok(ControlFlow::Resume(kind)),
                    }
                    i += step_v;
                }
                Ok(ControlFlow::Continue)
            }
            Stmt::ForEach {
                var,
                iterable,
                body,
            } => {
                let iterable = self.eval_expr(frame, iterable)?;
                let mut items: Vec<VbaValue> = Vec::new();
                match iterable {
                    VbaValue::Array(arr) => {
                        items.extend(arr.borrow().values.iter().cloned());
                    }
                    VbaValue::Object(obj) => match &*obj.borrow() {
                        VbaObject::Collection { items: coll } => items.extend(coll.iter().cloned()),
                        VbaObject::Dictionary { items: dict } => {
                            items.extend(dict.keys().map(|k| VbaValue::String(k.clone())));
                        }
                        _ => {
                            return Err(VbaError::Runtime(
                                "For Each requires an array or collection".to_string(),
                            ))
                        }
                    },
                    _ => {
                        return Err(VbaError::Runtime(
                            "For Each requires an array or collection".to_string(),
                        ))
                    }
                };

                for item in items {
                    self.tick()?;
                    self.assign_var(frame, var, item)?;
                    match self.exec_block(frame, body)? {
                        ControlFlow::Continue => {}
                        ControlFlow::ExitFor => break,
                        ControlFlow::ExitDo => break,
                        ControlFlow::ExitSub => return Ok(ControlFlow::ExitSub),
                        ControlFlow::ExitFunction => return Ok(ControlFlow::ExitFunction),
                        ControlFlow::Goto(label) => return Ok(ControlFlow::Goto(label)),
                        ControlFlow::ErrorGoto(label) => return Ok(ControlFlow::ErrorGoto(label)),
                        ControlFlow::Resume(kind) => return Ok(ControlFlow::Resume(kind)),
                    }
                }
                Ok(ControlFlow::Continue)
            }
            Stmt::DoLoop {
                pre_condition,
                post_condition,
                body,
            } => {
                if let Some((kind, cond)) = pre_condition {
                    loop {
                        if !self.eval_loop_condition(frame, *kind, cond)? {
                            break;
                        }
                        match self.exec_block(frame, body)? {
                            ControlFlow::Continue => {}
                            ControlFlow::ExitDo => break,
                            ControlFlow::ExitFor => break,
                            ControlFlow::ExitSub => return Ok(ControlFlow::ExitSub),
                            ControlFlow::ExitFunction => return Ok(ControlFlow::ExitFunction),
                            ControlFlow::Goto(label) => return Ok(ControlFlow::Goto(label)),
                            ControlFlow::ErrorGoto(label) => return Ok(ControlFlow::ErrorGoto(label)),
                            ControlFlow::Resume(kind) => return Ok(ControlFlow::Resume(kind)),
                        }
                    }
                    return Ok(ControlFlow::Continue);
                }

                loop {
                    match self.exec_block(frame, body)? {
                        ControlFlow::Continue => {}
                        ControlFlow::ExitDo => break,
                        ControlFlow::ExitFor => break,
                        ControlFlow::ExitSub => return Ok(ControlFlow::ExitSub),
                        ControlFlow::ExitFunction => return Ok(ControlFlow::ExitFunction),
                        ControlFlow::Goto(label) => return Ok(ControlFlow::Goto(label)),
                        ControlFlow::ErrorGoto(label) => return Ok(ControlFlow::ErrorGoto(label)),
                        ControlFlow::Resume(kind) => return Ok(ControlFlow::Resume(kind)),
                    }
                    if let Some((kind, cond)) = post_condition {
                        if !self.eval_loop_condition(frame, *kind, cond)? {
                            break;
                        }
                    }
                }
                Ok(ControlFlow::Continue)
            }
            Stmt::While { cond, body } => {
                while self.eval_expr(frame, cond)?.is_truthy() {
                    match self.exec_block(frame, body)? {
                        ControlFlow::Continue => {}
                        ControlFlow::ExitDo => break,
                        ControlFlow::ExitFor => break,
                        ControlFlow::ExitSub => return Ok(ControlFlow::ExitSub),
                        ControlFlow::ExitFunction => return Ok(ControlFlow::ExitFunction),
                        ControlFlow::Goto(label) => return Ok(ControlFlow::Goto(label)),
                        ControlFlow::ErrorGoto(label) => return Ok(ControlFlow::ErrorGoto(label)),
                        ControlFlow::Resume(kind) => return Ok(ControlFlow::Resume(kind)),
                    }
                }
                Ok(ControlFlow::Continue)
            }
            Stmt::SelectCase {
                expr,
                cases,
                else_body,
            } => {
                let selector = self.eval_expr(frame, expr)?;
                for arm in cases {
                    if self.select_case_matches(frame, &selector, arm)? {
                        return self.exec_block(frame, &arm.body);
                    }
                }
                self.exec_block(frame, else_body)
            }
            Stmt::With { object, body } => {
                let obj = self.eval_expr(frame, object)?;
                let obj = obj.as_object().ok_or_else(|| {
                    VbaError::Runtime("With expression must evaluate to an object".to_string())
                })?;
                self.with_stack.push(obj);
                let res = self.exec_block(frame, body);
                self.with_stack.pop();
                res
            }
            Stmt::ExitSub => Ok(ControlFlow::ExitSub),
            Stmt::ExitFunction => Ok(ControlFlow::ExitFunction),
            Stmt::ExitFor => Ok(ControlFlow::ExitFor),
            Stmt::ExitDo => Ok(ControlFlow::ExitDo),
            Stmt::OnErrorResumeNext => {
                frame.error_mode = ErrorMode::ResumeNext;
                Ok(ControlFlow::Continue)
            }
            Stmt::OnErrorGoto0 => {
                frame.error_mode = ErrorMode::Default;
                self.clear_err();
                Ok(ControlFlow::Continue)
            }
            Stmt::OnErrorGotoLabel(label) => {
                frame.error_mode = ErrorMode::GotoLabel(label.to_ascii_lowercase());
                Ok(ControlFlow::Continue)
            }
            Stmt::ResumeNext => Ok(ControlFlow::Resume(ResumeKind::Next)),
            Stmt::Resume => Ok(ControlFlow::Resume(ResumeKind::Same)),
            Stmt::ResumeLabel(label) => Ok(ControlFlow::Resume(ResumeKind::Label(
                label.to_ascii_lowercase(),
            ))),
            Stmt::Label(_) => Ok(ControlFlow::Continue),
            Stmt::Goto(label) => Ok(ControlFlow::Goto(label.to_ascii_lowercase())),
        }
    }

    fn eval_loop_condition(
        &mut self,
        frame: &mut Frame,
        kind: LoopConditionKind,
        cond: &Expr,
    ) -> Result<bool, VbaError> {
        let v = self.eval_expr(frame, cond)?.is_truthy();
        Ok(match kind {
            LoopConditionKind::While => v,
            LoopConditionKind::Until => !v,
        })
    }

    fn select_case_matches(
        &mut self,
        frame: &mut Frame,
        selector: &VbaValue,
        arm: &SelectCaseArm,
    ) -> Result<bool, VbaError> {
        for cond in &arm.conditions {
            if self.select_case_condition_matches(frame, selector, cond)? {
                return Ok(true);
            }
        }
        Ok(false)
    }

    fn select_case_condition_matches(
        &mut self,
        frame: &mut Frame,
        selector: &VbaValue,
        cond: &CaseCondition,
    ) -> Result<bool, VbaError> {
        match cond {
            CaseCondition::Expr(expr) => {
                let v = self.eval_expr(frame, expr)?;
                let eq = self.eval_binop(BinOp::Eq, selector.clone(), v)?;
                Ok(eq.is_truthy())
            }
            CaseCondition::Range { start, end } => {
                let s = self.eval_expr(frame, start)?.to_f64().unwrap_or(0.0);
                let e = self.eval_expr(frame, end)?.to_f64().unwrap_or(0.0);
                let v = selector.to_f64().unwrap_or(0.0);
                Ok(v >= s.min(e) && v <= s.max(e))
            }
            CaseCondition::Is { op, expr } => {
                let rhs = self.eval_expr(frame, expr)?;
                let bin = match op {
                    CaseComparisonOp::Eq => BinOp::Eq,
                    CaseComparisonOp::Ne => BinOp::Ne,
                    CaseComparisonOp::Lt => BinOp::Lt,
                    CaseComparisonOp::Le => BinOp::Le,
                    CaseComparisonOp::Gt => BinOp::Gt,
                    CaseComparisonOp::Ge => BinOp::Ge,
                };
                let v = self.eval_binop(bin, selector.clone(), rhs)?;
                Ok(v.is_truthy())
            }
        }
    }

    fn coerce_to_type(
        &mut self,
        frame: &mut Frame,
        value: VbaValue,
        ty: VbaType,
    ) -> Result<VbaValue, VbaError> {
        if matches!(value, VbaValue::Null) {
            return Err(VbaError::Runtime("Invalid use of Null".to_string()));
        }

        let v = match ty {
            VbaType::Variant => value,
            VbaType::String => VbaValue::String(self.coerce_to_string(frame, value)?),
            VbaType::Boolean => VbaValue::Boolean(self.coerce_to_bool(frame, value)?),
            VbaType::Double => VbaValue::Double(self.coerce_to_f64(frame, value)?),
            VbaType::Integer | VbaType::Long => {
                let n = self.coerce_to_f64(frame, value)?;
                VbaValue::Double(vba_round_bankers(n) as f64)
            }
            VbaType::Date => VbaValue::Date(self.coerce_to_date_serial(frame, value)?),
        };
        Ok(v)
    }

    fn coerce_to_string(&mut self, frame: &mut Frame, value: VbaValue) -> Result<String, VbaError> {
        let value = self.coerce_default_member(frame, value)?;
        match value {
            VbaValue::Null => Err(VbaError::Runtime("Invalid use of Null".to_string())),
            other => Ok(other.to_string_lossy()),
        }
    }

    fn coerce_to_f64(&mut self, frame: &mut Frame, value: VbaValue) -> Result<f64, VbaError> {
        let value = self.coerce_default_member(frame, value)?;
        match value {
            VbaValue::Null => Err(VbaError::Runtime("Invalid use of Null".to_string())),
            VbaValue::String(s) => s.parse::<f64>().map_err(|_| {
                VbaError::Runtime(format!("Type mismatch: cannot coerce `{s}` to number"))
            }),
            other => Ok(other.to_f64().unwrap_or(0.0)),
        }
    }

    fn coerce_cells_index(
        &mut self,
        frame: &mut Frame,
        value: VbaValue,
        is_row: bool,
    ) -> Result<u32, VbaError> {
        let value = self.coerce_default_member(frame, value)?;
        let raw = match value {
            VbaValue::Empty => return Ok(1),
            VbaValue::Null => return Err(VbaError::Runtime("Invalid use of Null".to_string())),
            VbaValue::String(s) => {
                let s = s.trim();
                if let Ok(n) = s.parse::<f64>() {
                    n
                } else if !is_row {
                    // VBA allows `Cells(1, "B")` where the column is given as letters.
                    let col0 = column_label_to_index(s).map_err(|_| {
                        VbaError::Runtime(format!(
                            "Invalid column index `{s}` (expected number or Excel column label)"
                        ))
                    })?;
                    return Ok(col0 + 1);
                } else {
                    return Err(VbaError::Runtime(format!(
                        "Invalid row index `{s}` (expected a number)"
                    )));
                }
            }
            other => other.to_f64().unwrap_or(0.0),
        };

        let idx = raw as i64;
        if idx <= 0 {
            return Err(VbaError::Runtime(format!(
                "Cells index must be >= 1 (got {raw})"
            )));
        }
        Ok(idx as u32)
    }

    fn coerce_to_bool(&mut self, frame: &mut Frame, value: VbaValue) -> Result<bool, VbaError> {
        let value = self.coerce_default_member(frame, value)?;
        match value {
            VbaValue::Null => Err(VbaError::Runtime("Invalid use of Null".to_string())),
            VbaValue::Boolean(b) => Ok(b),
            VbaValue::String(s) => {
                let s_trim = s.trim();
                if s_trim.eq_ignore_ascii_case("true") {
                    return Ok(true);
                }
                if s_trim.eq_ignore_ascii_case("false") {
                    return Ok(false);
                }
                if s_trim.is_empty() {
                    return Ok(false);
                }
                let n = s_trim.parse::<f64>().map_err(|_| {
                    VbaError::Runtime(format!("Type mismatch: cannot coerce `{s}` to Boolean"))
                })?;
                Ok(n != 0.0)
            }
            other => Ok(other.is_truthy()),
        }
    }

    fn coerce_to_date_serial(
        &mut self,
        frame: &mut Frame,
        value: VbaValue,
    ) -> Result<f64, VbaError> {
        let value = self.coerce_default_member(frame, value)?;
        match value {
            VbaValue::Date(d) => Ok(d),
            VbaValue::Double(n) => Ok(n),
            VbaValue::String(s) => parse_vba_date_string(&s).map(datetime_to_ole_date).ok_or_else(|| {
                VbaError::Runtime(format!("Type mismatch: cannot coerce `{s}` to Date"))
            }),
            VbaValue::Empty => Ok(0.0),
            VbaValue::Null => Err(VbaError::Runtime("Invalid use of Null".to_string())),
            other => Ok(other.to_f64().unwrap_or(0.0)),
        }
    }

    fn assign(&mut self, frame: &mut Frame, target: &Expr, value: VbaValue) -> Result<(), VbaError> {
        match target {
            Expr::Var(name) => self.assign_var(frame, name, value),
            Expr::Member { object, member } => {
                let obj = self.eval_expr(frame, object)?;
                let obj = obj.as_object().ok_or_else(|| {
                    VbaError::Runtime("Assignment target is not an object".to_string())
                })?;
                self.set_object_member(obj, member, value)
            }
            Expr::Call { callee, args } => {
                // Array element assignment: `arr(i) = v`
                if let Expr::Var(name) = &**callee {
                    let name_lc = name.to_ascii_lowercase();
                    if let Some(VbaValue::Array(arr)) = frame.locals.get(&name_lc).cloned() {
                        let idx_expr = args
                            .first()
                            .ok_or_else(|| VbaError::Runtime("Array index missing".to_string()))?;
                        let idx = self
                            .eval_expr(frame, &idx_expr.expr)?
                            .to_f64()
                            .unwrap_or(0.0) as i32;
                        if let Some(slot) = arr.borrow_mut().get_mut(idx) {
                            *slot = value;
                            return Ok(());
                        }
                        return Err(VbaError::Runtime("Subscript out of range".to_string()));
                    }
                }

                // Default property assignment: `Range("A1") = 1`
                let obj_val = self.eval_expr(frame, target)?;
                let obj = obj_val.as_object().ok_or_else(|| {
                    VbaError::Runtime("Unsupported assignment target".to_string())
                })?;
                self.set_object_member(obj, "Value", value)
            }
            _ => Err(VbaError::Runtime(
                "Unsupported assignment target".to_string(),
            )),
        }
    }

    fn assign_var(&mut self, frame: &mut Frame, name: &str, value: VbaValue) -> Result<(), VbaError> {
        let name_lc = name.to_ascii_lowercase();

        if frame.consts.contains(&name_lc) {
            return Err(VbaError::Runtime(format!(
                "Cannot assign to constant `{name}`"
            )));
        }
        if self.globals.borrow().consts.contains(&name_lc) {
            return Err(VbaError::Runtime(format!(
                "Cannot assign to constant `{name}`"
            )));
        }

        let declared_ty = frame
            .types
            .get(&name_lc)
            .copied()
            .or_else(|| self.globals.borrow().types.get(&name_lc).copied());

        let value = if let Some(ty) = declared_ty {
            self.coerce_to_type(frame, value, ty)?
        } else {
            value
        };

        if let Some(slot) = frame.locals.get_mut(&name_lc) {
            *slot = value;
            return Ok(());
        }

        if self.globals.borrow().values.contains_key(&name_lc) {
            self.globals.borrow_mut().values.insert(name_lc, value);
            return Ok(());
        }

        if self.program.option_explicit {
            return Err(VbaError::Runtime(format!(
                "Variable not defined: `{name}`"
            )));
        }

        frame.locals.insert(name_lc, value);
        Ok(())
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
                "value" | "value2" => self.set_range_value(*range, value),
                "formula" => self.set_range_formula(*range, value),
                "formular1c1" => self.set_range_formula(*range, value),
                _ => Err(VbaError::Runtime(format!(
                    "Unknown Range member `{member}`"
                ))),
            },
            VbaObject::Application => match member_lc.as_str() {
                "cutcopymode" => {
                    // Excel uses `Application.CutCopyMode = False` to clear the clipboard and
                    // exit copy mode. We model this by clearing our internal clipboard.
                    if !value.is_truthy() {
                        self.clipboard = None;
                    }
                    Ok(())
                }
                _ => Err(VbaError::Runtime(format!(
                    "Cannot assign to Application member `{member}`"
                ))),
            },
            VbaObject::Err(err) => match member_lc.as_str() {
                "number" => {
                    err.number = value.to_f64().unwrap_or(0.0) as i32;
                    Ok(())
                }
                "description" => {
                    err.description = value.to_string_lossy();
                    Ok(())
                }
                _ => Err(VbaError::Runtime(format!("Unknown Err member `{member}`"))),
            },
            _ => Err(VbaError::Runtime(format!(
                "Cannot assign to member `{member}`"
            ))),
        }
    }

    fn eval_expr(&mut self, frame: &mut Frame, expr: &Expr) -> Result<VbaValue, VbaError> {
        self.tick()?;
        match expr {
            Expr::Literal(v) => Ok(v.clone()),
            Expr::Missing => Ok(VbaValue::Empty),
            Expr::With => self
                .with_stack
                .last()
                .cloned()
                .map(VbaValue::Object)
                .ok_or_else(|| VbaError::Runtime("`.` used outside of `With`".to_string())),
            Expr::Var(name) => self.eval_var(frame, name),
            Expr::Unary { op, expr } => {
                let v = self.eval_expr(frame, expr)?;
                match op {
                    UnOp::Neg => Ok(VbaValue::Double(-self.coerce_to_f64(frame, v)?)),
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
            Expr::Call { callee, args } => self.eval_call(frame, callee, args),
            Expr::Index { .. } => Err(VbaError::Runtime("Indexing not implemented".to_string())),
        }
    }

    fn eval_var(&mut self, frame: &mut Frame, name: &str) -> Result<VbaValue, VbaError> {
        let name_lc = name.to_ascii_lowercase();

        if let Some(v) = frame.locals.get(&name_lc).cloned() {
            return Ok(v);
        }

        // Lazy constant eval.
        if self.globals.borrow().const_exprs.contains_key(&name_lc) {
            let expr = self
                .globals
                .borrow_mut()
                .const_exprs
                .remove(&name_lc)
                .unwrap();
            let value = self.eval_expr(frame, &expr)?;
            self.globals.borrow_mut().values.insert(name_lc.clone(), value.clone());
            return Ok(value);
        }

        if let Some(v) = self.globals.borrow().values.get(&name_lc).cloned() {
            return Ok(v);
        }

        // Built-in global properties without explicit object.
        if name_lc == "activesheet" {
            return Ok(VbaValue::Object(VbaObjectRef::new(VbaObject::Worksheet {
                sheet: self.sheet.active_sheet(),
            })));
        }
        if name_lc == "activecell" {
            let (r, c) = self.sheet.active_cell();
            return Ok(VbaValue::Object(VbaObjectRef::new(VbaObject::Range(
                VbaRangeRef {
                    sheet: self.sheet.active_sheet(),
                    start_row: r,
                    start_col: c,
                    end_row: r,
                    end_col: c,
                },
            ))));
        }
        if name_lc == "selection" {
            let sel = self.selection.unwrap_or_else(|| {
                let (r, c) = self.sheet.active_cell();
                VbaRangeRef {
                    sheet: self.sheet.active_sheet(),
                    start_row: r,
                    start_col: c,
                    end_row: r,
                    end_col: c,
                }
            });
            return Ok(VbaValue::Object(VbaObjectRef::new(VbaObject::Range(sel))));
        }
        if name_lc == "cells" {
            let sheet = self.sheet.active_sheet();
            return Ok(VbaValue::Object(VbaObjectRef::new(VbaObject::Range(
                sheet_entire_range(self.sheet, sheet),
            ))));
        }
        if name_lc == "rows" {
            let sheet = self.sheet.active_sheet();
            return Ok(VbaValue::Object(VbaObjectRef::new(VbaObject::RangeRows {
                range: sheet_entire_range(self.sheet, sheet),
            })));
        }
        if name_lc == "columns" {
            let sheet = self.sheet.active_sheet();
            return Ok(VbaValue::Object(VbaObjectRef::new(VbaObject::RangeColumns {
                range: sheet_entire_range(self.sheet, sheet),
            })));
        }
        if name_lc == "activeworkbook" {
            return Ok(VbaValue::Object(VbaObjectRef::new(VbaObject::Workbook)));
        }

        // VBA allows calling some 0-argument functions without parentheses (e.g. `Now`, `Date`).
        if matches!(name_lc.as_str(), "now" | "date" | "time") {
            return self.call_global(frame, name, &[]);
        }

        // User-defined no-arg functions can also be called without parentheses.
        if let Some(proc) = self.program.get(&name_lc) {
            if proc.kind == ProcedureKind::Function && proc.params.is_empty() {
                let res = self.call_procedure(proc, &[])?;
                return Ok(res.returned.unwrap_or(VbaValue::Empty));
            }
        }

        if self.program.option_explicit {
            Err(VbaError::Runtime(format!(
                "Variable not defined: `{name}`"
            )))
        } else {
            Ok(VbaValue::Empty)
        }
    }

    fn eval_call(
        &mut self,
        frame: &mut Frame,
        callee: &Expr,
        args: &[crate::ast::CallArg],
    ) -> Result<VbaValue, VbaError> {
        if let Expr::Var(name) = callee {
            let name_lc = name.to_ascii_lowercase();
            if let Some(VbaValue::Array(arr)) = frame.locals.get(&name_lc).cloned() {
                return self.index_array_value(frame, VbaValue::Array(arr), args);
            }
            return self.call_global(frame, name, args);
        }

        if let Expr::Member { object, member } = callee {
            let obj = self.eval_expr(frame, object)?;
            let obj = obj
                .as_object()
                .ok_or_else(|| VbaError::Runtime("Call on non-object".to_string()))?;
            return self.call_object_method(frame, obj, member, args);
        }

        // Default member call for objects / array indexing.
        let callee_val = self.eval_expr(frame, callee)?;
        if let VbaValue::Array(arr) = callee_val {
            return self.index_array_value(frame, VbaValue::Array(arr), args);
        }

        if let VbaValue::Object(obj) = callee_val {
            // Collection/Dictionary default member is Item; Worksheet default is Range.
            let snapshot = obj.borrow().clone();
            match snapshot {
                VbaObject::Collection { .. } => {
                    return self.call_object_method(frame, obj, "Item", args);
                }
                VbaObject::Dictionary { .. } => {
                    return self.call_object_method(frame, obj, "Item", args);
                }
                VbaObject::Worksheet { .. } => {
                    return self.call_object_method(frame, obj, "Range", args);
                }
                _ => {}
            }
        }

        Err(VbaError::Runtime("Unsupported call target".to_string()))
    }

    fn coerce_default_member(
        &mut self,
        _frame: &mut Frame,
        value: VbaValue,
    ) -> Result<VbaValue, VbaError> {
        match value {
            VbaValue::Object(obj) => {
                let is_range = { matches!(&*obj.borrow(), VbaObject::Range(_)) };
                if is_range {
                    self.get_object_member(obj, "Value")
                } else {
                    Ok(VbaValue::Object(obj))
                }
            }
            other => Ok(other),
        }
    }

    fn eval_binop(&mut self, op: BinOp, l: VbaValue, r: VbaValue) -> Result<VbaValue, VbaError> {
        let l = self.coerce_default_member(&mut Frame::dummy(), l)?;
        let r = self.coerce_default_member(&mut Frame::dummy(), r)?;

        if matches!(l, VbaValue::Null) || matches!(r, VbaValue::Null) {
            // Propagate Null (best-effort).
            return Ok(VbaValue::Null);
        }

        match op {
            BinOp::Add => Ok(VbaValue::Double(l.to_f64().unwrap_or(0.0) + r.to_f64().unwrap_or(0.0))),
            BinOp::Sub => Ok(VbaValue::Double(l.to_f64().unwrap_or(0.0) - r.to_f64().unwrap_or(0.0))),
            BinOp::Mul => Ok(VbaValue::Double(l.to_f64().unwrap_or(0.0) * r.to_f64().unwrap_or(0.0))),
            BinOp::Div => Ok(VbaValue::Double(l.to_f64().unwrap_or(0.0) / r.to_f64().unwrap_or(0.0))),
            BinOp::IntDiv => Ok(VbaValue::Double((l.to_f64().unwrap_or(0.0) / r.to_f64().unwrap_or(0.0)).floor())),
            BinOp::Mod => {
                let a = l.to_f64().unwrap_or(0.0);
                let b = r.to_f64().unwrap_or(0.0);
                Ok(VbaValue::Double(a - b * (a / b).floor()))
            }
            BinOp::Pow => Ok(VbaValue::Double(l.to_f64().unwrap_or(0.0).powf(r.to_f64().unwrap_or(0.0)))),
            BinOp::Concat => Ok(VbaValue::String(format!("{}{}", l.to_string_lossy(), r.to_string_lossy()))),
            BinOp::Eq | BinOp::Ne | BinOp::Lt | BinOp::Le | BinOp::Gt | BinOp::Ge => {
                self.eval_compare(op, l, r)
            }
            BinOp::And => Ok(VbaValue::Boolean(l.is_truthy() && r.is_truthy())),
            BinOp::Or => Ok(VbaValue::Boolean(l.is_truthy() || r.is_truthy())),
        }
    }

    fn eval_compare(&self, op: BinOp, l: VbaValue, r: VbaValue) -> Result<VbaValue, VbaError> {
        // Best-effort VBA-style coercion:
        // - Prefer numeric compare when both sides can coerce to number.
        // - Otherwise fall back to case-insensitive string compare.
        let ln = l.to_f64();
        let rn = r.to_f64();
        if let (Some(a), Some(b)) = (ln, rn) {
            let res = match op {
                BinOp::Eq => a == b,
                BinOp::Ne => a != b,
                BinOp::Lt => a < b,
                BinOp::Le => a <= b,
                BinOp::Gt => a > b,
                BinOp::Ge => a >= b,
                _ => unreachable!(),
            };
            return Ok(VbaValue::Boolean(res));
        }

        let a = l.to_string_lossy();
        let b = r.to_string_lossy();
        let cmp = ascii_lowercase_cmp(a.as_str(), b.as_str());
        let res = match op {
            BinOp::Eq => cmp == std::cmp::Ordering::Equal,
            BinOp::Ne => cmp != std::cmp::Ordering::Equal,
            BinOp::Lt => cmp == std::cmp::Ordering::Less,
            BinOp::Le => cmp != std::cmp::Ordering::Greater,
            BinOp::Gt => cmp == std::cmp::Ordering::Greater,
            BinOp::Ge => cmp != std::cmp::Ordering::Less,
            _ => unreachable!(),
        };
        Ok(VbaValue::Boolean(res))
    }

    fn call_global(
        &mut self,
        frame: &mut Frame,
        name: &str,
        args: &[crate::ast::CallArg],
    ) -> Result<VbaValue, VbaError> {
        match () {
            _ if name.eq_ignore_ascii_case("range") => {
                if args.is_empty() {
                    return Err(VbaError::Runtime("Range() missing argument".to_string()));
                }

                if args.len() == 1 {
                    let cell1 = arg_named_or_pos(args, "cell1", 0).unwrap_or(&args[0]);
                    if let Expr::Literal(VbaValue::String(a1)) = &cell1.expr {
                        return Ok(VbaValue::Object(crate::object_model::range_on_active_sheet(
                            self.sheet, a1,
                        )?));
                    }
                }

                let sheet = self.sheet.active_sheet();
                let range = self.eval_range_args(frame, sheet, args)?;
                Ok(VbaValue::Object(VbaObjectRef::new(VbaObject::Range(range))))
            }
            _ if name.eq_ignore_ascii_case("cells") => {
                let sheet = self.sheet.active_sheet();
                if args.is_empty() {
                    return Ok(VbaValue::Object(VbaObjectRef::new(VbaObject::Range(
                        sheet_entire_range(self.sheet, sheet),
                    ))));
                }
                if args.len() < 2 {
                    return Err(VbaError::Runtime("Cells() missing arguments".to_string()));
                }
                let row = self.eval_expr(frame, &args[0].expr)?;
                let col = self.eval_expr(frame, &args[1].expr)?;
                let row = self.coerce_cells_index(frame, row, true)?;
                let col = self.coerce_cells_index(frame, col, false)?;
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
            _ if name.eq_ignore_ascii_case("rows") => {
                if args.is_empty() {
                    let sheet = self.sheet.active_sheet();
                    return Ok(VbaValue::Object(VbaObjectRef::new(VbaObject::RangeRows {
                        range: sheet_entire_range(self.sheet, sheet),
                    })));
                }
                let sheet = self.sheet.active_sheet();
                let v = self.eval_expr(frame, &args[0].expr)?;
                let range = match v {
                    VbaValue::String(s) => self.sheet_range(sheet, &s)?,
                    other => {
                        let row = self.coerce_cells_index(frame, other, true)?;
                        let mut range = sheet_entire_range(self.sheet, sheet);
                        range.start_row = row;
                        range.end_row = row;
                        range
                    }
                };
                Ok(VbaValue::Object(VbaObjectRef::new(VbaObject::Range(range))))
            }
            _ if name.eq_ignore_ascii_case("columns") => {
                if args.is_empty() {
                    let sheet = self.sheet.active_sheet();
                    return Ok(VbaValue::Object(VbaObjectRef::new(VbaObject::RangeColumns {
                        range: sheet_entire_range(self.sheet, sheet),
                    })));
                }
                let sheet = self.sheet.active_sheet();
                let v = self.eval_expr(frame, &args[0].expr)?;
                let range = match v {
                    VbaValue::String(s) => self.sheet_range(sheet, &s)?,
                    other => {
                        let col = self.coerce_cells_index(frame, other, false)?;
                        let mut range = sheet_entire_range(self.sheet, sheet);
                        range.start_col = col;
                        range.end_col = col;
                        range
                    }
                };
                Ok(VbaValue::Object(VbaObjectRef::new(VbaObject::Range(range))))
            }
            _ if name.eq_ignore_ascii_case("msgbox") => {
                let msg = self
                    .eval_expr(
                        frame,
                        &args
                            .first()
                            .ok_or_else(|| VbaError::Runtime("MsgBox() missing prompt".to_string()))?
                            .expr,
                    )?
                    .to_string_lossy();
                self.sheet.log(format!("MsgBox: {msg}"));
                Ok(VbaValue::Double(0.0))
            }
            _ if name.eq_ignore_ascii_case("debugprint") => {
                let mut parts = Vec::new();
                for arg in args {
                    parts.push(self.eval_expr(frame, &arg.expr)?.to_string_lossy());
                }
                self.sheet.log(format!("Debug.Print {}", parts.join(" ")));
                Ok(VbaValue::Empty)
            }
            _ if name.eq_ignore_ascii_case("array") => {
                let mut values = Vec::new();
                for arg in args {
                    values.push(self.eval_expr(frame, &arg.expr)?);
                }
                Ok(VbaValue::Array(std::rc::Rc::new(RefCell::new(
                    VbaArray::new(0, values),
                ))))
            }
            _ if name.eq_ignore_ascii_case("worksheets") || name.eq_ignore_ascii_case("sheets") => {
                let arg = self.eval_expr(
                    frame,
                    &args
                        .first()
                        .ok_or_else(|| VbaError::Runtime("Worksheets() missing name".to_string()))?
                        .expr,
                )?;
                let idx = match arg {
                    VbaValue::String(name) => self
                        .sheet
                        .sheet_index(&name)
                        .ok_or_else(|| VbaError::Runtime(format!("Unknown worksheet `{name}`")))?,
                    other => {
                        let n = other.to_f64().unwrap_or(0.0) as isize;
                        if n <= 0 {
                            return Err(VbaError::Runtime(format!(
                                "Worksheet index must be >= 1 (got {n})"
                            )));
                        }
                        let idx = (n - 1) as usize;
                        if idx >= self.sheet.sheet_count() {
                            return Err(VbaError::Runtime(format!(
                                "Worksheet index out of range: {n}"
                            )));
                        }
                        idx
                    }
                };
                Ok(VbaValue::Object(VbaObjectRef::new(VbaObject::Worksheet {
                    sheet: idx,
                })))
            }
            _ if name.eq_ignore_ascii_case("__new") => {
                let class = self
                    .eval_expr(
                        frame,
                        &args
                            .first()
                            .ok_or_else(|| VbaError::Runtime("__new() missing class".to_string()))?
                            .expr,
                    )?
                    .to_string_lossy();
                if class.eq_ignore_ascii_case("collection") {
                    return Ok(VbaValue::Object(VbaObjectRef::new(VbaObject::Collection {
                        items: Vec::new(),
                    })));
                }
                Err(VbaError::Runtime(format!(
                    "Unsupported class in New: {class}"
                )))
            }
            _ if name.eq_ignore_ascii_case("createobject") => {
                let progid = self
                    .eval_expr(
                        frame,
                        &args
                            .first()
                            .ok_or_else(|| {
                                VbaError::Runtime("CreateObject() missing ProgID".to_string())
                            })?
                            .expr,
                    )?
                    .to_string_lossy();

                if progid.eq_ignore_ascii_case("scripting.dictionary")
                    || progid.eq_ignore_ascii_case("scripting.dictionary.1")
                {
                    self.require_permission(Permission::ObjectCreation, "CreateObject")?;
                    return Ok(VbaValue::Object(VbaObjectRef::new(VbaObject::Dictionary {
                        items: HashMap::new(),
                    })));
                }

                self.require_permission(Permission::ObjectCreation, "CreateObject")?;
                Err(VbaError::Sandbox(format!(
                    "CreateObject is not supported for `{progid}`"
                )))
            }
            // ---- Conversions / string helpers ----
            _ if name.eq_ignore_ascii_case("cstr") => {
                let v = self.eval_required_arg(frame, args, 0, "CStr")?;
                if matches!(v, VbaValue::Null) {
                    return Err(VbaError::Runtime("Invalid use of Null".to_string()));
                }
                Ok(VbaValue::String(v.to_string_lossy()))
            }
            _ if name.eq_ignore_ascii_case("clng")
                || name.eq_ignore_ascii_case("cint")
                || name.eq_ignore_ascii_case("cdbl") =>
            {
                let v = self.eval_required_arg(frame, args, 0, name)?;
                if matches!(v, VbaValue::Null) {
                    return Err(VbaError::Runtime("Invalid use of Null".to_string()));
                }
                let n = self.coerce_to_f64(frame, v)?;
                let out = if name.eq_ignore_ascii_case("cdbl") {
                    n
                } else {
                    vba_round_bankers(n) as f64
                };
                Ok(VbaValue::Double(out))
            }
            _ if name.eq_ignore_ascii_case("cbool") => {
                let v = self.eval_required_arg(frame, args, 0, "CBool")?;
                Ok(VbaValue::Boolean(self.coerce_to_bool(frame, v)?))
            }
            _ if name.eq_ignore_ascii_case("cdate") => {
                let v = self.eval_required_arg(frame, args, 0, "CDate")?;
                Ok(VbaValue::Date(self.coerce_to_date_serial(frame, v)?))
            }
            _ if name.eq_ignore_ascii_case("ucase") => {
                let s = self.eval_required_arg(frame, args, 0, "UCase")?.to_string_lossy();
                Ok(VbaValue::String(s.to_uppercase()))
            }
            _ if name.eq_ignore_ascii_case("lcase") => {
                let s = self.eval_required_arg(frame, args, 0, "LCase")?.to_string_lossy();
                Ok(VbaValue::String(s.to_lowercase()))
            }
            _ if name.eq_ignore_ascii_case("trim") => {
                let s = self.eval_required_arg(frame, args, 0, "Trim")?.to_string_lossy();
                Ok(VbaValue::String(s.trim().to_string()))
            }
            _ if name.eq_ignore_ascii_case("left") => {
                let s = self.eval_required_arg(frame, args, 0, "Left")?.to_string_lossy();
                let n = self
                    .eval_required_arg(frame, args, 1, "Left")?
                    .to_f64()
                    .unwrap_or(0.0) as usize;
                Ok(VbaValue::String(s.chars().take(n).collect()))
            }
            _ if name.eq_ignore_ascii_case("right") => {
                let s = self.eval_required_arg(frame, args, 0, "Right")?.to_string_lossy();
                let n = self
                    .eval_required_arg(frame, args, 1, "Right")?
                    .to_f64()
                    .unwrap_or(0.0) as usize;
                let len = s.chars().count();
                Ok(VbaValue::String(s.chars().skip(len.saturating_sub(n)).collect()))
            }
            _ if name.eq_ignore_ascii_case("mid") => {
                let s = self.eval_required_arg(frame, args, 0, "Mid")?.to_string_lossy();
                let start = self
                    .eval_required_arg(frame, args, 1, "Mid")?
                    .to_f64()
                    .unwrap_or(1.0) as isize;
                let start = (start - 1).max(0) as usize;
                let chars: Vec<char> = s.chars().collect();
                let out = match args.get(2) {
                    None => chars.into_iter().skip(start).collect::<String>(),
                    Some(len_arg) if matches!(len_arg.expr, Expr::Missing) => {
                        chars.into_iter().skip(start).collect::<String>()
                    }
                    Some(len_arg) => {
                        let len =
                            self.eval_expr(frame, &len_arg.expr)?.to_f64().unwrap_or(0.0) as usize;
                        chars
                            .into_iter()
                            .skip(start)
                            .take(len)
                            .collect::<String>()
                    }
                };
                Ok(VbaValue::String(out))
            }
            _ if name.eq_ignore_ascii_case("len") => {
                let s = self.eval_required_arg(frame, args, 0, "Len")?.to_string_lossy();
                Ok(VbaValue::Double(s.chars().count() as f64))
            }
            _ if name.eq_ignore_ascii_case("replace") => {
                let expr = self.eval_required_arg(frame, args, 0, "Replace")?.to_string_lossy();
                let find = self.eval_required_arg(frame, args, 1, "Replace")?.to_string_lossy();
                let repl = self.eval_required_arg(frame, args, 2, "Replace")?.to_string_lossy();
                // Optional start position (1-based).
                let start = args
                    .get(3)
                    .map(|a| self.eval_expr(frame, &a.expr))
                    .transpose()?
                    .and_then(|v| v.to_f64())
                    .unwrap_or(1.0) as isize;
                let start = (start - 1).max(0) as usize;
                let chars: Vec<char> = expr.chars().collect();
                let prefix: String = chars.iter().take(start).collect();
                let suffix: String = chars.iter().skip(start).collect();
                Ok(VbaValue::String(format!(
                    "{}{}",
                    prefix,
                    suffix.replace(&find, &repl)
                )))
            }
            // ---- Date/time ----
            _ if name.eq_ignore_ascii_case("now") => {
                Ok(VbaValue::Date(datetime_to_ole_date(Local::now().naive_local())))
            }
            _ if name.eq_ignore_ascii_case("date") => {
                let today = Local::now().date_naive();
                Ok(VbaValue::Date(datetime_to_ole_date(
                    today.and_hms_opt(0, 0, 0).unwrap(),
                )))
            }
            _ if name.eq_ignore_ascii_case("time") => {
                let now = Local::now().naive_local().time();
                let secs = now.num_seconds_from_midnight() as f64
                    + (now.nanosecond() as f64) / 1_000_000_000.0;
                Ok(VbaValue::Date(secs / 86_400.0))
            }
            _ if name.eq_ignore_ascii_case("dateadd") => {
                let interval = self.eval_required_arg(frame, args, 0, "DateAdd")?.to_string_lossy();
                let number = self
                    .eval_required_arg(frame, args, 1, "DateAdd")?
                    .to_f64()
                    .unwrap_or(0.0) as i64;
                let date = self.eval_required_arg(frame, args, 2, "DateAdd")?;
                let serial = self.coerce_to_date_serial(frame, date)?;
                let dt = ole_date_to_datetime(serial);
                let out = date_add(&interval, number, dt)?;
                Ok(VbaValue::Date(datetime_to_ole_date(out)))
            }
            _ if name.eq_ignore_ascii_case("datediff") => {
                let interval = self.eval_required_arg(frame, args, 0, "DateDiff")?.to_string_lossy();
                let d1 = self.eval_required_arg(frame, args, 1, "DateDiff")?;
                let d2 = self.eval_required_arg(frame, args, 2, "DateDiff")?;
                let t1 = ole_date_to_datetime(self.coerce_to_date_serial(frame, d1)?);
                let t2 = ole_date_to_datetime(self.coerce_to_date_serial(frame, d2)?);
                Ok(VbaValue::Double(date_diff(&interval, t1, t2)? as f64))
            }
            _ if name.eq_ignore_ascii_case("format") => {
                let value = self.eval_required_arg(frame, args, 0, "Format")?;
                let fmt = args
                    .get(1)
                    .map(|a| self.eval_expr(frame, &a.expr))
                    .transpose()?
                    .map(|v| v.to_string_lossy());
                Ok(VbaValue::String(format_value(value, fmt.as_deref())))
            }
            _ => {
                // Procedure call.
                let name_lc = name.to_ascii_lowercase();
                if let Some(proc) = self.program.get(&name_lc) {
                    let arg_vals = build_proc_args(frame, args, proc, self)?;
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

    fn eval_required_arg(
        &mut self,
        frame: &mut Frame,
        args: &[crate::ast::CallArg],
        idx: usize,
        name: &str,
    ) -> Result<VbaValue, VbaError> {
        let expr = args
            .get(idx)
            .ok_or_else(|| VbaError::Runtime(format!("{name}() missing argument {idx}")))?;
        self.eval_expr(frame, &expr.expr)
    }

    fn call_object_method(
        &mut self,
        frame: &mut Frame,
        obj: VbaObjectRef,
        member: &str,
        args: &[crate::ast::CallArg],
    ) -> Result<VbaValue, VbaError> {
        let snapshot = obj.borrow().clone();
        match snapshot {
            VbaObject::Worksheet { sheet } => match () {
                _ if member.eq_ignore_ascii_case("range") => {
                    if args.is_empty() {
                        return Err(VbaError::Runtime("Range() missing argument".to_string()));
                    }
                    let range_ref = self.eval_range_args(frame, sheet, args)?;
                    Ok(VbaValue::Object(VbaObjectRef::new(VbaObject::Range(range_ref))))
                }
                _ if member.eq_ignore_ascii_case("cells") => {
                    if args.is_empty() {
                        return Ok(VbaValue::Object(VbaObjectRef::new(VbaObject::Range(
                            sheet_entire_range(self.sheet, sheet),
                        ))));
                    }
                    if args.len() < 2 {
                        return Err(VbaError::Runtime("Cells() missing arguments".to_string()));
                    }
                    let row = self.eval_required_arg(frame, args, 0, "Cells")?;
                    let col = self.eval_required_arg(frame, args, 1, "Cells")?;
                    let row = self.coerce_cells_index(frame, row, true)?;
                    let col = self.coerce_cells_index(frame, col, false)?;
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
                _ if member.eq_ignore_ascii_case("rows") => {
                    if args.is_empty() {
                        return Ok(VbaValue::Object(VbaObjectRef::new(VbaObject::RangeRows {
                            range: sheet_entire_range(self.sheet, sheet),
                        })));
                    }
                    let v = self.eval_required_arg(frame, args, 0, "Rows")?;
                    let range = match v {
                        VbaValue::String(s) => self.sheet_range(sheet, &s)?,
                        other => {
                            let row = self.coerce_cells_index(frame, other, true)?;
                            let mut range = sheet_entire_range(self.sheet, sheet);
                            range.start_row = row;
                            range.end_row = row;
                            range
                        }
                    };
                    Ok(VbaValue::Object(VbaObjectRef::new(VbaObject::Range(range))))
                }
                _ if member.eq_ignore_ascii_case("columns") => {
                    if args.is_empty() {
                        return Ok(VbaValue::Object(VbaObjectRef::new(VbaObject::RangeColumns {
                            range: sheet_entire_range(self.sheet, sheet),
                        })));
                    }
                    let v = self.eval_required_arg(frame, args, 0, "Columns")?;
                    let range = match v {
                        VbaValue::String(s) => self.sheet_range(sheet, &s)?,
                        other => {
                            let col = self.coerce_cells_index(frame, other, false)?;
                            let mut range = sheet_entire_range(self.sheet, sheet);
                            range.start_col = col;
                            range.end_col = col;
                            range
                        }
                    };
                    Ok(VbaValue::Object(VbaObjectRef::new(VbaObject::Range(range))))
                }
                _ if member.eq_ignore_ascii_case("paste")
                    || member.eq_ignore_ascii_case("pastespecial") =>
                {
                    // Best-effort: paste clipboard at the current selection/active cell.
                    //
                    // Excel's recorder frequently emits `ActiveSheet.Paste` after selecting the
                    // destination cell. We treat this like a "paste all" (values+formulas).
                    let dest = if args.is_empty() {
                        if let Some(sel) = self.selection {
                            if sel.sheet == sheet {
                                sel
                            } else if self.sheet.active_sheet() == sheet {
                                let (r, c) = self.sheet.active_cell();
                                VbaRangeRef {
                                    sheet,
                                    start_row: r,
                                    start_col: c,
                                    end_row: r,
                                    end_col: c,
                                }
                            } else {
                                VbaRangeRef {
                                    sheet,
                                    start_row: 1,
                                    start_col: 1,
                                    end_row: 1,
                                    end_col: 1,
                                }
                            }
                        } else if self.sheet.active_sheet() == sheet {
                            let (r, c) = self.sheet.active_cell();
                            VbaRangeRef {
                                sheet,
                                start_row: r,
                                start_col: c,
                                end_row: r,
                                end_col: c,
                            }
                        } else {
                            VbaRangeRef {
                                sheet,
                                start_row: 1,
                                start_col: 1,
                                end_row: 1,
                                end_col: 1,
                            }
                        }
                    } else {
                        let dest_arg = arg_named_or_pos(args, "destination", 0).ok_or_else(|| {
                            VbaError::Runtime("Paste() missing destination".to_string())
                        })?;
                        let dest_val = self.eval_expr(frame, &dest_arg.expr)?;
                        let dest_obj = dest_val.as_object().ok_or_else(|| {
                            VbaError::Runtime("Paste destination must be a Range".to_string())
                        })?;
                        let dest_range = {
                            let borrowed = dest_obj.borrow();
                            match &*borrowed {
                                VbaObject::Range(r) => *r,
                                _ => {
                                    return Err(VbaError::Runtime(
                                        "Paste destination must be a Range".to_string(),
                                    ))
                                }
                            }
                        };
                        dest_range
                    };

                    if let Some(clip) = self.clipboard.clone() {
                        self.paste_clipboard(&clip, dest, true)?;
                    }
                    Ok(VbaValue::Empty)
                }
                _ if member.eq_ignore_ascii_case("activate")
                    || member.eq_ignore_ascii_case("select") =>
                {
                    self.sheet.set_active_sheet(sheet)?;
                    self.selection = None;
                    Ok(VbaValue::Empty)
                }
                _ => Err(VbaError::Runtime(format!(
                    "Unknown Worksheet method `{member}`"
                ))),
            },
            VbaObject::Range(range) => match () {
                _ if member.eq_ignore_ascii_case("select") => {
                    self.sheet.set_active_sheet(range.sheet)?;
                    self.sheet.set_active_cell(range.start_row, range.start_col)?;
                    self.selection = Some(range);
                    Ok(VbaValue::Empty)
                }
                _ if member.eq_ignore_ascii_case("copy") => {
                    if args.is_empty() {
                        self.clipboard = Some(self.snapshot_range(range)?);
                        return Ok(VbaValue::Empty);
                    }
                    let dest_arg = arg_named_or_pos(args, "destination", 0)
                        .ok_or_else(|| VbaError::Runtime("Copy() missing destination".to_string()))?;
                    let dest = self.eval_expr(frame, &dest_arg.expr)?;
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
                    let dest_range = expand_single_cell_destination(dest_range, range);
                    self.copy_range(range, dest_range)?;
                    Ok(VbaValue::Empty)
                }
                _ if member.eq_ignore_ascii_case("pastespecial") => {
                    // Best-effort:
                    // - `PasteSpecial` with no args behaves like "paste everything" (values+formulas).
                    // - `PasteSpecial ...` (typically `Paste:=xlPasteValues`) pastes values only.
                    let include_formulas = if args.is_empty() {
                        true
                    } else if let Some(paste_arg) = arg_named_or_pos(args, "paste", 0) {
                        if matches!(paste_arg.expr, Expr::Missing) {
                            true
                        } else {
                            let paste = self.eval_expr(frame, &paste_arg.expr)?;
                            let code = paste.to_f64().unwrap_or(0.0) as i64;
                            match code {
                                // xlPasteAll
                                -4104 => true,
                                // xlPasteFormulas
                                -4123 => true,
                                // xlPasteValues (and everything else we don't understand)
                                _ => false,
                            }
                        }
                    } else {
                        // If Paste is omitted (but other args like Operation/SkipBlanks are
                        // supplied), Excel defaults to xlPasteAll.
                        true
                    };
                    if let Some(clip) = self.clipboard.clone() {
                        self.paste_clipboard(&clip, range, include_formulas)?;
                    }
                    Ok(VbaValue::Empty)
                }
                _ if member.eq_ignore_ascii_case("autofill") => {
                    let dest_arg = arg_named_or_pos(args, "destination", 0)
                        .ok_or_else(|| {
                            VbaError::Runtime("AutoFill() missing destination".to_string())
                        })?;
                    let dest = self.eval_expr(frame, &dest_arg.expr)?;
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
                _ if member.eq_ignore_ascii_case("offset") => {
                    let row_off = args
                        .first()
                        .map(|a| self.eval_expr(frame, &a.expr))
                        .transpose()?
                        .unwrap_or(VbaValue::Double(0.0))
                        .to_f64()
                        .unwrap_or(0.0) as i32;
                    let col_off = args
                        .get(1)
                        .map(|a| self.eval_expr(frame, &a.expr))
                        .transpose()?
                        .unwrap_or(VbaValue::Double(0.0))
                        .to_f64()
                        .unwrap_or(0.0) as i32;
                    let new = offset_range(range, row_off, col_off)?;
                    Ok(VbaValue::Object(VbaObjectRef::new(VbaObject::Range(
                        new,
                    ))))
                }
                _ if member.eq_ignore_ascii_case("resize") => {
                    let rows = match args.first() {
                        None => None,
                        Some(arg) if matches!(arg.expr, Expr::Missing) => None,
                        Some(arg) => self
                            .eval_expr(frame, &arg.expr)?
                            .to_f64()
                            .map(|v| v as i32),
                    };
                    let cols = match args.get(1) {
                        None => None,
                        Some(arg) if matches!(arg.expr, Expr::Missing) => None,
                        Some(arg) => self
                            .eval_expr(frame, &arg.expr)?
                            .to_f64()
                            .map(|v| v as i32),
                    };
                    let new = resize_range(range, rows, cols)?;
                    Ok(VbaValue::Object(VbaObjectRef::new(VbaObject::Range(
                        new,
                    ))))
                }
                _ if member.eq_ignore_ascii_case("end") => {
                    let dir = self.eval_required_arg(frame, args, 0, "End")?;
                    let dir = dir.to_f64().unwrap_or(0.0) as i64;
                    let end = self.range_end(range, dir)?;
                    Ok(VbaValue::Object(VbaObjectRef::new(VbaObject::Range(end))))
                }
                _ if member.eq_ignore_ascii_case("clearcontents") => {
                    if let Some(cells) = self.sheet.used_cells_in_range(range) {
                        for (r, c) in cells {
                            self.tick()?;
                            self.sheet.clear_cell_contents(range.sheet, r, c)?;
                        }
                    } else {
                        for r in range.start_row..=range.end_row {
                            for c in range.start_col..=range.end_col {
                                self.tick()?;
                                self.sheet.clear_cell_contents(range.sheet, r, c)?;
                            }
                        }
                    }
                    Ok(VbaValue::Empty)
                }
                _ if member.eq_ignore_ascii_case("clear") => {
                    // `Clear` clears contents + formatting. We only model contents for now.
                    self.call_object_method(frame, obj, "ClearContents", &[])?;
                    Ok(VbaValue::Empty)
                }
                _ => Err(VbaError::Runtime(format!(
                    "Unknown Range method `{member}`"
                ))),
            },
            VbaObject::Application => match () {
                _ if member.eq_ignore_ascii_case("range") => {
                    if args.is_empty() {
                        return Err(VbaError::Runtime("Range() missing argument".to_string()));
                    }
                    let sheet = self.sheet.active_sheet();
                    let range_ref = self.eval_range_args(frame, sheet, args)?;
                    Ok(VbaValue::Object(VbaObjectRef::new(VbaObject::Range(range_ref))))
                }
                _ if member.eq_ignore_ascii_case("worksheets")
                    || member.eq_ignore_ascii_case("sheets") =>
                {
                    let arg = self.eval_required_arg(frame, args, 0, "Worksheets")?;
                    let idx = match arg {
                        VbaValue::String(name) => self
                            .sheet
                            .sheet_index(&name)
                            .ok_or_else(|| VbaError::Runtime(format!("Unknown worksheet `{name}`")))?,
                        other => {
                            let n = other.to_f64().unwrap_or(0.0) as isize;
                            if n <= 0 {
                                return Err(VbaError::Runtime(format!(
                                    "Worksheet index must be >= 1 (got {n})"
                                )));
                            }
                            let idx = (n - 1) as usize;
                            if idx >= self.sheet.sheet_count() {
                                return Err(VbaError::Runtime(format!(
                                    "Worksheet index out of range: {n}"
                                )));
                            }
                            idx
                        }
                    };
                    Ok(VbaValue::Object(VbaObjectRef::new(VbaObject::Worksheet {
                        sheet: idx,
                    })))
                }
                _ => Err(VbaError::Runtime(format!(
                    "Unknown Application method `{member}`"
                ))),
            },
            VbaObject::Workbook => match () {
                _ if member.eq_ignore_ascii_case("worksheets")
                    || member.eq_ignore_ascii_case("sheets") =>
                {
                    let arg = self.eval_required_arg(frame, args, 0, "Worksheets")?;
                    let idx = match arg {
                        VbaValue::String(name) => self
                            .sheet
                            .sheet_index(&name)
                            .ok_or_else(|| VbaError::Runtime(format!("Unknown worksheet `{name}`")))?,
                        other => {
                            let n = other.to_f64().unwrap_or(0.0) as isize;
                            if n <= 0 {
                                return Err(VbaError::Runtime(format!(
                                    "Worksheet index must be >= 1 (got {n})"
                                )));
                            }
                            let idx = (n - 1) as usize;
                            if idx >= self.sheet.sheet_count() {
                                return Err(VbaError::Runtime(format!(
                                    "Worksheet index out of range: {n}"
                                )));
                            }
                            idx
                        }
                    };
                    Ok(VbaValue::Object(VbaObjectRef::new(VbaObject::Worksheet {
                        sheet: idx,
                    })))
                }
                _ => Err(VbaError::Runtime(format!(
                    "Unknown Workbook method `{member}`"
                ))),
            },
            VbaObject::Collection { .. } => match () {
                _ if member.eq_ignore_ascii_case("add") => {
                    let item = self.eval_required_arg(frame, args, 0, "Add")?;
                    if let VbaObject::Collection { items } = &mut *obj.borrow_mut() {
                        items.push(item);
                    }
                    Ok(VbaValue::Empty)
                }
                _ if member.eq_ignore_ascii_case("item") => {
                    let index = self
                        .eval_required_arg(frame, args, 0, "Item")?
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
            VbaObject::Dictionary { .. } => match () {
                _ if member.eq_ignore_ascii_case("add") => {
                    let key = self.eval_required_arg(frame, args, 0, "Add")?.to_string_lossy();
                    let item = self.eval_required_arg(frame, args, 1, "Add")?;
                    if let VbaObject::Dictionary { items } = &mut *obj.borrow_mut() {
                        items.insert(key, item);
                    }
                    Ok(VbaValue::Empty)
                }
                _ if member.eq_ignore_ascii_case("exists") => {
                    let key = self
                        .eval_required_arg(frame, args, 0, "Exists")?
                        .to_string_lossy();
                    if let VbaObject::Dictionary { items } = &*obj.borrow() {
                        Ok(VbaValue::Boolean(items.contains_key(&key)))
                    } else {
                        Ok(VbaValue::Boolean(false))
                    }
                }
                _ if member.eq_ignore_ascii_case("item") => {
                    let key = self.eval_required_arg(frame, args, 0, "Item")?.to_string_lossy();
                    if let VbaObject::Dictionary { items } = &*obj.borrow() {
                        Ok(items.get(&key).cloned().unwrap_or(VbaValue::Empty))
                    } else {
                        Ok(VbaValue::Empty)
                    }
                }
                _ if member.eq_ignore_ascii_case("keys") => {
                    if let VbaObject::Dictionary { items } = &*obj.borrow() {
                        Ok(VbaValue::Array(std::rc::Rc::new(RefCell::new(
                            VbaArray::new(
                                0,
                                items
                                    .keys()
                                    .cloned()
                                    .map(VbaValue::String)
                                    .collect::<Vec<_>>(),
                            ),
                        ))))
                    } else {
                        Ok(VbaValue::Array(std::rc::Rc::new(RefCell::new(
                            VbaArray::new(0, Vec::new()),
                        ))))
                    }
                }
                _ if member.eq_ignore_ascii_case("remove") => {
                    let key = self
                        .eval_required_arg(frame, args, 0, "Remove")?
                        .to_string_lossy();
                    if let VbaObject::Dictionary { items } = &mut *obj.borrow_mut() {
                        items.remove(&key);
                    }
                    Ok(VbaValue::Empty)
                }
                _ if member.eq_ignore_ascii_case("removeall") => {
                    if let VbaObject::Dictionary { items } = &mut *obj.borrow_mut() {
                        items.clear();
                    }
                    Ok(VbaValue::Empty)
                }
                _ => Err(VbaError::Runtime(format!(
                    "Unknown Dictionary method `{member}`"
                ))),
            },
            VbaObject::Err(_) => match () {
                _ if member.eq_ignore_ascii_case("clear") => {
                    self.clear_err();
                    Ok(VbaValue::Empty)
                }
                _ => Err(VbaError::Runtime(format!("Unknown Err method `{member}`"))),
            },
            VbaObject::RangeRows { .. } | VbaObject::RangeColumns { .. } => Err(VbaError::Runtime(
                format!("Cannot call `{member}` on range dimension object"),
            )),
        }
    }

    fn sheet_range(&self, sheet: usize, a1: &str) -> Result<VbaRangeRef, VbaError> {
        let (max_row, max_col) = self.sheet.sheet_dimensions(sheet);
        let (r1, c1, r2, c2) =
            crate::object_model::parse_range_a1_with_bounds(a1, max_row, max_col)?;
        Ok(VbaRangeRef {
            sheet,
            start_row: r1,
            start_col: c1,
            end_row: r2,
            end_col: c2,
        })
    }

    fn eval_range_args(
        &mut self,
        frame: &mut Frame,
        sheet: usize,
        args: &[crate::ast::CallArg],
    ) -> Result<VbaRangeRef, VbaError> {
        if args.is_empty() {
            return Err(VbaError::Runtime("Range() missing argument".to_string()));
        }
        if args.len() > 2 {
            return Err(VbaError::Runtime(format!(
                "Range() expects 1 or 2 arguments (got {})",
                args.len()
            )));
        }

        let cell1 = arg_named_or_pos(args, "cell1", 0).unwrap_or(&args[0]);
        let r1 = self.eval_range_arg_value(frame, sheet, &cell1.expr)?;

        let cell2 = arg_named_or_pos(args, "cell2", 1);
        if let Some(cell2) = cell2 {
            let r2 = self.eval_range_arg_value(frame, sheet, &cell2.expr)?;
            return Ok(VbaRangeRef {
                sheet,
                start_row: r1.start_row.min(r2.start_row),
                start_col: r1.start_col.min(r2.start_col),
                end_row: r1.end_row.max(r2.end_row),
                end_col: r1.end_col.max(r2.end_col),
            });
        }

        Ok(r1)
    }

    fn eval_range_arg_value(
        &mut self,
        frame: &mut Frame,
        sheet: usize,
        expr: &Expr,
    ) -> Result<VbaRangeRef, VbaError> {
        let value = self.eval_expr(frame, expr)?;
        match value {
            VbaValue::String(a1) => self.sheet_range(sheet, &a1),
            VbaValue::Object(obj) => match &*obj.borrow() {
                VbaObject::Range(range) => {
                    if range.sheet != sheet {
                        return Err(VbaError::Runtime(
                            "Range arguments must be on the same sheet".to_string(),
                        ));
                    }
                    Ok(*range)
                }
                _ => Err(VbaError::Runtime(
                    "Range() arguments must be strings or Range objects".to_string(),
                )),
            },
            _ => Err(VbaError::Runtime(
                "Range() arguments must be strings or Range objects".to_string(),
            )),
        }
    }

    fn get_object_member(&mut self, obj: VbaObjectRef, member: &str) -> Result<VbaValue, VbaError> {
        match &*obj.borrow() {
            VbaObject::Application => match () {
                _ if member.eq_ignore_ascii_case("activesheet") => {
                    Ok(VbaValue::Object(VbaObjectRef::new(VbaObject::Worksheet {
                    sheet: self.sheet.active_sheet(),
                    })))
                }
                _ if member.eq_ignore_ascii_case("activecell") => {
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
                _ if member.eq_ignore_ascii_case("selection") => {
                    let sel = self.selection.unwrap_or_else(|| {
                        let (r, c) = self.sheet.active_cell();
                        VbaRangeRef {
                            sheet: self.sheet.active_sheet(),
                            start_row: r,
                            start_col: c,
                            end_row: r,
                            end_col: c,
                        }
                    });
                    Ok(VbaValue::Object(VbaObjectRef::new(VbaObject::Range(sel))))
                }
                _ if member.eq_ignore_ascii_case("cutcopymode") => {
                    Ok(VbaValue::Boolean(self.clipboard.is_some()))
                }
                _ if member.eq_ignore_ascii_case("activeworkbook") => {
                    Ok(VbaValue::Object(VbaObjectRef::new(VbaObject::Workbook)))
                }
                _ => Err(VbaError::Runtime(format!(
                    "Unknown Application member `{member}`"
                ))),
            },
            VbaObject::Workbook => match () {
                _ if member.eq_ignore_ascii_case("activesheet") => {
                    Ok(VbaValue::Object(VbaObjectRef::new(VbaObject::Worksheet {
                    sheet: self.sheet.active_sheet(),
                    })))
                }
                _ => Err(VbaError::Runtime(format!(
                    "Unknown Workbook member `{member}`"
                ))),
            },
            VbaObject::Worksheet { sheet } => match () {
                _ if member.eq_ignore_ascii_case("name") => Ok(VbaValue::String(
                    self.sheet.sheet_name(*sheet).unwrap_or("").to_string(),
                )),
                _ if member.eq_ignore_ascii_case("cells") => Ok(VbaValue::Object(VbaObjectRef::new(VbaObject::Range(
                    sheet_entire_range(self.sheet, *sheet),
                )))),
                _ if member.eq_ignore_ascii_case("rows") => Ok(VbaValue::Object(VbaObjectRef::new(VbaObject::RangeRows {
                    range: sheet_entire_range(self.sheet, *sheet),
                }))),
                _ if member.eq_ignore_ascii_case("columns") => Ok(VbaValue::Object(VbaObjectRef::new(VbaObject::RangeColumns {
                    range: sheet_entire_range(self.sheet, *sheet),
                }))),
                _ => Err(VbaError::Runtime(format!(
                    "Unknown Worksheet member `{member}`"
                ))),
            },
            VbaObject::Range(range) => match () {
                _ if member.eq_ignore_ascii_case("value") || member.eq_ignore_ascii_case("value2") => {
                    self.get_range_value(*range)
                }
                _ if member.eq_ignore_ascii_case("formula") || member.eq_ignore_ascii_case("formular1c1") => {
                    self.get_range_formula(*range)
                }
                _ if member.eq_ignore_ascii_case("text") => Ok(self
                    .sheet
                    .get_cell_value(range.sheet, range.start_row, range.start_col)?
                    .to_string_lossy()
                    .into()),
                _ if member.eq_ignore_ascii_case("address") => Ok(VbaValue::String(range_address(*range)?)),
                _ if member.eq_ignore_ascii_case("row") => Ok(VbaValue::Double(range.start_row as f64)),
                _ if member.eq_ignore_ascii_case("column") => Ok(VbaValue::Double(range.start_col as f64)),
                _ if member.eq_ignore_ascii_case("rows") => Ok(VbaValue::Object(VbaObjectRef::new(VbaObject::RangeRows {
                    range: *range,
                }))),
                _ if member.eq_ignore_ascii_case("columns") => Ok(VbaValue::Object(VbaObjectRef::new(VbaObject::RangeColumns {
                    range: *range,
                }))),
                _ => Err(VbaError::Runtime(format!("Unknown Range member `{member}`"))),
            },
            VbaObject::RangeRows { range } => match () {
                _ if member.eq_ignore_ascii_case("count") => Ok(VbaValue::Double(
                    (range.end_row - range.start_row + 1) as f64,
                )),
                _ => Err(VbaError::Runtime(format!(
                    "Unknown Rows member `{member}`"
                ))),
            },
            VbaObject::RangeColumns { range } => match () {
                _ if member.eq_ignore_ascii_case("count") => Ok(VbaValue::Double(
                    (range.end_col - range.start_col + 1) as f64,
                )),
                _ => Err(VbaError::Runtime(format!(
                    "Unknown Columns member `{member}`"
                ))),
            },
            VbaObject::Collection { items } => match () {
                _ if member.eq_ignore_ascii_case("count") => Ok(VbaValue::Double(items.len() as f64)),
                _ => Err(VbaError::Runtime(format!(
                    "Unknown Collection member `{member}`"
                ))),
            },
            VbaObject::Dictionary { items } => match () {
                _ if member.eq_ignore_ascii_case("count") => Ok(VbaValue::Double(items.len() as f64)),
                _ => Err(VbaError::Runtime(format!(
                    "Unknown Dictionary member `{member}`"
                ))),
            },
            VbaObject::Err(err) => match () {
                _ if member.eq_ignore_ascii_case("number") => Ok(VbaValue::Double(err.number as f64)),
                _ if member.eq_ignore_ascii_case("description") => Ok(VbaValue::String(err.description.clone())),
                _ => Err(VbaError::Runtime(format!("Unknown Err member `{member}`"))),
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

    fn range_end(&mut self, range: VbaRangeRef, dir: i64) -> Result<VbaRangeRef, VbaError> {
        // Excel direction constants (subset).
        const XL_DOWN: i64 = -4121;
        const XL_UP: i64 = -4162;
        const XL_TO_LEFT: i64 = -4159;
        const XL_TO_RIGHT: i64 = -4161;

        let (drow, dcol): (i64, i64) = match dir {
            XL_DOWN => (1, 0),
            XL_UP => (-1, 0),
            XL_TO_LEFT => (0, -1),
            XL_TO_RIGHT => (0, 1),
            _ => {
                return Err(VbaError::Runtime(format!(
                    "Range.End unsupported direction: {dir}"
                )))
            }
        };

        let sheet = range.sheet;
        let (max_row, max_col) = self.sheet.sheet_dimensions(sheet);
        let max_row = max_row as i64;
        let max_col = max_col as i64;
        let mut row = range.start_row;
        let mut col = range.start_col;

        let next_cell = |row: u32, col: u32| -> Option<(u32, u32)> {
            let nr = row as i64 + drow;
            let nc = col as i64 + dcol;
            if nr <= 0 || nc <= 0 {
                return None;
            }
            if nr > max_row || nc > max_col {
                return None;
            }
            Some((nr as u32, nc as u32))
        };

        let start_has_content = self.cell_has_content(sheet, row, col)?;

        if !start_has_content {
            // Excel semantics: from an empty cell, move to the *first* non-empty cell in that
            // direction (do not continue through a contiguous block).
            //
            // This is a critical compatibility point for patterns like:
            //   Cells(Rows.Count, 1).End(xlUp).Row
            // which would otherwise scan the entire worksheet and/or land on the wrong edge of
            // the data block.
            let fast = match dir {
                XL_UP => self
                    .sheet
                    .last_used_row_in_column(sheet, col, row)
                    .map(|r| (r, col)),
                XL_DOWN => self
                    .sheet
                    .next_used_row_in_column(sheet, col, row)
                    .map(|r| (r, col)),
                XL_TO_LEFT => self
                    .sheet
                    .last_used_col_in_row(sheet, row, col)
                    .map(|c| (row, c)),
                XL_TO_RIGHT => self
                    .sheet
                    .next_used_col_in_row(sheet, row, col)
                    .map(|c| (row, c)),
                _ => None,
            };

            if let Some((r, c)) = fast {
                if self.cell_has_content(sheet, r, c)? {
                    return Ok(VbaRangeRef {
                        sheet,
                        start_row: r,
                        start_col: c,
                        end_row: r,
                        end_col: c,
                    });
                }
            }

            // Fallback: scan forward until we find content (best-effort).
            let mut cursor = (row, col);
            loop {
                let Some(next) = next_cell(cursor.0, cursor.1) else {
                    break;
                };
                cursor = next;
                if self.cell_has_content(sheet, cursor.0, cursor.1)? {
                    row = cursor.0;
                    col = cursor.1;
                    break;
                }
            }
        } else {
            // From a non-empty cell, move until the next cell is empty (end of the contiguous
            // block).
            loop {
                let Some((nr, nc)) = next_cell(row, col) else {
                    break;
                };
                if !self.cell_has_content(sheet, nr, nc)? {
                    break;
                }
                row = nr;
                col = nc;
            }
        }

        Ok(VbaRangeRef {
            sheet,
            start_row: row,
            start_col: col,
            end_row: row,
            end_col: col,
        })
    }

    fn cell_has_content(&mut self, sheet: usize, row: u32, col: u32) -> Result<bool, VbaError> {
        self.tick()?;
        let value = self.sheet.get_cell_value(sheet, row, col)?;
        if !matches!(value, VbaValue::Empty) {
            return Ok(true);
        }
        Ok(self.sheet.get_cell_formula(sheet, row, col)?.is_some())
    }

    fn snapshot_range(&mut self, range: VbaRangeRef) -> Result<Clipboard, VbaError> {
        let rows = range.end_row - range.start_row + 1;
        let cols = range.end_col - range.start_col + 1;
        let mut values = Vec::new();
        let mut formulas = Vec::new();
        for r in 0..rows {
            for c in 0..cols {
                self.tick()?;
                let row = range.start_row + r;
                let col = range.start_col + c;
                values.push(self.sheet.get_cell_value(range.sheet, row, col)?);
                formulas.push(self.sheet.get_cell_formula(range.sheet, row, col)?);
            }
        }
        Ok(Clipboard {
            rows,
            cols,
            values,
            formulas,
        })
    }

    fn paste_clipboard(
        &mut self,
        clip: &Clipboard,
        dest: VbaRangeRef,
        include_formulas: bool,
    ) -> Result<(), VbaError> {
        let mut dest = dest;
        if dest.start_row == dest.end_row
            && dest.start_col == dest.end_col
            && (clip.rows > 1 || clip.cols > 1)
        {
            dest.end_row = dest.start_row + clip.rows - 1;
            dest.end_col = dest.start_col + clip.cols - 1;
        }
        let dest_rows = dest.end_row - dest.start_row + 1;
        let dest_cols = dest.end_col - dest.start_col + 1;
        for dr in 0..dest_rows {
            for dc in 0..dest_cols {
                let sr = (dr % clip.rows) as usize;
                let sc = (dc % clip.cols) as usize;
                let idx = sr * (clip.cols as usize) + sc;
                let value = clip.values.get(idx).cloned().unwrap_or(VbaValue::Empty);
                let formula = clip.formulas.get(idx).cloned().unwrap_or(None);
                let tr = dest.start_row + dr;
                let tc = dest.start_col + dc;
                self.tick()?;
                self.sheet.clear_cell_contents(dest.sheet, tr, tc)?;
                if include_formulas {
                    if let Some(formula) = formula {
                        self.sheet.set_cell_formula(dest.sheet, tr, tc, formula)?;
                    } else {
                        self.sheet.set_cell_value(dest.sheet, tr, tc, value)?;
                    }
                } else {
                    self.sheet.set_cell_value(dest.sheet, tr, tc, value)?;
                }
            }
        }
        Ok(())
    }

    fn copy_range(&mut self, src: VbaRangeRef, dest: VbaRangeRef) -> Result<(), VbaError> {
        let src_rows = src.end_row.saturating_sub(src.start_row) + 1;
        let src_cols = src.end_col.saturating_sub(src.start_col) + 1;
        let dest_rows = dest.end_row.saturating_sub(dest.start_row) + 1;
        let dest_cols = dest.end_col.saturating_sub(dest.start_col) + 1;

        for dr in 0..dest_rows {
            for dc in 0..dest_cols {
                self.tick()?;
                let sr = src.start_row + (dr % src_rows);
                let sc = src.start_col + (dc % src_cols);
                let value = self.sheet.get_cell_value(src.sheet, sr, sc)?;
                let formula = self.sheet.get_cell_formula(src.sheet, sr, sc)?;

                let tr = dest.start_row + dr;
                let tc = dest.start_col + dc;
                self.sheet.clear_cell_contents(dest.sheet, tr, tc)?;
                if let Some(formula) = formula {
                    self.sheet.set_cell_formula(dest.sheet, tr, tc, formula)?;
                } else {
                    self.sheet.set_cell_value(dest.sheet, tr, tc, value)?;
                }
            }
        }

        Ok(())
    }

    fn index_array_value(
        &mut self,
        frame: &mut Frame,
        mut value: VbaValue,
        args: &[crate::ast::CallArg],
    ) -> Result<VbaValue, VbaError> {
        if args.is_empty() {
            return Err(VbaError::Runtime("Array index missing".to_string()));
        }

        for arg in args {
            let idx = self
                .eval_expr(frame, &arg.expr)?
                .to_f64()
                .unwrap_or(0.0) as i32;
            let VbaValue::Array(arr) = value else {
                return Ok(VbaValue::Empty);
            };
            value = arr
                .borrow()
                .get(idx)
                .cloned()
                .unwrap_or(VbaValue::Empty);
        }
        Ok(value)
    }

    fn get_range_value(&mut self, range: VbaRangeRef) -> Result<VbaValue, VbaError> {
        if range.start_row == range.end_row && range.start_col == range.end_col {
            return self
                .sheet
                .get_cell_value(range.sheet, range.start_row, range.start_col);
        }

        let rows = range.end_row.saturating_sub(range.start_row) + 1;
        let cols = range.end_col.saturating_sub(range.start_col) + 1;

        let mut outer = Vec::with_capacity(rows as usize);
        for r_off in 0..rows {
            let mut inner = Vec::with_capacity(cols as usize);
            for c_off in 0..cols {
                let value = self.sheet.get_cell_value(
                    range.sheet,
                    range.start_row + r_off,
                    range.start_col + c_off,
                )?;
                inner.push(value);
            }
            outer.push(VbaValue::Array(std::rc::Rc::new(RefCell::new(VbaArray::new(
                1, inner,
            )))));
        }
        Ok(VbaValue::Array(std::rc::Rc::new(RefCell::new(VbaArray::new(
            1, outer,
        )))))
    }

    fn get_range_formula(&mut self, range: VbaRangeRef) -> Result<VbaValue, VbaError> {
        if range.start_row == range.end_row && range.start_col == range.end_col {
            return Ok(self
                .sheet
                .get_cell_formula(range.sheet, range.start_row, range.start_col)?
                .map(VbaValue::String)
                .unwrap_or(VbaValue::Empty));
        }

        let rows = range.end_row.saturating_sub(range.start_row) + 1;
        let cols = range.end_col.saturating_sub(range.start_col) + 1;

        let mut outer = Vec::with_capacity(rows as usize);
        for r_off in 0..rows {
            let mut inner = Vec::with_capacity(cols as usize);
            for c_off in 0..cols {
                let value = self
                    .sheet
                    .get_cell_formula(
                        range.sheet,
                        range.start_row + r_off,
                        range.start_col + c_off,
                    )?
                    .map(VbaValue::String)
                    .unwrap_or(VbaValue::Empty);
                inner.push(value);
            }
            outer.push(VbaValue::Array(std::rc::Rc::new(RefCell::new(VbaArray::new(
                1, inner,
            )))));
        }
        Ok(VbaValue::Array(std::rc::Rc::new(RefCell::new(VbaArray::new(
            1, outer,
        )))))
    }

    fn set_range_value(&mut self, range: VbaRangeRef, value: VbaValue) -> Result<(), VbaError> {
        let rows = range.end_row.saturating_sub(range.start_row) + 1;
        let cols = range.end_col.saturating_sub(range.start_col) + 1;
        if rows == 1 && cols == 1 {
            self.tick()?;
            return self
                .sheet
                .set_cell_value(range.sheet, range.start_row, range.start_col, value);
        }

        match value {
            VbaValue::Array(arr) => self.set_range_value_from_array(range, arr),
            scalar => {
                for r in range.start_row..=range.end_row {
                    for c in range.start_col..=range.end_col {
                        self.tick()?;
                        self.sheet
                            .set_cell_value(range.sheet, r, c, scalar.clone())?;
                    }
                }
                Ok(())
            }
        }
    }

    fn set_range_formula(&mut self, range: VbaRangeRef, value: VbaValue) -> Result<(), VbaError> {
        let rows = range.end_row.saturating_sub(range.start_row) + 1;
        let cols = range.end_col.saturating_sub(range.start_col) + 1;
        if rows == 1 && cols == 1 {
            self.tick()?;
            return self.sheet.set_cell_formula(
                range.sheet,
                range.start_row,
                range.start_col,
                value.to_string_lossy(),
            );
        }

        match value {
            VbaValue::Array(arr) => self.set_range_formula_from_array(range, arr),
            scalar => {
                let formula = scalar.to_string_lossy();
                for r in range.start_row..=range.end_row {
                    for c in range.start_col..=range.end_col {
                        self.tick()?;
                        self.sheet
                            .set_cell_formula(range.sheet, r, c, formula.clone())?;
                    }
                }
                Ok(())
            }
        }
    }

    fn set_range_value_from_array(
        &mut self,
        range: VbaRangeRef,
        arr: VbaArrayRef,
    ) -> Result<(), VbaError> {
        let rows = range.end_row.saturating_sub(range.start_row) + 1;
        let cols = range.end_col.saturating_sub(range.start_col) + 1;
        let rows_usize = rows as usize;
        let cols_usize = cols as usize;

        // 2D array-of-arrays (rows x cols).
        {
            let outer = arr.borrow();
            if outer.values.len() == rows_usize
                && outer.values.iter().all(|v| matches!(v, VbaValue::Array(_)))
            {
                for (r_idx, row_val) in outer.values.iter().enumerate() {
                    let VbaValue::Array(inner) = row_val else {
                        continue;
                    };
                    let inner = inner.borrow();
                    if inner.values.len() != cols_usize {
                        return Err(VbaError::Runtime(format!(
                            "Array size mismatch: expected {cols_usize} values in row, got {}",
                            inner.values.len()
                        )));
                    }

                    for (c_idx, cell_val) in inner.values.iter().enumerate() {
                        self.tick()?;
                        self.sheet.set_cell_value(
                            range.sheet,
                            range.start_row + r_idx as u32,
                            range.start_col + c_idx as u32,
                            cell_val.clone(),
                        )?;
                    }
                }
                return Ok(());
            }
        }

        // Flat array.
        let total = rows_usize.saturating_mul(cols_usize);
        let expected = if rows_usize == 1 {
            cols_usize
        } else if cols_usize == 1 {
            rows_usize
        } else {
            total
        };
        let outer = arr.borrow();
        if outer.values.len() != expected {
            return Err(VbaError::Runtime(format!(
                "Array size mismatch: expected {expected} values for range, got {}",
                outer.values.len()
            )));
        }
        let flat = outer.values.as_slice();

        if rows_usize == 1 && cols_usize >= 1 {
            for (c_idx, cell_val) in flat.iter().enumerate().take(cols_usize) {
                self.tick()?;
                self.sheet.set_cell_value(
                    range.sheet,
                    range.start_row,
                    range.start_col + c_idx as u32,
                    cell_val.clone(),
                )?;
            }
            return Ok(());
        }

        if cols_usize == 1 && rows_usize >= 1 {
            for (r_idx, cell_val) in flat.iter().enumerate().take(rows_usize) {
                self.tick()?;
                self.sheet.set_cell_value(
                    range.sheet,
                    range.start_row + r_idx as u32,
                    range.start_col,
                    cell_val.clone(),
                )?;
            }
            return Ok(());
        }

        for r_idx in 0..rows_usize {
            for c_idx in 0..cols_usize {
                let idx = r_idx * cols_usize + c_idx;
                self.tick()?;
                self.sheet.set_cell_value(
                    range.sheet,
                    range.start_row + r_idx as u32,
                    range.start_col + c_idx as u32,
                    flat[idx].clone(),
                )?;
            }
        }
        Ok(())
    }

    fn set_range_formula_from_array(
        &mut self,
        range: VbaRangeRef,
        arr: VbaArrayRef,
    ) -> Result<(), VbaError> {
        let rows = range.end_row.saturating_sub(range.start_row) + 1;
        let cols = range.end_col.saturating_sub(range.start_col) + 1;
        let rows_usize = rows as usize;
        let cols_usize = cols as usize;

        // 2D array-of-arrays (rows x cols).
        {
            let outer = arr.borrow();
            if outer.values.len() == rows_usize
                && outer.values.iter().all(|v| matches!(v, VbaValue::Array(_)))
            {
                for (r_idx, row_val) in outer.values.iter().enumerate() {
                    let VbaValue::Array(inner) = row_val else {
                        continue;
                    };
                    let inner = inner.borrow();
                    if inner.values.len() != cols_usize {
                        return Err(VbaError::Runtime(format!(
                            "Array size mismatch: expected {cols_usize} formulas in row, got {}",
                            inner.values.len()
                        )));
                    }

                    for (c_idx, cell_val) in inner.values.iter().enumerate() {
                        self.tick()?;
                        self.sheet.set_cell_formula(
                            range.sheet,
                            range.start_row + r_idx as u32,
                            range.start_col + c_idx as u32,
                            cell_val.to_string_lossy(),
                        )?;
                    }
                }
                return Ok(());
            }
        }

        // Flat array.
        let total = rows_usize.saturating_mul(cols_usize);
        let expected = if rows_usize == 1 {
            cols_usize
        } else if cols_usize == 1 {
            rows_usize
        } else {
            total
        };
        let outer = arr.borrow();
        if outer.values.len() != expected {
            return Err(VbaError::Runtime(format!(
                "Array size mismatch: expected {expected} formulas for range, got {}",
                outer.values.len()
            )));
        }
        let flat = outer.values.as_slice();

        if rows_usize == 1 && cols_usize >= 1 {
            for (c_idx, cell_val) in flat.iter().enumerate().take(cols_usize) {
                self.tick()?;
                self.sheet.set_cell_formula(
                    range.sheet,
                    range.start_row,
                    range.start_col + c_idx as u32,
                    cell_val.to_string_lossy(),
                )?;
            }
            return Ok(());
        }

        if cols_usize == 1 && rows_usize >= 1 {
            for (r_idx, cell_val) in flat.iter().enumerate().take(rows_usize) {
                self.tick()?;
                self.sheet.set_cell_formula(
                    range.sheet,
                    range.start_row + r_idx as u32,
                    range.start_col,
                    cell_val.to_string_lossy(),
                )?;
            }
            return Ok(());
        }

        for r_idx in 0..rows_usize {
            for c_idx in 0..cols_usize {
                let idx = r_idx * cols_usize + c_idx;
                self.tick()?;
                self.sheet.set_cell_formula(
                    range.sheet,
                    range.start_row + r_idx as u32,
                    range.start_col + c_idx as u32,
                    flat[idx].to_string_lossy(),
                )?;
            }
        }
        Ok(())
    }
}

impl Frame {
    fn dummy() -> Self {
        Self {
            locals: HashMap::new(),
            types: HashMap::new(),
            consts: HashSet::new(),
            error_mode: ErrorMode::Default,
            resume: ResumeState::default(),
        }
    }
}

#[derive(Debug)]
enum ControlFlow {
    Continue,
    ExitSub,
    ExitFunction,
    ExitFor,
    ExitDo,
    Goto(String),
    ErrorGoto(String),
    Resume(ResumeKind),
}

#[derive(Debug, Clone)]
enum ResumeKind {
    Next,
    Same,
    Label(String),
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

fn vba_round_bankers(n: f64) -> i64 {
    let frac = n.fract().abs();
    if (frac - 0.5).abs() < f64::EPSILON {
        let int = n.trunc() as i64;
        if int % 2 == 0 {
            int
        } else if n.is_sign_negative() {
            int - 1
        } else {
            int + 1
        }
    } else {
        n.round() as i64
    }
}

fn arg_named_or_pos<'a>(
    args: &'a [crate::ast::CallArg],
    name: &str,
    pos: usize,
) -> Option<&'a crate::ast::CallArg> {
    args.iter()
        .find(|a| a.name.as_deref().is_some_and(|n| n.eq_ignore_ascii_case(name)))
        .or_else(|| args.get(pos).filter(|a| a.name.is_none()))
}

fn expand_single_cell_destination(dest: VbaRangeRef, template: VbaRangeRef) -> VbaRangeRef {
    if dest.start_row != dest.end_row || dest.start_col != dest.end_col {
        return dest;
    }
    let rows = template.end_row.saturating_sub(template.start_row) + 1;
    let cols = template.end_col.saturating_sub(template.start_col) + 1;
    VbaRangeRef {
        sheet: dest.sheet,
        start_row: dest.start_row,
        start_col: dest.start_col,
        end_row: dest.start_row + rows - 1,
        end_col: dest.start_col + cols - 1,
    }
}

fn sheet_entire_range(spreadsheet: &dyn Spreadsheet, sheet: usize) -> VbaRangeRef {
    let (max_rows, max_cols) = spreadsheet.sheet_dimensions(sheet);
    VbaRangeRef {
        sheet,
        start_row: 1,
        start_col: 1,
        end_row: max_rows,
        end_col: max_cols,
    }
}

fn offset_range(range: VbaRangeRef, row_off: i32, col_off: i32) -> Result<VbaRangeRef, VbaError> {
    let height = range.end_row - range.start_row;
    let width = range.end_col - range.start_col;
    let sr = (range.start_row as i32 + row_off).max(1) as u32;
    let sc = (range.start_col as i32 + col_off).max(1) as u32;
    Ok(VbaRangeRef {
        sheet: range.sheet,
        start_row: sr,
        start_col: sc,
        end_row: sr + height,
        end_col: sc + width,
    })
}

fn resize_range(
    range: VbaRangeRef,
    rows: Option<i32>,
    cols: Option<i32>,
) -> Result<VbaRangeRef, VbaError> {
    let cur_rows = (range.end_row - range.start_row + 1) as i32;
    let cur_cols = (range.end_col - range.start_col + 1) as i32;
    let rows = rows.unwrap_or(cur_rows);
    let cols = cols.unwrap_or(cur_cols);
    if rows <= 0 || cols <= 0 {
        return Err(VbaError::Runtime("Resize rows/cols must be >= 1".to_string()));
    }
    Ok(VbaRangeRef {
        sheet: range.sheet,
        start_row: range.start_row,
        start_col: range.start_col,
        end_row: range.start_row + (rows as u32) - 1,
        end_col: range.start_col + (cols as u32) - 1,
    })
}

fn range_address(range: VbaRangeRef) -> Result<String, VbaError> {
    let a = absolute_a1(range.start_row, range.start_col)?;
    if range.start_row == range.end_row && range.start_col == range.end_col {
        return Ok(a);
    }
    let b = absolute_a1(range.end_row, range.end_col)?;
    Ok(format!("{a}:{b}"))
}

fn absolute_a1(row: u32, col: u32) -> Result<String, VbaError> {
    let a1 = row_col_to_a1(row, col)?;
    let mut out = String::with_capacity(a1.len() + 2);
    let mut seen_digit = false;
    for ch in a1.chars() {
        if !seen_digit && ch.is_ascii_digit() {
            out.push('$');
            seen_digit = true;
        }
        if out.is_empty() {
            out.push('$');
        }
        out.push(ch);
    }
    Ok(out)
}

fn build_proc_args(
    frame: &mut Frame,
    args: &[crate::ast::CallArg],
    proc: &ProcedureDef,
    exec: &mut Executor<'_>,
) -> Result<Vec<VbaValue>, VbaError> {
    if args.iter().any(|a| a.name.is_some()) {
        let mut values = vec![VbaValue::Empty; proc.params.len()];
        let mut next_pos = 0usize;
        for arg in args {
            let value = exec.eval_expr(frame, &arg.expr)?;
            if let Some(name) = &arg.name {
                if let Some((idx, _)) = proc
                    .params
                    .iter()
                    .enumerate()
                    .find(|(_, p)| p.name.eq_ignore_ascii_case(name))
                {
                    values[idx] = value;
                    continue;
                }
            }
            if next_pos < values.len() {
                values[next_pos] = value;
                next_pos += 1;
            }
        }
        return Ok(values);
    }

    let mut values = Vec::new();
    for arg in args {
        values.push(exec.eval_expr(frame, &arg.expr)?);
    }
    Ok(values)
}

fn parse_vba_date_string(s: &str) -> Option<NaiveDateTime> {
    let s = s.trim();
    // Common recorded-macro patterns (best-effort).
    let fmts = [
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d",
        "%m/%d/%Y %H:%M:%S",
        "%m/%d/%Y %H:%M",
        "%m/%d/%Y",
    ];
    for fmt in fmts {
        if let Ok(dt) = NaiveDateTime::parse_from_str(s, fmt) {
            return Some(dt);
        }
        if let Ok(d) = NaiveDate::parse_from_str(s, fmt) {
            return Some(d.and_hms_opt(0, 0, 0).unwrap());
        }
    }
    None
}

fn ascii_lowercase_cmp(a: &str, b: &str) -> std::cmp::Ordering {
    let a = a.as_bytes();
    let b = b.as_bytes();
    let min = a.len().min(b.len());
    for i in 0..min {
        let ca = a[i].to_ascii_lowercase();
        let cb = b[i].to_ascii_lowercase();
        if ca != cb {
            return ca.cmp(&cb);
        }
    }
    a.len().cmp(&b.len())
}

fn contains_ignore_ascii_case(haystack: &str, needle: &str) -> bool {
    let haystack = haystack.as_bytes();
    let needle = needle.as_bytes();
    if needle.is_empty() {
        return true;
    }
    if needle.len() > haystack.len() {
        return false;
    }
    for start in 0..=haystack.len() - needle.len() {
        if haystack[start..start + needle.len()].eq_ignore_ascii_case(needle) {
            return true;
        }
    }
    false
}

fn ole_base_datetime() -> NaiveDateTime {
    NaiveDate::from_ymd_opt(1899, 12, 30)
        .unwrap()
        .and_hms_opt(0, 0, 0)
        .unwrap()
}

fn datetime_to_ole_date(dt: NaiveDateTime) -> f64 {
    let base = ole_base_datetime();
    let delta = dt - base;
    delta.num_seconds() as f64 / 86_400.0
}

fn ole_date_to_datetime(serial: f64) -> NaiveDateTime {
    let base = ole_base_datetime();
    let secs = (serial * 86_400.0).round() as i64;
    base + ChronoDuration::seconds(secs)
}

fn date_add(interval: &str, number: i64, dt: NaiveDateTime) -> Result<NaiveDateTime, VbaError> {
    if interval.eq_ignore_ascii_case("d") || interval.eq_ignore_ascii_case("dd") {
        return Ok(dt + ChronoDuration::days(number));
    }
    if interval.eq_ignore_ascii_case("h") || interval.eq_ignore_ascii_case("hh") {
        return Ok(dt + ChronoDuration::hours(number));
    }
    if interval.eq_ignore_ascii_case("n") || interval.eq_ignore_ascii_case("nn") {
        return Ok(dt + ChronoDuration::minutes(number));
    }
    if interval.eq_ignore_ascii_case("s") || interval.eq_ignore_ascii_case("ss") {
        return Ok(dt + ChronoDuration::seconds(number));
    }
    if interval.eq_ignore_ascii_case("m") || interval.eq_ignore_ascii_case("mm") {
        return Ok(add_months(dt, number));
    }
    if interval.eq_ignore_ascii_case("yyyy") {
        return Ok(add_months(dt, number * 12));
    }
    let interval = interval.to_ascii_lowercase();
    Err(VbaError::Runtime(format!(
        "DateAdd unsupported interval `{interval}`"
    )))
}

fn date_diff(interval: &str, d1: NaiveDateTime, d2: NaiveDateTime) -> Result<i64, VbaError> {
    let delta = d2 - d1;
    if interval.eq_ignore_ascii_case("d") || interval.eq_ignore_ascii_case("dd") {
        return Ok(delta.num_days());
    }
    if interval.eq_ignore_ascii_case("h") || interval.eq_ignore_ascii_case("hh") {
        return Ok(delta.num_hours());
    }
    if interval.eq_ignore_ascii_case("n") || interval.eq_ignore_ascii_case("nn") {
        return Ok(delta.num_minutes());
    }
    if interval.eq_ignore_ascii_case("s") || interval.eq_ignore_ascii_case("ss") {
        return Ok(delta.num_seconds());
    }
    if interval.eq_ignore_ascii_case("m") || interval.eq_ignore_ascii_case("mm") {
        return Ok((d2.year() as i64 * 12 + d2.month() as i64)
            - (d1.year() as i64 * 12 + d1.month() as i64));
    }
    if interval.eq_ignore_ascii_case("yyyy") {
        return Ok(d2.year() as i64 - d1.year() as i64);
    }
    let interval = interval.to_ascii_lowercase();
    Err(VbaError::Runtime(format!(
        "DateDiff unsupported interval `{interval}`"
    )))
}

fn add_months(dt: NaiveDateTime, months: i64) -> NaiveDateTime {
    let date = dt.date();
    let mut year = date.year() as i64;
    let mut month0 = date.month0() as i64;
    let total = month0 + months;
    year += total.div_euclid(12);
    month0 = total.rem_euclid(12);
    let month = (month0 + 1) as u32;
    let day = date.day();

    let last_day = last_day_of_month(year as i32, month);
    let day = day.min(last_day);

    NaiveDate::from_ymd_opt(year as i32, month, day)
        .unwrap()
        .and_time(dt.time())
}

fn last_day_of_month(year: i32, month: u32) -> u32 {
    let (next_year, next_month) = if month == 12 {
        (year + 1, 1)
    } else {
        (year, month + 1)
    };
    let first_next =
        NaiveDate::from_ymd_opt(next_year, next_month, 1).expect("valid date for next month");
    (first_next - ChronoDuration::days(1)).day()
}

fn format_value(value: VbaValue, fmt: Option<&str>) -> String {
    match (value, fmt) {
        (VbaValue::Date(serial), Some(f)) => {
            let dt = ole_date_to_datetime(serial);
            format_datetime_vba(dt, f)
        }
        (VbaValue::Double(n), Some(f)) => format_number_vba(n, f),
        (v, _) => v.to_string_lossy(),
    }
}

fn format_number_vba(n: f64, fmt: &str) -> String {
    // Very small subset: formats like `0`, `0.00`.
    let fmt = fmt.trim();
    if let Some((_, frac)) = fmt.split_once('.') {
        let decimals = frac.chars().take_while(|c| *c == '0').count();
        return format!("{n:.*}", decimals);
    }
    if fmt.contains('0') {
        return format!("{n:.0}");
    }
    n.to_string()
}

fn format_datetime_vba(dt: NaiveDateTime, fmt: &str) -> String {
    let use_12h = contains_ignore_ascii_case(fmt, "am/pm");

    // Token replacement (best-effort). We intentionally keep this simple; it is not a full VBA
    // formatter.
    let mut out = String::new();
    let mut i = 0;
    while i < fmt.len() {
        let rest = &fmt[i..];
        if rest.get(..4).is_some_and(|p| p.eq_ignore_ascii_case("yyyy")) {
            out.push_str("%Y");
            i += 4;
        } else if rest.get(..2).is_some_and(|p| p.eq_ignore_ascii_case("yy")) {
            out.push_str("%y");
            i += 2;
        } else if rest.get(..4).is_some_and(|p| p.eq_ignore_ascii_case("mmmm")) {
            out.push_str("%B");
            i += 4;
        } else if rest.get(..3).is_some_and(|p| p.eq_ignore_ascii_case("mmm")) {
            out.push_str("%b");
            i += 3;
        } else if rest.get(..2).is_some_and(|p| p.eq_ignore_ascii_case("mm")) {
            out.push_str("%m");
            i += 2;
        } else if rest.get(..1).is_some_and(|p| p.eq_ignore_ascii_case("m")) {
            out.push_str("%-m");
            i += 1;
        } else if rest.get(..2).is_some_and(|p| p.eq_ignore_ascii_case("dd")) {
            out.push_str("%d");
            i += 2;
        } else if rest.get(..1).is_some_and(|p| p.eq_ignore_ascii_case("d")) {
            out.push_str("%-d");
            i += 1;
        } else if rest.get(..2).is_some_and(|p| p.eq_ignore_ascii_case("hh")) {
            out.push_str(if use_12h { "%I" } else { "%H" });
            i += 2;
        } else if rest.get(..1).is_some_and(|p| p.eq_ignore_ascii_case("h")) {
            out.push_str(if use_12h { "%-I" } else { "%-H" });
            i += 1;
        } else if rest.get(..2).is_some_and(|p| p.eq_ignore_ascii_case("nn")) {
            out.push_str("%M");
            i += 2;
        } else if rest.get(..2).is_some_and(|p| p.eq_ignore_ascii_case("ss")) {
            out.push_str("%S");
            i += 2;
        } else if rest.get(..5).is_some_and(|p| p.eq_ignore_ascii_case("am/pm")) {
            out.push_str("%p");
            i += 5;
        } else {
            let ch = rest.chars().next().unwrap();
            out.push(ch);
            i += ch.len_utf8();
        }
    }

    dt.format(&out).to_string()
}
