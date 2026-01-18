//! Excel-like Solver
//!
//! This module provides a small-but-functional optimization engine intended to
//! be integrated with the spreadsheet recalculation loop. The API is designed
//! around a [`SolverModel`] trait so the solver can run against a real sheet or
//! a unit-test mock.

mod engine_model;
mod evolutionary;
mod grg;
mod simplex;

use std::fmt;

pub use engine_model::EngineSolverModel;
pub use evolutionary::EvolutionaryOptions;
pub use grg::GrgOptions;
pub use simplex::SimplexOptions;

/// A model that can be "recalculated" after decision variables change.
///
/// In the real application this will be backed by the spreadsheet engine:
/// setting decision variables mutates cells, and `recalc` triggers dependency
/// propagation so the objective and constraint cells update.
pub trait SolverModel {
    /// Number of decision variables.
    fn num_vars(&self) -> usize;

    /// Number of constraints exposed by this model.
    fn num_constraints(&self) -> usize;

    /// Read the current variable values into `out` (length `num_vars()`).
    fn get_vars(&self, out: &mut [f64]);

    /// Update decision variables.
    fn set_vars(&mut self, vars: &[f64]) -> Result<(), SolverError>;

    /// Recalculate spreadsheet state after variables change.
    fn recalc(&mut self) -> Result<(), SolverError>;

    /// Read the objective value (after `recalc`).
    fn objective(&self) -> f64;

    /// Read the constraint values (after `recalc`) into `out` (length `num_constraints()`).
    fn constraints(&self, out: &mut [f64]);
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum SolveMethod {
    /// Linear programming (LP) with (optional) integer constraints.
    Simplex,
    /// Nonlinear optimization with constraints (continuous problems).
    GrgNonlinear,
    /// Genetic algorithm intended for non-smooth / discontinuous problems.
    Evolutionary,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ObjectiveKind {
    Maximize,
    Minimize,
    /// Drive the objective cell to a specific value.
    Target,
}

#[derive(Clone, Copy, Debug)]
pub struct Objective {
    pub kind: ObjectiveKind,
    /// Only used for [`ObjectiveKind::Target`].
    pub target_value: f64,
    /// Consider the target met if `|objective - target_value| <= target_tolerance`.
    pub target_tolerance: f64,
}

impl Objective {
    pub fn maximize() -> Self {
        Self {
            kind: ObjectiveKind::Maximize,
            target_value: 0.0,
            target_tolerance: 0.0,
        }
    }

    pub fn minimize() -> Self {
        Self {
            kind: ObjectiveKind::Minimize,
            target_value: 0.0,
            target_tolerance: 0.0,
        }
    }

    pub fn target(value: f64, tolerance: f64) -> Self {
        Self {
            kind: ObjectiveKind::Target,
            target_value: value,
            target_tolerance: tolerance.max(0.0),
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Relation {
    LessEqual,
    GreaterEqual,
    Equal,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum VarType {
    Continuous,
    Integer,
    Binary,
}

#[derive(Clone, Copy, Debug)]
pub struct VarSpec {
    pub lower: f64,
    pub upper: f64,
    pub var_type: VarType,
}

impl VarSpec {
    pub fn continuous(lower: f64, upper: f64) -> Self {
        Self {
            lower,
            upper,
            var_type: VarType::Continuous,
        }
    }

    pub fn integer(lower: f64, upper: f64) -> Self {
        Self {
            lower,
            upper,
            var_type: VarType::Integer,
        }
    }

    pub fn binary() -> Self {
        Self {
            lower: 0.0,
            upper: 1.0,
            var_type: VarType::Binary,
        }
    }
}

#[derive(Clone, Debug)]
pub struct Constraint {
    /// Index into the model's constraint vector.
    pub index: usize,
    pub relation: Relation,
    pub rhs: f64,
    /// Constraint is considered satisfied if violation <= `tolerance`.
    pub tolerance: f64,
}

impl Constraint {
    pub fn new(index: usize, relation: Relation, rhs: f64) -> Self {
        Self {
            index,
            relation,
            rhs,
            tolerance: 1e-8,
        }
    }

    pub fn with_tolerance(mut self, tolerance: f64) -> Self {
        self.tolerance = tolerance.max(0.0);
        self
    }
}

#[derive(Clone, Debug)]
pub struct SolverProblem {
    pub objective: Objective,
    pub variables: Vec<VarSpec>,
    pub constraints: Vec<Constraint>,
}

#[derive(Clone, Debug)]
pub struct Progress {
    pub iteration: usize,
    pub best_objective: f64,
    pub current_objective: f64,
    pub max_constraint_violation: f64,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum SolveStatus {
    Optimal,
    Feasible,
    Infeasible,
    Unbounded,
    IterationLimit,
    Cancelled,
}

#[derive(Clone, Debug)]
pub struct SolveOutcome {
    pub status: SolveStatus,
    pub iterations: usize,
    pub original_vars: Vec<f64>,
    pub best_vars: Vec<f64>,
    pub best_objective: f64,
    pub max_constraint_violation: f64,
}

pub struct SolveOptions<'a> {
    pub method: SolveMethod,
    /// Iteration / generation limit.
    pub max_iterations: usize,
    /// General numeric tolerance (used by all methods).
    pub tolerance: f64,
    /// Optional progress callback. Return `false` to cancel.
    pub progress: Option<&'a mut dyn FnMut(Progress) -> bool>,
    /// Whether to apply the best solution to `model` before returning.
    pub apply_solution: bool,
    pub simplex: SimplexOptions,
    pub grg: GrgOptions,
    pub evolutionary: EvolutionaryOptions,
}

impl fmt::Debug for SolveOptions<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("SolveOptions")
            .field("method", &self.method)
            .field("max_iterations", &self.max_iterations)
            .field("tolerance", &self.tolerance)
            .field("apply_solution", &self.apply_solution)
            .field("simplex", &self.simplex)
            .field("grg", &self.grg)
            .field("evolutionary", &self.evolutionary)
            .finish()
    }
}

impl<'a> SolveOptions<'a> {
    pub fn with_progress(mut self, progress: &'a mut dyn FnMut(Progress) -> bool) -> Self {
        self.progress = Some(progress);
        self
    }
}

impl<'a> Default for SolveOptions<'a> {
    fn default() -> Self {
        Self {
            method: SolveMethod::GrgNonlinear,
            max_iterations: 500,
            tolerance: 1e-8,
            progress: None,
            apply_solution: true,
            simplex: SimplexOptions::default(),
            grg: GrgOptions::default(),
            evolutionary: EvolutionaryOptions::default(),
        }
    }
}

#[derive(Debug, Clone)]
pub struct SolverError {
    pub message: String,
}

impl SolverError {
    pub fn new(message: impl Into<String>) -> Self {
        Self {
            message: message.into(),
        }
    }
}

impl fmt::Display for SolverError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}", self.message)
    }
}

impl std::error::Error for SolverError {}

pub struct Solver;

impl Solver {
    pub fn solve<M: SolverModel>(
        model: &mut M,
        problem: &SolverProblem,
        mut options: SolveOptions<'_>,
    ) -> Result<SolveOutcome, SolverError> {
        if problem.variables.len() != model.num_vars() {
            return Err(SolverError::new(format!(
                "variable spec count ({}) does not match model vars ({})",
                problem.variables.len(),
                model.num_vars()
            )));
        }
        for c in &problem.constraints {
            if c.index >= model.num_constraints() {
                return Err(SolverError::new(format!(
                    "constraint index {} out of range (model has {})",
                    c.index,
                    model.num_constraints()
                )));
            }
        }

        let n = model.num_vars();
        let mut original_vars: Vec<f64> = Vec::new();
        if original_vars.try_reserve_exact(n).is_err() {
            debug_assert!(false, "solver allocation failed (original_vars={n})");
            return Err(SolverError::new("allocation failed"));
        }
        original_vars.resize(n, 0.0);
        model.get_vars(&mut original_vars);

        // Normalize variable bounds for integer/binary vars.
        let mut normalized_problem = problem.clone();
        normalize_integer_bounds(&mut normalized_problem.variables)?;

        let mut outcome = match options.method {
            SolveMethod::Simplex => {
                simplex::solve_simplex(model, &normalized_problem, &mut options)?
            }
            SolveMethod::GrgNonlinear => grg::solve_grg(model, &normalized_problem, &mut options)?,
            SolveMethod::Evolutionary => {
                evolutionary::solve_evolutionary(model, &normalized_problem, &mut options)?
            }
        };

        if options.apply_solution && !outcome.best_vars.is_empty() {
            model.set_vars(&outcome.best_vars)?;
            model.recalc()?;
        } else {
            // Restore original state if we didn't apply the solution.
            model.set_vars(&original_vars)?;
            model.recalc()?;
        }

        outcome.original_vars = original_vars;
        Ok(outcome)
    }
}

fn normalize_integer_bounds(vars: &mut [VarSpec]) -> Result<(), SolverError> {
    for (idx, v) in vars.iter_mut().enumerate() {
        match v.var_type {
            VarType::Continuous => {}
            VarType::Integer => {
                if v.lower.is_finite() {
                    v.lower = v.lower.ceil();
                }
                if v.upper.is_finite() {
                    v.upper = v.upper.floor();
                }
                if v.lower > v.upper {
                    return Err(SolverError::new(format!(
                        "integer var {idx} has empty bounds [{}, {}]",
                        v.lower, v.upper
                    )));
                }
            }
            VarType::Binary => {
                v.lower = 0.0;
                v.upper = 1.0;
            }
        }
    }
    Ok(())
}

const NON_FINITE_PENALTY: f64 = 1e30;

fn constraint_violation(lhs: f64, constraint: &Constraint) -> f64 {
    if !lhs.is_finite() || !constraint.rhs.is_finite() {
        return NON_FINITE_PENALTY;
    }
    match constraint.relation {
        Relation::LessEqual => (lhs - constraint.rhs - constraint.tolerance).max(0.0),
        Relation::GreaterEqual => (constraint.rhs - lhs - constraint.tolerance).max(0.0),
        Relation::Equal => ((lhs - constraint.rhs).abs() - constraint.tolerance).max(0.0),
    }
}

fn max_constraint_violation(values: &[f64], constraints: &[Constraint]) -> f64 {
    constraints
        .iter()
        .map(|c| constraint_violation(values[c.index], c))
        .fold(0.0, f64::max)
}

fn is_feasible(values: &[f64], constraints: &[Constraint], tol: f64) -> bool {
    max_constraint_violation(values, constraints) <= tol
}

fn clamp_vars(vars: &mut [f64], specs: &[VarSpec]) {
    for (v, spec) in vars.iter_mut().zip(specs.iter()) {
        if spec.lower.is_finite() {
            *v = v.max(spec.lower);
        }
        if spec.upper.is_finite() {
            *v = v.min(spec.upper);
        }
        match spec.var_type {
            VarType::Continuous => {}
            VarType::Integer => {
                *v = v.round();
                if spec.lower.is_finite() {
                    *v = v.max(spec.lower);
                }
                if spec.upper.is_finite() {
                    *v = v.min(spec.upper);
                }
            }
            VarType::Binary => {
                *v = if *v >= 0.5 { 1.0 } else { 0.0 };
            }
        }
    }
}

fn objective_merit(objective: &Objective, objective_value: f64) -> f64 {
    if !objective_value.is_finite() {
        return NON_FINITE_PENALTY;
    }
    match objective.kind {
        ObjectiveKind::Maximize => -objective_value,
        ObjectiveKind::Minimize => objective_value,
        ObjectiveKind::Target => {
            let d = objective_value - objective.target_value;
            d * d
        }
    }
}

#[cfg(test)]
mod tests;
