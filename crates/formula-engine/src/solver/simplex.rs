use super::{
    clamp_vars, max_constraint_violation, ObjectiveKind, Progress, Relation, SolveOptions,
    SolveOutcome, SolveStatus, SolverError, SolverModel, SolverProblem, VarSpec, VarType,
};

#[derive(Clone, Copy, Debug)]
pub struct SimplexOptions {
    /// Maximum number of pivot operations per simplex run (phase I + phase II combined).
    pub max_pivots: usize,
    /// Maximum branch-and-bound nodes for integer/binary problems.
    pub max_bnb_nodes: usize,
    /// Integer feasibility tolerance.
    pub integer_tolerance: f64,
}

impl Default for SimplexOptions {
    fn default() -> Self {
        Self {
            max_pivots: 10_000,
            max_bnb_nodes: 1_000,
            integer_tolerance: 1e-6,
        }
    }
}

#[derive(Clone, Debug)]
struct LinearFunction {
    coeffs: Vec<f64>,
    constant: f64,
}

#[derive(Clone, Debug)]
struct LinearConstraint {
    coeffs: Vec<f64>,
    relation: Relation,
    rhs: f64,
}

#[derive(Clone, Debug)]
struct LinearProgram {
    /// Maximize `objective^T x`
    objective: Vec<f64>,
    constraints: Vec<LinearConstraint>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum LpStatus {
    Optimal,
    Infeasible,
    Unbounded,
    IterationLimit,
}

#[derive(Clone, Debug)]
struct LpSolution {
    status: LpStatus,
    x: Vec<f64>,
    objective: f64,
}

fn try_clone_f64_slice(src: &[f64]) -> Result<Vec<f64>, SolverError> {
    let mut out: Vec<f64> = Vec::new();
    if out.try_reserve_exact(src.len()).is_err() {
        debug_assert!(false, "solver allocation failed (len={})", src.len());
        return Err(SolverError::new("allocation failed"));
    }
    out.extend_from_slice(src);
    Ok(out)
}

fn try_zeros_f64(len: usize) -> Result<Vec<f64>, SolverError> {
    let mut out: Vec<f64> = Vec::new();
    if out.try_reserve_exact(len).is_err() {
        debug_assert!(false, "solver allocation failed (len={len})");
        return Err(SolverError::new("allocation failed"));
    }
    out.resize(len, 0.0);
    Ok(out)
}

fn try_clone_lp(lp: &LinearProgram) -> Result<LinearProgram, SolverError> {
    let objective = try_clone_f64_slice(&lp.objective)?;

    let mut constraints: Vec<LinearConstraint> = Vec::new();
    if constraints.try_reserve_exact(lp.constraints.len()).is_err() {
        debug_assert!(
            false,
            "solver allocation failed (constraints={})",
            lp.constraints.len()
        );
        return Err(SolverError::new("allocation failed"));
    }
    for c in &lp.constraints {
        let coeffs = try_clone_f64_slice(&c.coeffs)?;
        constraints.push(LinearConstraint {
            coeffs,
            relation: c.relation,
            rhs: c.rhs,
        });
    }

    Ok(LinearProgram {
        objective,
        constraints,
    })
}

pub(crate) fn solve_simplex<M: SolverModel>(
    model: &mut M,
    problem: &SolverProblem,
    options: &mut SolveOptions<'_>,
) -> Result<SolveOutcome, SolverError> {
    // Evaluate at current values to define the sampling base point.
    let n_vars = model.num_vars();
    let mut vars: Vec<f64> = Vec::new();
    if vars.try_reserve_exact(n_vars).is_err() {
        debug_assert!(false, "solver allocation failed (vars={n_vars})");
        return Err(SolverError::new("allocation failed"));
    }
    vars.resize(n_vars, 0.0);
    model.get_vars(&mut vars);
    clamp_vars(&mut vars, &problem.variables);

    let (objective_fn, constraint_fns) = infer_linear_functions(model, problem, &vars)?;

    let mut bounds_lower: Vec<f64> = Vec::new();
    if bounds_lower.try_reserve_exact(problem.variables.len()).is_err() {
        debug_assert!(false, "solver allocation failed (bounds_lower)");
        return Err(SolverError::new("allocation failed"));
    }
    for v in &problem.variables {
        bounds_lower.push(v.lower);
    }
    for (idx, v) in problem.variables.iter().enumerate() {
        if !bounds_lower[idx].is_finite() {
            // In Excel, "Make Unconstrained Variables Non-Negative" is the default for simplex.
            // We follow that convention here when lower bounds are not supplied.
            bounds_lower[idx] = 0.0;
        }
        if v.upper.is_finite() && bounds_lower[idx] > v.upper {
            return Ok(SolveOutcome {
                status: SolveStatus::Infeasible,
                iterations: 0,
                original_vars: Vec::new(),
                best_vars: Vec::new(),
                best_objective: f64::NAN,
                max_constraint_violation: f64::INFINITY,
            });
        }
    }

    let (lp, shift, decision_len) =
        build_lp(problem, &objective_fn, &constraint_fns, &bounds_lower)?;

    let mut integer_indices: Vec<usize> = Vec::new();
    if integer_indices
        .try_reserve_exact(problem.variables.len())
        .is_err()
    {
        debug_assert!(false, "solver allocation failed (integer_indices)");
        return Err(SolverError::new("allocation failed"));
    }
    for (idx, v) in problem.variables.iter().enumerate() {
        if matches!(v.var_type, VarType::Integer | VarType::Binary) {
            integer_indices.push(idx);
        }
    }

    let mut nodes_searched = 0usize;
    let mut best_lp_solution: Option<LpSolution> = None;
    let mut cancelled = false;

    branch_and_bound(
        &lp,
        problem,
        &objective_fn,
        &constraint_fns,
        &shift,
        decision_len,
        &integer_indices,
        options.simplex,
        &mut nodes_searched,
        &mut best_lp_solution,
        &mut options.progress,
        &mut cancelled,
    )?;

    let best_lp_solution = best_lp_solution.unwrap_or(LpSolution {
        status: LpStatus::Infeasible,
        x: Vec::new(),
        objective: f64::NAN,
    });

    let mut status = match best_lp_solution.status {
        LpStatus::Optimal => SolveStatus::Optimal,
        LpStatus::Infeasible => SolveStatus::Infeasible,
        LpStatus::Unbounded => SolveStatus::Unbounded,
        LpStatus::IterationLimit => SolveStatus::IterationLimit,
    };

    if cancelled {
        status = SolveStatus::Cancelled;
    } else if !integer_indices.is_empty() && nodes_searched >= options.simplex.max_bnb_nodes {
        // Search was truncated by the node limit; the best solution we have is
        // not guaranteed globally optimal.
        status = SolveStatus::IterationLimit;
    }

    if best_lp_solution.x.is_empty() {
        return Ok(SolveOutcome {
            status,
            iterations: nodes_searched,
            original_vars: Vec::new(),
            best_vars: Vec::new(),
            best_objective: f64::NAN,
            max_constraint_violation: f64::INFINITY,
        });
    }

    let mut best_vars: Vec<f64> = Vec::new();
    if best_vars.try_reserve_exact(decision_len).is_err() {
        debug_assert!(false, "solver allocation failed (best_vars={decision_len})");
        return Err(SolverError::new("allocation failed"));
    }
    for (y, l) in best_lp_solution.x[..decision_len].iter().zip(shift.iter()) {
        best_vars.push(y + l);
    }

    // Project back into bounds / variable domains.
    clamp_vars(&mut best_vars, &problem.variables);

    let m_constraints = model.num_constraints();
    let mut constraint_values: Vec<f64> = Vec::new();
    if constraint_values.try_reserve_exact(m_constraints).is_err() {
        debug_assert!(false, "solver allocation failed (constraints={m_constraints})");
        return Err(SolverError::new("allocation failed"));
    }
    constraint_values.resize(m_constraints, 0.0);
    model.set_vars(&best_vars)?;
    model.recalc()?;
    let best_objective = model.objective();
    model.constraints(&mut constraint_values);
    let max_violation = max_constraint_violation(&constraint_values, &problem.constraints);

    if problem.objective.kind == ObjectiveKind::Target
        && matches!(status, SolveStatus::Optimal | SolveStatus::Feasible)
    {
        let tol = problem
            .objective
            .target_tolerance
            .max(options.tolerance.max(0.0));
        if (best_objective - problem.objective.target_value).abs() <= tol {
            status = SolveStatus::Optimal;
        } else {
            status = SolveStatus::Feasible;
        }
    }

    // Progress update (a single update at end; simplex is not iterative in the same way).
    if let Some(progress) = options.progress.as_deref_mut() {
        let _ = progress(Progress {
            iteration: nodes_searched,
            best_objective,
            current_objective: best_objective,
            max_constraint_violation: max_violation,
        });
    }

    Ok(SolveOutcome {
        status,
        iterations: nodes_searched,
        original_vars: Vec::new(),
        best_vars,
        best_objective,
        max_constraint_violation: max_violation,
    })
}

fn infer_linear_functions<M: SolverModel>(
    model: &mut M,
    problem: &SolverProblem,
    base_vars: &[f64],
) -> Result<(LinearFunction, Vec<LinearFunction>), SolverError> {
    let n = model.num_vars();
    let m = model.num_constraints();

    let mut base_constraints: Vec<f64> = Vec::new();
    if base_constraints.try_reserve_exact(m).is_err() {
        debug_assert!(false, "solver allocation failed (constraints={m})");
        return Err(SolverError::new("allocation failed"));
    }
    base_constraints.resize(m, 0.0);
    model.set_vars(base_vars)?;
    model.recalc()?;
    let base_obj = model.objective();
    model.constraints(&mut base_constraints);

    if !base_obj.is_finite() {
        return Err(SolverError::new(format!(
            "objective is not finite at the starting point ({base_obj}); simplex requires a valid linear model"
        )));
    }
    if let Some((idx, val)) = base_constraints
        .iter()
        .enumerate()
        .find(|(_, v)| !v.is_finite())
    {
        return Err(SolverError::new(format!(
            "constraint {idx} is not finite at the starting point ({val}); simplex requires a valid linear model"
        )));
    }

    let mut obj_coeffs: Vec<f64> = Vec::new();
    if obj_coeffs.try_reserve_exact(n).is_err() {
        debug_assert!(false, "solver allocation failed (obj_coeffs={n})");
        return Err(SolverError::new("allocation failed"));
    }
    obj_coeffs.resize(n, 0.0);

    let mut constraint_coeffs: Vec<Vec<f64>> = Vec::new();
    if constraint_coeffs.try_reserve_exact(m).is_err() {
        debug_assert!(false, "solver allocation failed (constraint_coeffs={m})");
        return Err(SolverError::new("allocation failed"));
    }
    for _ in 0..m {
        let mut row: Vec<f64> = Vec::new();
        if row.try_reserve_exact(n).is_err() {
            debug_assert!(false, "solver allocation failed (constraint_row={n})");
            return Err(SolverError::new("allocation failed"));
        }
        row.resize(n, 0.0);
        constraint_coeffs.push(row);
    }

    for j in 0..n {
        let step = choose_step(base_vars[j], &problem.variables[j]);
        let mut vars: Vec<f64> = Vec::new();
        if vars.try_reserve_exact(base_vars.len()).is_err() {
            debug_assert!(false, "solver allocation failed (vars={})", base_vars.len());
            return Err(SolverError::new("allocation failed"));
        }
        vars.extend_from_slice(base_vars);
        vars[j] += step;
        clamp_vars(&mut vars, &problem.variables);

        let denom = vars[j] - base_vars[j];
        if denom.abs() < 1e-12 {
            // Variable cannot move (e.g. fixed bounds); treat as having a zero
            // coefficient for the inferred linear model.
            obj_coeffs[j] = 0.0;
            for i in 0..m {
                constraint_coeffs[i][j] = 0.0;
            }
            continue;
        }

        let mut constraints: Vec<f64> = Vec::new();
        if constraints.try_reserve_exact(m).is_err() {
            debug_assert!(false, "solver allocation failed (constraints={m})");
            return Err(SolverError::new("allocation failed"));
        }
        constraints.resize(m, 0.0);
        model.set_vars(&vars)?;
        model.recalc()?;
        let obj = model.objective();
        model.constraints(&mut constraints);
        if !obj.is_finite() {
            return Err(SolverError::new(format!(
                "objective is not finite while inferring coefficient for var {j} ({obj})"
            )));
        }
        if let Some((idx, val)) = constraints.iter().enumerate().find(|(_, v)| !v.is_finite()) {
            return Err(SolverError::new(format!(
                "constraint {idx} is not finite while inferring coefficient for var {j} ({val})"
            )));
        }

        obj_coeffs[j] = (obj - base_obj) / denom;
        for i in 0..m {
            constraint_coeffs[i][j] = (constraints[i] - base_constraints[i]) / denom;
        }
    }

    let obj_constant = base_obj - dot(&obj_coeffs, base_vars);
    let objective_fn = LinearFunction {
        coeffs: obj_coeffs,
        constant: obj_constant,
    };

    let mut constraint_fns: Vec<LinearFunction> = Vec::new();
    if constraint_fns.try_reserve_exact(m).is_err() {
        debug_assert!(false, "solver allocation failed (constraint_fns={m})");
        return Err(SolverError::new("allocation failed"));
    }
    for (i, coeffs) in constraint_coeffs.into_iter().enumerate() {
        let cst = base_constraints[i] - dot(&coeffs, base_vars);
        constraint_fns.push(LinearFunction { coeffs, constant: cst });
    }

    Ok((objective_fn, constraint_fns))
}

fn choose_step(x: f64, spec: &VarSpec) -> f64 {
    let base = 1.0_f64.max(x.abs());
    let mut step = 1.0;
    if spec.var_type == VarType::Binary {
        // Ensure we actually flip the bit when inferring coefficients.
        step = if x < 0.5 { 1.0 } else { -1.0 };
    }

    if spec.lower.is_finite() && x + step < spec.lower {
        step = (spec.lower - x).max(1e-3 * base);
    }
    if spec.upper.is_finite() && x + step > spec.upper {
        step = -(x - spec.upper).max(1e-3 * base);
    }
    if step == 0.0 {
        step = 1e-3 * base;
    }
    step
}

fn build_lp(
    problem: &SolverProblem,
    objective_fn: &LinearFunction,
    constraint_fns: &[LinearFunction],
    lower_bounds: &[f64],
) -> Result<(LinearProgram, Vec<f64>, usize), SolverError> {
    let n = problem.variables.len();
    let mut shift: Vec<f64> = Vec::new();
    if shift.try_reserve_exact(lower_bounds.len()).is_err() {
        debug_assert!(false, "solver allocation failed (shift={})", lower_bounds.len());
        return Err(SolverError::new("allocation failed"));
    }
    shift.extend_from_slice(lower_bounds);

    // Decision variables are shifted so they are all >= 0.
    let mut obj: Vec<f64> = Vec::new();
    if obj.try_reserve_exact(n).is_err() {
        debug_assert!(false, "solver allocation failed (obj={n})");
        return Err(SolverError::new("allocation failed"));
    }
    obj.resize(n, 0.0);
    match problem.objective.kind {
        ObjectiveKind::Maximize => obj.clone_from_slice(&objective_fn.coeffs),
        ObjectiveKind::Minimize => {
            for j in 0..n {
                obj[j] = -objective_fn.coeffs[j];
            }
        }
        ObjectiveKind::Target => {
            // We'll add an auxiliary variable `t` and minimize it (by maximizing -t).
            let mut tmp: Vec<f64> = Vec::new();
            if tmp.try_reserve_exact(n + 1).is_err() {
                debug_assert!(false, "solver allocation failed (obj={})", n + 1);
                return Err(SolverError::new("allocation failed"));
            }
            tmp.resize(n + 1, 0.0);
            obj = tmp;
            obj[n] = -1.0;
        }
    }

    let upper_bound_constraints = problem
        .variables
        .iter()
        .filter(|v| v.upper.is_finite())
        .count();
    let target_constraints = if problem.objective.kind == ObjectiveKind::Target {
        2
    } else {
        0
    };
    let expected_constraints = problem.constraints.len() + upper_bound_constraints + target_constraints;

    let mut constraints: Vec<LinearConstraint> = Vec::new();
    if constraints.try_reserve_exact(expected_constraints).is_err() {
        debug_assert!(
            false,
            "solver allocation failed (constraints={expected_constraints})"
        );
        return Err(SolverError::new("allocation failed"));
    }

    // User constraints.
    for constraint in &problem.constraints {
        let f = &constraint_fns[constraint.index];
        let mut coeffs = try_clone_f64_slice(&f.coeffs)?;
        let mut rhs = constraint.rhs - f.constant - dot(&coeffs, &shift);
        let mut relation = constraint.relation;

        // Ensure RHS >= 0 for a cleaner initial basis.
        if rhs < 0.0 {
            rhs = -rhs;
            for c in &mut coeffs {
                *c = -*c;
            }
            relation = match relation {
                Relation::LessEqual => Relation::GreaterEqual,
                Relation::GreaterEqual => Relation::LessEqual,
                Relation::Equal => Relation::Equal,
            };
        }

        constraints.push(LinearConstraint {
            coeffs,
            relation,
            rhs,
        });
    }

    // Upper bounds as constraints.
    for (j, var) in problem.variables.iter().enumerate() {
        if var.upper.is_finite() {
            let ub = var.upper - shift[j];
            let mut coeffs: Vec<f64> = Vec::new();
            if coeffs.try_reserve_exact(n).is_err() {
                debug_assert!(false, "solver allocation failed (coeffs={n})");
                return Err(SolverError::new("allocation failed"));
            }
            coeffs.resize(n, 0.0);
            coeffs[j] = 1.0;
            constraints.push(LinearConstraint {
                coeffs,
                relation: Relation::LessEqual,
                rhs: ub,
            });
        }
    }

    // Target objective constraints.
    let decision_len = if problem.objective.kind == ObjectiveKind::Target {
        n + 1
    } else {
        n
    };

    if problem.objective.kind == ObjectiveKind::Target {
        let constant_y = objective_fn.constant + dot(&objective_fn.coeffs, &shift);
        let target_rhs = problem.objective.target_value - constant_y;

        let mut c1: Vec<f64> = Vec::new();
        if c1.try_reserve_exact(objective_fn.coeffs.len() + 1).is_err() {
            debug_assert!(
                false,
                "solver allocation failed (coeffs={})",
                objective_fn.coeffs.len() + 1
            );
            return Err(SolverError::new("allocation failed"));
        }
        c1.extend_from_slice(&objective_fn.coeffs);
        c1.push(-1.0); // -t
        constraints.push(LinearConstraint {
            coeffs: c1,
            relation: Relation::LessEqual,
            rhs: target_rhs,
        });

        let mut c2: Vec<f64> = Vec::new();
        if c2.try_reserve_exact(objective_fn.coeffs.len() + 1).is_err() {
            debug_assert!(
                false,
                "solver allocation failed (coeffs={})",
                objective_fn.coeffs.len() + 1
            );
            return Err(SolverError::new("allocation failed"));
        }
        for v in &objective_fn.coeffs {
            c2.push(-*v);
        }
        c2.push(-1.0); // -t
        constraints.push(LinearConstraint {
            coeffs: c2,
            relation: Relation::LessEqual,
            rhs: -target_rhs,
        });

        // t >= 0 already holds as a decision var (non-negative in simplex).
    }

    // Make all constraints have the right coefficient length.
    for c in &mut constraints {
        if c.coeffs.len() < decision_len {
            let additional = decision_len - c.coeffs.len();
            if c.coeffs.try_reserve_exact(additional).is_err() {
                debug_assert!(
                    false,
                    "solver allocation failed (coeffs resize -> {decision_len})"
                );
                return Err(SolverError::new("allocation failed"));
            }
            c.coeffs.resize(decision_len, 0.0);
        }
    }

    Ok((
        LinearProgram {
            objective: obj,
            constraints,
        },
        shift,
        decision_len,
    ))
}

fn branch_and_bound(
    lp: &LinearProgram,
    problem: &SolverProblem,
    objective_fn: &LinearFunction,
    constraint_fns: &[LinearFunction],
    shift: &[f64],
    decision_len: usize,
    integer_indices: &[usize],
    options: SimplexOptions,
    nodes_searched: &mut usize,
    best_solution: &mut Option<LpSolution>,
    progress: &mut Option<&mut dyn FnMut(Progress) -> bool>,
    cancelled: &mut bool,
) -> Result<(), SolverError> {
    if *cancelled || *nodes_searched >= options.max_bnb_nodes {
        return Ok(());
    }
    *nodes_searched += 1;

    let mut lp_solution = solve_lp(lp, options.max_pivots)?;

    if lp_solution.status != LpStatus::Optimal || *cancelled {
        return Ok(());
    }

    // Prune using relaxation bound.
    if let Some(best) = best_solution {
        if lp_solution.objective <= best.objective + 1e-10 {
            return Ok(());
        }
    }

    let current_objective = evaluate_objective_for_solution(objective_fn, shift, &lp_solution.x);
    let current_violation =
        evaluate_max_violation_for_solution(constraint_fns, shift, problem, &lp_solution.x);

    // Check integrality.
    let fractional_var = integer_indices
        .iter()
        .copied()
        .filter(|&idx| idx < decision_len)
        .find_map(|idx| {
            let v = lp_solution.x[idx];
            let nearest = v.round();
            if (v - nearest).abs() > options.integer_tolerance {
                Some((idx, v))
            } else {
                None
            }
        });

    if fractional_var.is_none() {
        // Integer-feasible; accept if better.
        if best_solution
            .as_ref()
            .map_or(true, |best| lp_solution.objective > best.objective)
        {
            // Ensure non-decision vars don't leak into the solution.
            lp_solution.x.truncate(decision_len);
            *best_solution = Some(lp_solution);
        }
    }

    if let Some(callback) = progress.as_deref_mut() {
        let best_objective = best_solution
            .as_ref()
            .map(|s| evaluate_objective_for_solution(objective_fn, shift, &s.x))
            .unwrap_or(f64::NAN);

        if !callback(Progress {
            iteration: *nodes_searched,
            best_objective,
            current_objective,
            max_constraint_violation: current_violation,
        }) {
            *cancelled = true;
        }
    }

    if *cancelled {
        return Ok(());
    }

    if let Some((idx, value)) = fractional_var {
        let floor_v = value.floor();
        let ceil_v = value.ceil();

        // Branch 1: x_idx <= floor_v
        if floor_v.is_finite() {
            let mut lp1 = try_clone_lp(lp)?;
            if lp1.constraints.try_reserve_exact(1).is_err() {
                debug_assert!(false, "solver allocation failed (constraints +1)");
                return Err(SolverError::new("allocation failed"));
            }
            let mut coeffs = try_zeros_f64(decision_len)?;
            coeffs[idx] = 1.0;
            lp1.constraints.push(LinearConstraint {
                coeffs,
                relation: Relation::LessEqual,
                rhs: floor_v,
            });
            branch_and_bound(
                &lp1,
                problem,
                objective_fn,
                constraint_fns,
                shift,
                decision_len,
                integer_indices,
                options,
                nodes_searched,
                best_solution,
                progress,
                cancelled,
            )?;
        }

        // Branch 2: x_idx >= ceil_v  ->  -x_idx <= -ceil_v
        if ceil_v.is_finite() {
            let mut lp2 = try_clone_lp(lp)?;
            if lp2.constraints.try_reserve_exact(1).is_err() {
                debug_assert!(false, "solver allocation failed (constraints +1)");
                return Err(SolverError::new("allocation failed"));
            }
            let mut coeffs = try_zeros_f64(decision_len)?;
            coeffs[idx] = -1.0;
            lp2.constraints.push(LinearConstraint {
                coeffs,
                relation: Relation::LessEqual,
                rhs: -ceil_v,
            });
            branch_and_bound(
                &lp2,
                problem,
                objective_fn,
                constraint_fns,
                shift,
                decision_len,
                integer_indices,
                options,
                nodes_searched,
                best_solution,
                progress,
                cancelled,
            )?;
        }
    }

    Ok(())
}

fn solve_lp(lp: &LinearProgram, max_pivots: usize) -> Result<LpSolution, SolverError> {
    let n = lp.objective.len();

    let mut slack_count = 0usize;
    let mut surplus_count = 0usize;
    let mut artificial_count = 0usize;
    for c in &lp.constraints {
        let relation = if c.rhs < 0.0 {
            match c.relation {
                Relation::LessEqual => Relation::GreaterEqual,
                Relation::GreaterEqual => Relation::LessEqual,
                Relation::Equal => Relation::Equal,
            }
        } else {
            c.relation
        };

        match relation {
            Relation::LessEqual => slack_count += 1,
            Relation::GreaterEqual => {
                surplus_count += 1;
                artificial_count += 1;
            }
            Relation::Equal => artificial_count += 1,
        }
    }

    let total_vars = n + slack_count + surplus_count + artificial_count;
    let m = lp.constraints.len();

    let cols = total_vars + 1;
    let rows = m + 1;

    let mut tableau: Vec<Vec<f64>> = Vec::new();
    if tableau.try_reserve_exact(rows).is_err() {
        debug_assert!(false, "solver allocation failed (tableau rows={rows})");
        return Err(SolverError::new("allocation failed"));
    }
    for _ in 0..rows {
        tableau.push(try_zeros_f64(cols)?);
    }

    let mut basis: Vec<usize> = Vec::new();
    if basis.try_reserve_exact(m).is_err() {
        debug_assert!(false, "solver allocation failed (basis={m})");
        return Err(SolverError::new("allocation failed"));
    }
    basis.resize(m, 0);

    let slack_offset = n;
    let surplus_offset = slack_offset + slack_count;
    let artificial_offset = surplus_offset + surplus_count;

    let mut slack_idx = 0usize;
    let mut surplus_idx = 0usize;
    let mut artificial_idx = 0usize;

    for (row, c) in lp.constraints.iter().enumerate() {
        let (coeff_sign, rhs, relation) = if c.rhs < 0.0 {
            let relation = match c.relation {
                Relation::LessEqual => Relation::GreaterEqual,
                Relation::GreaterEqual => Relation::LessEqual,
                Relation::Equal => Relation::Equal,
            };
            (-1.0, -c.rhs, relation)
        } else {
            (1.0, c.rhs, c.relation)
        };

        for j in 0..n {
            tableau[row][j] = coeff_sign * c.coeffs[j];
        }
        tableau[row][total_vars] = rhs;

        match relation {
            Relation::LessEqual => {
                let col = slack_offset + slack_idx;
                slack_idx += 1;
                tableau[row][col] = 1.0;
                basis[row] = col;
            }
            Relation::GreaterEqual => {
                let surplus_col = surplus_offset + surplus_idx;
                surplus_idx += 1;
                tableau[row][surplus_col] = -1.0;

                let artificial_col = artificial_offset + artificial_idx;
                artificial_idx += 1;
                tableau[row][artificial_col] = 1.0;
                basis[row] = artificial_col;
            }
            Relation::Equal => {
                let artificial_col = artificial_offset + artificial_idx;
                artificial_idx += 1;
                tableau[row][artificial_col] = 1.0;
                basis[row] = artificial_col;
            }
        }
    }

    // Phase I objective: maximize -sum(artificial)
    for col in artificial_offset..total_vars {
        tableau[m][col] = -1.0;
    }

    // Make objective row canonical for the initial basis.
    for row in 0..m {
        let basic = basis[row];
        if basic >= artificial_offset {
            // Add this row to the objective to eliminate the -1 coefficient.
            let factor = 1.0;
            for col in 0..=total_vars {
                tableau[m][col] += factor * tableau[row][col];
            }
        }
    }

    let mut pivots_used = 0usize;
    match simplex_iterate(
        &mut tableau,
        &mut basis,
        max_pivots,
        &mut pivots_used,
        |col| col < total_vars, // allow all vars in phase I
    ) {
        LpStatus::Optimal => {}
        status => {
            return Ok(LpSolution {
                status,
                x: Vec::new(),
                objective: f64::NAN,
            });
        }
    }

    let phase1_obj = tableau[m][total_vars];
    if phase1_obj < -1e-8 {
        return Ok(LpSolution {
            status: LpStatus::Infeasible,
            x: Vec::new(),
            objective: f64::NAN,
        });
    }

    // Phase II: set real objective.
    for col in 0..total_vars {
        tableau[m][col] = if col < n { lp.objective[col] } else { 0.0 };
    }
    tableau[m][total_vars] = 0.0;

    // Canonicalize w.r.t current basis.
    for row in 0..m {
        let basic = basis[row];
        if basic < n {
            let factor = tableau[m][basic];
            if factor.abs() > 1e-12 {
                for col in 0..=total_vars {
                    tableau[m][col] -= factor * tableau[row][col];
                }
            }
        }
    }

    let phase2_status = simplex_iterate(
        &mut tableau,
        &mut basis,
        max_pivots,
        &mut pivots_used,
        |col| col < artificial_offset, // do not allow artificial vars to enter
    );

    let mut x = try_zeros_f64(n)?;
    for row in 0..m {
        let basic = basis[row];
        if basic < n {
            x[basic] = tableau[row][total_vars];
        }
    }

    Ok(LpSolution {
        status: phase2_status,
        x,
        objective: tableau[m][total_vars],
    })
}

fn simplex_iterate<F: Fn(usize) -> bool>(
    tableau: &mut [Vec<f64>],
    basis: &mut [usize],
    max_pivots: usize,
    pivots_used: &mut usize,
    allow_entering: F,
) -> LpStatus {
    let m = basis.len();
    let total_vars = tableau[0].len() - 1;
    let rhs_col = total_vars;

    let eps = 1e-10;

    while *pivots_used < max_pivots {
        // Choose entering variable (Bland's rule).
        let mut entering: Option<usize> = None;
        for col in 0..total_vars {
            if !allow_entering(col) {
                continue;
            }
            if tableau[m][col] > eps {
                entering = Some(col);
                break;
            }
        }

        let Some(entering) = entering else {
            return LpStatus::Optimal;
        };

        // Leaving variable via minimum ratio test.
        let mut leaving_row: Option<usize> = None;
        let mut best_ratio = f64::INFINITY;
        for row in 0..m {
            let a = tableau[row][entering];
            if a > eps {
                let ratio = tableau[row][rhs_col] / a;
                if ratio < best_ratio - 1e-12
                    || ((ratio - best_ratio).abs() <= 1e-12
                        && basis[row] < basis[leaving_row.unwrap_or(row)])
                {
                    best_ratio = ratio;
                    leaving_row = Some(row);
                }
            }
        }

        let Some(leaving_row) = leaving_row else {
            return LpStatus::Unbounded;
        };

        pivot(tableau, basis, leaving_row, entering);
        *pivots_used += 1;
    }

    LpStatus::IterationLimit
}

fn pivot(tableau: &mut [Vec<f64>], basis: &mut [usize], leaving_row: usize, entering: usize) {
    let m = basis.len();
    let total_vars = tableau[0].len() - 1;
    let pivot_val = tableau[leaving_row][entering];

    // Normalize leaving row.
    for col in 0..=total_vars {
        tableau[leaving_row][col] /= pivot_val;
    }

    // Eliminate pivot column from all other rows including objective.
    for row in 0..=m {
        if row == leaving_row {
            continue;
        }
        let factor = tableau[row][entering];
        if factor.abs() < 1e-12 {
            continue;
        }
        for col in 0..=total_vars {
            tableau[row][col] -= factor * tableau[leaving_row][col];
        }
    }

    basis[leaving_row] = entering;
}

fn dot(a: &[f64], b: &[f64]) -> f64 {
    a.iter().zip(b.iter()).map(|(x, y)| x * y).sum()
}

fn evaluate_objective_for_solution(objective_fn: &LinearFunction, shift: &[f64], y: &[f64]) -> f64 {
    let n = shift.len().min(y.len());
    let mut acc = objective_fn.constant;
    acc += dot(&objective_fn.coeffs[..n], &shift[..n]);
    acc += dot(&objective_fn.coeffs[..n], &y[..n]);
    acc
}

fn evaluate_max_violation_for_solution(
    constraint_fns: &[LinearFunction],
    shift: &[f64],
    problem: &SolverProblem,
    y: &[f64],
) -> f64 {
    let n = shift.len().min(y.len());
    problem
        .constraints
        .iter()
        .map(|c| {
            let f = &constraint_fns[c.index];
            let mut lhs = f.constant + dot(&f.coeffs[..n], &shift[..n]);
            lhs += dot(&f.coeffs[..n], &y[..n]);
            super::constraint_violation(lhs, c)
        })
        .fold(0.0, f64::max)
}
