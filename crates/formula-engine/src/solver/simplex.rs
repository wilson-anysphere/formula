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

pub(crate) fn solve_simplex<M: SolverModel>(
    model: &mut M,
    problem: &SolverProblem,
    options: &mut SolveOptions<'_>,
) -> Result<SolveOutcome, SolverError> {
    // Evaluate at current values to define the sampling base point.
    let mut vars = vec![0.0; model.num_vars()];
    model.get_vars(&mut vars);
    clamp_vars(&mut vars, &problem.variables);

    let (objective_fn, constraint_fns) = infer_linear_functions(model, problem, &vars)?;

    let mut bounds_lower: Vec<f64> = problem.variables.iter().map(|v| v.lower).collect();
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

    let integer_indices: Vec<usize> = problem
        .variables
        .iter()
        .enumerate()
        .filter_map(|(idx, v)| match v.var_type {
            VarType::Continuous => None,
            VarType::Integer | VarType::Binary => Some(idx),
        })
        .collect();

    let mut nodes_searched = 0usize;
    let mut best_lp_solution: Option<LpSolution> = None;

    branch_and_bound(
        &lp,
        decision_len,
        &integer_indices,
        options.simplex,
        &mut nodes_searched,
        &mut best_lp_solution,
    );

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

    let mut best_vars: Vec<f64> = best_lp_solution.x[..decision_len]
        .iter()
        .zip(shift.iter())
        .map(|(y, l)| y + l)
        .collect();

    // Project back into bounds / variable domains.
    clamp_vars(&mut best_vars, &problem.variables);

    let mut constraint_values = vec![0.0; model.num_constraints()];
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

    let mut base_constraints = vec![0.0; m];
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

    let mut obj_coeffs = vec![0.0; n];
    let mut constraint_coeffs: Vec<Vec<f64>> = vec![vec![0.0; n]; m];

    for j in 0..n {
        let step = choose_step(base_vars[j], &problem.variables[j]);
        let mut vars = base_vars.to_vec();
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

        let mut constraints = vec![0.0; m];
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

    let mut constraint_fns = Vec::with_capacity(m);
    for i in 0..m {
        let cst = base_constraints[i] - dot(&constraint_coeffs[i], base_vars);
        constraint_fns.push(LinearFunction {
            coeffs: constraint_coeffs[i].clone(),
            constant: cst,
        });
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
    let shift: Vec<f64> = lower_bounds.to_vec();

    // Decision variables are shifted so they are all >= 0.
    let mut obj = vec![0.0; n];
    match problem.objective.kind {
        ObjectiveKind::Maximize => obj.clone_from_slice(&objective_fn.coeffs),
        ObjectiveKind::Minimize => {
            for j in 0..n {
                obj[j] = -objective_fn.coeffs[j];
            }
        }
        ObjectiveKind::Target => {
            // We'll add an auxiliary variable `t` and minimize it (by maximizing -t).
            obj = vec![0.0; n + 1];
            obj[n] = -1.0;
        }
    }

    let mut constraints: Vec<LinearConstraint> = Vec::new();

    // User constraints.
    for constraint in &problem.constraints {
        let f = &constraint_fns[constraint.index];
        let mut coeffs = f.coeffs.clone();
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
            let mut coeffs = vec![0.0; n];
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

        let mut c1 = objective_fn.coeffs.clone();
        c1.push(-1.0); // -t
        constraints.push(LinearConstraint {
            coeffs: c1,
            relation: Relation::LessEqual,
            rhs: target_rhs,
        });

        let mut c2: Vec<f64> = objective_fn.coeffs.iter().map(|v| -*v).collect();
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
    decision_len: usize,
    integer_indices: &[usize],
    options: SimplexOptions,
    nodes_searched: &mut usize,
    best_solution: &mut Option<LpSolution>,
) {
    if *nodes_searched >= options.max_bnb_nodes {
        return;
    }
    *nodes_searched += 1;

    let mut lp_solution = solve_lp(lp, options.max_pivots);

    if lp_solution.status != LpStatus::Optimal {
        return;
    }

    // Prune using relaxation bound.
    if let Some(best) = best_solution {
        if lp_solution.objective <= best.objective + 1e-10 {
            return;
        }
    }

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

    if let Some((idx, value)) = fractional_var {
        let floor_v = value.floor();
        let ceil_v = value.ceil();

        // Branch 1: x_idx <= floor_v
        if floor_v.is_finite() {
            let mut lp1 = lp.clone();
            let mut coeffs = vec![0.0; decision_len];
            coeffs[idx] = 1.0;
            lp1.constraints.push(LinearConstraint {
                coeffs,
                relation: Relation::LessEqual,
                rhs: floor_v,
            });
            branch_and_bound(
                &lp1,
                decision_len,
                integer_indices,
                options,
                nodes_searched,
                best_solution,
            );
        }

        // Branch 2: x_idx >= ceil_v  ->  -x_idx <= -ceil_v
        if ceil_v.is_finite() {
            let mut lp2 = lp.clone();
            let mut coeffs = vec![0.0; decision_len];
            coeffs[idx] = -1.0;
            lp2.constraints.push(LinearConstraint {
                coeffs,
                relation: Relation::LessEqual,
                rhs: -ceil_v,
            });
            branch_and_bound(
                &lp2,
                decision_len,
                integer_indices,
                options,
                nodes_searched,
                best_solution,
            );
        }
        return;
    }

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

fn solve_lp(lp: &LinearProgram, max_pivots: usize) -> LpSolution {
    let n = lp.objective.len();
    let mut constraints = lp.constraints.clone();

    // Ensure RHS >= 0 for tableau construction.
    for c in &mut constraints {
        if c.rhs < 0.0 {
            c.rhs = -c.rhs;
            for v in &mut c.coeffs {
                *v = -*v;
            }
            c.relation = match c.relation {
                Relation::LessEqual => Relation::GreaterEqual,
                Relation::GreaterEqual => Relation::LessEqual,
                Relation::Equal => Relation::Equal,
            };
        }
    }

    let slack_count = constraints
        .iter()
        .filter(|c| c.relation == Relation::LessEqual)
        .count();
    let surplus_count = constraints
        .iter()
        .filter(|c| c.relation == Relation::GreaterEqual)
        .count();
    let artificial_count = constraints
        .iter()
        .filter(|c| matches!(c.relation, Relation::GreaterEqual | Relation::Equal))
        .count();

    let total_vars = n + slack_count + surplus_count + artificial_count;
    let m = constraints.len();

    let mut tableau = vec![vec![0.0; total_vars + 1]; m + 1];
    let mut basis = vec![0usize; m];

    let slack_offset = n;
    let surplus_offset = slack_offset + slack_count;
    let artificial_offset = surplus_offset + surplus_count;

    let mut slack_idx = 0usize;
    let mut surplus_idx = 0usize;
    let mut artificial_idx = 0usize;

    for (row, c) in constraints.iter().enumerate() {
        for j in 0..n {
            tableau[row][j] = c.coeffs[j];
        }
        tableau[row][total_vars] = c.rhs;

        match c.relation {
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
            return LpSolution {
                status,
                x: vec![0.0; n],
                objective: f64::NAN,
            };
        }
    }

    let phase1_obj = tableau[m][total_vars];
    if phase1_obj < -1e-8 {
        return LpSolution {
            status: LpStatus::Infeasible,
            x: vec![0.0; n],
            objective: f64::NAN,
        };
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

    let mut x = vec![0.0; n];
    for row in 0..m {
        let basic = basis[row];
        if basic < n {
            x[basic] = tableau[row][total_vars];
        }
    }

    LpSolution {
        status: phase2_status,
        x,
        objective: tableau[m][total_vars],
    }
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
