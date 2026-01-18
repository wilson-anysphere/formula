use super::{
    clamp_vars, is_feasible, max_constraint_violation, objective_merit, Progress, Relation,
    SolveOptions, SolveOutcome, SolveStatus, SolverError, SolverModel, SolverProblem, VarType,
};

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

#[derive(Clone, Copy, Debug)]
pub struct GrgOptions {
    /// Initial step size for the line search.
    pub initial_step: f64,
    /// Finite difference base step (scaled by variable magnitude).
    pub diff_step: f64,
    /// Penalty weight applied to squared constraint violations.
    pub penalty_weight: f64,
    /// Factor by which the penalty weight grows when constraints remain violated.
    pub penalty_growth: f64,
    /// Backtracking shrink factor (0 < shrink < 1).
    pub line_search_shrink: f64,
    /// Maximum backtracking steps per iteration.
    pub line_search_max_steps: usize,
}

impl Default for GrgOptions {
    fn default() -> Self {
        Self {
            initial_step: 1.0,
            diff_step: 1e-5,
            penalty_weight: 10.0,
            penalty_growth: 2.0,
            line_search_shrink: 0.5,
            line_search_max_steps: 20,
        }
    }
}

pub(crate) fn solve_grg<M: SolverModel>(
    model: &mut M,
    problem: &SolverProblem,
    options: &mut SolveOptions<'_>,
) -> Result<SolveOutcome, SolverError> {
    let n = model.num_vars();
    let m = model.num_constraints();

    let mut x = try_zeros_f64(n)?;
    model.get_vars(&mut x);
    clamp_vars(&mut x, &problem.variables);

    let mut constraint_values = try_zeros_f64(m)?;
    let mut best_vars = try_clone_f64_slice(&x)?;
    let mut best_obj = f64::NAN;
    let mut best_violation = f64::INFINITY;

    evaluate(model, &x, &mut constraint_values)?;
    let mut current_obj = model.objective();
    let mut current_violation = max_constraint_violation(&constraint_values, &problem.constraints);
    let mut penalty_weight = options.grg.penalty_weight;

    let mut best_overall_vars = try_clone_f64_slice(&x)?;
    let mut best_overall_obj = current_obj;
    let mut best_overall_violation = current_violation;
    let mut best_overall_merit =
        merit_function(problem, current_obj, &constraint_values, penalty_weight);
    let mut best_violation_seen = current_violation;

    if is_feasible(&constraint_values, &problem.constraints, options.tolerance) {
        best_vars.clone_from(&x);
        best_obj = current_obj;
        best_violation = current_violation;
    }

    let mut direction = try_zeros_f64(n)?;
    let mut candidate = try_zeros_f64(n)?;
    let mut candidate_constraints = try_zeros_f64(m)?;

    let mut iterations = 0usize;
    for iter in 0..options.max_iterations {
        iterations = iter;

        let merit = merit_function(problem, current_obj, &constraint_values, penalty_weight);

        let (grad, grad_norm) = finite_difference_gradient(
            model,
            problem,
            &x,
            &constraint_values,
            penalty_weight,
            options.grg.diff_step,
        )?;

        if grad_norm <= options.tolerance
            && is_feasible(&constraint_values, &problem.constraints, options.tolerance)
        {
            break;
        }

        if let Some(progress) = options.progress.as_deref_mut() {
            if !progress(Progress {
                iteration: iter,
                best_objective: best_obj,
                current_objective: current_obj,
                max_constraint_violation: current_violation,
            }) {
                return Ok(SolveOutcome {
                    status: SolveStatus::Cancelled,
                    iterations,
                    original_vars: Vec::new(),
                    best_vars,
                    best_objective: best_obj,
                    max_constraint_violation: best_violation,
                });
            }
        }

        // Normalized steepest descent direction; normalization avoids line search
        // degenerate behavior when the penalty weight becomes large.
        direction.fill(0.0);
        if grad_norm > 0.0 {
            for j in 0..n {
                direction[j] = -grad[j] / grad_norm;
            }
        }
        let dir_derivative: f64 = grad.iter().zip(direction.iter()).map(|(g, d)| g * d).sum();

        // Line search.
        let mut step = options.grg.initial_step;
        let mut accepted = false;
        candidate.clone_from(&x);
        let mut candidate_obj = current_obj;

        for _ in 0..options.grg.line_search_max_steps {
            for j in 0..n {
                candidate[j] = x[j] + step * direction[j];
            }
            clamp_vars(&mut candidate, &problem.variables);

            evaluate(model, &candidate, &mut candidate_constraints)?;
            candidate_obj = model.objective();

            let candidate_merit = merit_function(
                problem,
                candidate_obj,
                &candidate_constraints,
                penalty_weight,
            );

            // Armijo condition: merit(x + a d) <= merit(x) + c1 * a * grad·d
            // where grad·d is negative for a descent direction.
            if candidate_merit <= merit + 1e-4 * step * dir_derivative {
                accepted = true;
                break;
            }

            step *= options.grg.line_search_shrink;
        }

        if !accepted {
            // If we're stuck and still infeasible, crank penalties and try again.
            if current_violation > options.tolerance {
                penalty_weight *= options.grg.penalty_growth;
                continue;
            }
            break;
        }

        x.clone_from(&candidate);
        constraint_values.clone_from(&candidate_constraints);
        current_obj = candidate_obj;
        current_violation = max_constraint_violation(&constraint_values, &problem.constraints);

        // Feasibility restoration pass for equality constraints.
        if current_violation > options.tolerance
            && problem
                .constraints
                .iter()
                .any(|c| c.relation == Relation::Equal)
        {
            repair_equalities(
                model,
                problem,
                &mut x,
                &mut constraint_values,
                options.grg.diff_step,
                options.tolerance,
            )?;
            current_obj = model.objective();
            current_violation = max_constraint_violation(&constraint_values, &problem.constraints);
        }
        if current_violation < best_violation_seen {
            best_violation_seen = current_violation;
        }

        let current_merit =
            merit_function(problem, current_obj, &constraint_values, penalty_weight);
        if current_merit < best_overall_merit {
            best_overall_merit = current_merit;
            best_overall_vars.clone_from(&x);
            best_overall_obj = current_obj;
            best_overall_violation = current_violation;
        }

        let feasible = is_feasible(&constraint_values, &problem.constraints, options.tolerance);
        if feasible {
            let improved = match problem.objective.kind {
                super::ObjectiveKind::Maximize => best_obj.is_nan() || current_obj > best_obj,
                super::ObjectiveKind::Minimize => best_obj.is_nan() || current_obj < best_obj,
                super::ObjectiveKind::Target => {
                    let best_dist = if best_obj.is_nan() {
                        f64::INFINITY
                    } else {
                        (best_obj - problem.objective.target_value).abs()
                    };
                    let cur_dist = (current_obj - problem.objective.target_value).abs();
                    cur_dist < best_dist
                }
            };

            if improved {
                best_vars.clone_from(&x);
                best_obj = current_obj;
                best_violation = current_violation;

                if problem.objective.kind == super::ObjectiveKind::Target
                    && (best_obj - problem.objective.target_value).abs()
                        <= problem.objective.target_tolerance.max(options.tolerance)
                {
                    break;
                }
            }
        } else {
            // If we're not making progress towards feasibility, increase penalties.
            if current_violation > 10.0 * options.tolerance
                && current_violation >= 0.99 * best_violation_seen
            {
                penalty_weight = (penalty_weight * options.grg.penalty_growth).min(1e12);
            }
        }
    }

    let status = if best_obj.is_nan() {
        SolveStatus::IterationLimit
    } else if problem.objective.kind == super::ObjectiveKind::Target
        && (best_obj - problem.objective.target_value).abs()
            <= problem.objective.target_tolerance.max(options.tolerance)
    {
        SolveStatus::Optimal
    } else if iterations + 1 >= options.max_iterations {
        SolveStatus::IterationLimit
    } else {
        SolveStatus::Feasible
    };

    let (best_vars, best_obj, best_violation) = if best_obj.is_nan() {
        (best_overall_vars, best_overall_obj, best_overall_violation)
    } else {
        (best_vars, best_obj, best_violation)
    };

    Ok(SolveOutcome {
        status,
        iterations: iterations + 1,
        original_vars: Vec::new(),
        best_vars,
        best_objective: best_obj,
        max_constraint_violation: best_violation,
    })
}

fn evaluate<M: SolverModel>(
    model: &mut M,
    vars: &[f64],
    constraint_values: &mut [f64],
) -> Result<(), SolverError> {
    model.set_vars(vars)?;
    model.recalc()?;
    model.constraints(constraint_values);
    Ok(())
}

fn merit_function(
    problem: &SolverProblem,
    objective_value: f64,
    constraint_values: &[f64],
    penalty_weight: f64,
) -> f64 {
    let mut merit = objective_merit(&problem.objective, objective_value);
    for c in &problem.constraints {
        let v = super::constraint_violation(constraint_values[c.index], c);
        merit += penalty_weight * v * v;
    }
    merit
}

fn finite_difference_gradient<M: SolverModel>(
    model: &mut M,
    problem: &SolverProblem,
    x: &[f64],
    current_constraints: &[f64],
    penalty_weight: f64,
    diff_step: f64,
) -> Result<(Vec<f64>, f64), SolverError> {
    let n = x.len();
    let mut grad = try_zeros_f64(n)?;
    let base_obj = model.objective();
    let base_merit = merit_function(problem, base_obj, current_constraints, penalty_weight);

    let mut tmp_constraints = try_zeros_f64(current_constraints.len())?;
    let mut x_fwd = try_clone_f64_slice(x)?;
    let mut x_bwd = try_clone_f64_slice(x)?;

    for j in 0..n {
        if problem.variables[j].var_type != VarType::Continuous {
            grad[j] = 0.0;
            continue;
        }

        let scale = 1.0_f64.max(x[j].abs());
        let h = diff_step * scale;
        x_fwd.as_mut_slice().copy_from_slice(x);
        x_fwd[j] += h;
        clamp_vars(&mut x_fwd, &problem.variables);

        evaluate(model, &x_fwd, &mut tmp_constraints)?;
        let obj_fwd = model.objective();
        let merit_fwd = merit_function(problem, obj_fwd, &tmp_constraints, penalty_weight);

        // Try central difference if we can move backwards.
        x_bwd.as_mut_slice().copy_from_slice(x);
        x_bwd[j] -= h;
        clamp_vars(&mut x_bwd, &problem.variables);

        if (x_bwd[j] - x[j]).abs() > 1e-12 {
            evaluate(model, &x_bwd, &mut tmp_constraints)?;
            let obj_bwd = model.objective();
            let merit_bwd = merit_function(problem, obj_bwd, &tmp_constraints, penalty_weight);

            let denom = x_fwd[j] - x_bwd[j];
            grad[j] = (merit_fwd - merit_bwd) / denom;
        } else {
            let denom = x_fwd[j] - x[j];
            grad[j] = (merit_fwd - base_merit) / denom;
        }
    }

    let norm = grad.iter().map(|v| v * v).sum::<f64>().sqrt();
    Ok((grad, norm))
}

fn repair_equalities<M: SolverModel>(
    model: &mut M,
    problem: &SolverProblem,
    x: &mut [f64],
    constraint_values: &mut [f64],
    diff_step: f64,
    tol: f64,
) -> Result<(), SolverError> {
    let n = x.len();
    let m = constraint_values.len();
    let max_iters = 10;
    let mut tmp_constraints = try_zeros_f64(m)?;
    let mut grad = try_zeros_f64(n)?;
    let mut x_fwd = try_zeros_f64(n)?;

    for _ in 0..max_iters {
        evaluate(model, x, constraint_values)?;
        if max_constraint_violation(constraint_values, &problem.constraints) <= tol {
            return Ok(());
        }

        for c in problem
            .constraints
            .iter()
            .filter(|c| c.relation == Relation::Equal)
        {
            let lhs = constraint_values[c.index];
            let residual = lhs - c.rhs;
            let eq_tol = c.tolerance.max(tol);
            if residual.abs() <= eq_tol {
                continue;
            }

            // Estimate gradient of the equality constraint.
            grad.fill(0.0);
            for j in 0..n {
                if problem.variables[j].var_type != VarType::Continuous {
                    continue;
                }
                let scale = 1.0_f64.max(x[j].abs());
                let h = diff_step * scale;
                x_fwd.as_mut_slice().copy_from_slice(x);
                x_fwd[j] += h;
                clamp_vars(&mut x_fwd, &problem.variables);

                evaluate(model, &x_fwd, &mut tmp_constraints)?;
                let lhs_fwd = tmp_constraints[c.index];
                let denom = x_fwd[j] - x[j];
                if denom.abs() > 1e-12 {
                    grad[j] = (lhs_fwd - lhs) / denom;
                }
            }

            let norm_sq = grad.iter().map(|v| v * v).sum::<f64>();
            if norm_sq <= 1e-12 {
                continue;
            }

            // Project residual out in the direction of the constraint gradient.
            for j in 0..n {
                x[j] -= residual * grad[j] / norm_sq;
            }
            clamp_vars(x, &problem.variables);
        }
    }

    evaluate(model, x, constraint_values)?;
    Ok(())
}
