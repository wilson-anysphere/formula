use super::{
    clamp_vars, is_feasible, max_constraint_violation, objective_merit, Progress, SolveOptions,
    SolveOutcome, SolveStatus, SolverError, SolverModel, SolverProblem, VarType,
};

#[derive(Clone, Copy, Debug)]
pub struct EvolutionaryOptions {
    pub population_size: usize,
    pub elite_count: usize,
    pub mutation_rate: f64,
    pub crossover_rate: f64,
    pub penalty_weight: f64,
    pub seed: u64,
}

impl Default for EvolutionaryOptions {
    fn default() -> Self {
        Self {
            population_size: 40,
            elite_count: 4,
            mutation_rate: 0.2,
            crossover_rate: 0.7,
            penalty_weight: 50.0,
            seed: 0x5EED_5EED_1234_5678,
        }
    }
}

#[derive(Clone)]
struct Individual {
    vars: Vec<f64>,
    fitness: f64,
    objective: f64,
    violation: f64,
    feasible: bool,
}

pub(crate) fn solve_evolutionary<M: SolverModel>(
    model: &mut M,
    problem: &SolverProblem,
    options: &mut SolveOptions<'_>,
) -> Result<SolveOutcome, SolverError> {
    let n = model.num_vars();
    let m = model.num_constraints();

    let mut rng = XorShift64::new(options.evolutionary.seed);

    let mut current: Vec<f64> = Vec::new();
    if current.try_reserve_exact(n).is_err() {
        debug_assert!(false, "solver allocation failed (vars={n})");
        return Err(SolverError::new("allocation failed"));
    }
    current.resize(n, 0.0);
    model.get_vars(&mut current);
    clamp_vars(&mut current, &problem.variables);

    let mut constraint_values: Vec<f64> = Vec::new();
    if constraint_values.try_reserve_exact(m).is_err() {
        debug_assert!(false, "solver allocation failed (constraints={m})");
        return Err(SolverError::new("allocation failed"));
    }
    constraint_values.resize(m, 0.0);

    let mut population: Vec<Individual> = Vec::new();
    if population
        .try_reserve_exact(options.evolutionary.population_size)
        .is_err()
    {
        debug_assert!(
            false,
            "solver allocation failed (population_size={})",
            options.evolutionary.population_size
        );
        return Err(SolverError::new("allocation failed"));
    }
    population.push(evaluate_individual(
        model,
        problem,
        &current,
        &mut constraint_values,
        options.evolutionary.penalty_weight,
        options.tolerance,
    )?);

    while population.len() < options.evolutionary.population_size {
        let vars = random_vars(&mut rng, &current, problem)?;
        population.push(evaluate_individual(
            model,
            problem,
            &vars,
            &mut constraint_values,
            options.evolutionary.penalty_weight,
            options.tolerance,
        )?);
    }

    let mut best_vars = current.clone();
    let mut best_obj = f64::NAN;
    let mut best_violation = f64::INFINITY;

    for ind in &population {
        if ind.feasible {
            if best_obj.is_nan() || better_objective(problem, ind.objective, best_obj) {
                best_obj = ind.objective;
                best_vars.clone_from(&ind.vars);
                best_violation = ind.violation;
            }
        }
    }

    let mut generations = 0usize;

    for gen in 0..options.max_iterations {
        generations = gen;

        population.sort_by(|a, b| {
            let a_fit = if a.fitness.is_nan() {
                f64::NEG_INFINITY
            } else {
                a.fitness
            };
            let b_fit = if b.fitness.is_nan() {
                f64::NEG_INFINITY
            } else {
                b.fitness
            };
            b_fit.total_cmp(&a_fit)
        });

        if let Some(progress) = options.progress.as_deref_mut() {
            let best_fit = &population[0];
            if !progress(Progress {
                iteration: gen,
                best_objective: best_obj,
                current_objective: best_fit.objective,
                max_constraint_violation: best_fit.violation,
            }) {
                return Ok(SolveOutcome {
                    status: SolveStatus::Cancelled,
                    iterations: generations,
                    original_vars: Vec::new(),
                    best_vars,
                    best_objective: best_obj,
                    max_constraint_violation: best_violation,
                });
            }
        }

        // Track best feasible.
        for ind in &population {
            if ind.feasible
                && (best_obj.is_nan() || better_objective(problem, ind.objective, best_obj))
            {
                best_obj = ind.objective;
                best_vars.clone_from(&ind.vars);
                best_violation = ind.violation;

                if problem.objective.kind == super::ObjectiveKind::Target
                    && (best_obj - problem.objective.target_value).abs()
                        <= problem.objective.target_tolerance.max(options.tolerance)
                {
                    generations = gen;
                    break;
                }
            }
        }

        // Next generation.
        let elite = options.evolutionary.elite_count.min(population.len());
        let mut next: Vec<Individual> = Vec::new();
        if next.try_reserve_exact(population.len()).is_err() {
            debug_assert!(false, "solver allocation failed (population={})", population.len());
            return Err(SolverError::new("allocation failed"));
        }
        next.extend_from_slice(&population[..elite]);

        while next.len() < options.evolutionary.population_size {
            let parent_a = tournament(&mut rng, &population);
            let parent_b = tournament(&mut rng, &population);

            let mut child_vars = if rng.next_f64() < options.evolutionary.crossover_rate {
                crossover(&mut rng, &parent_a.vars, &parent_b.vars)?
            } else {
                try_clone_f64_slice(&parent_a.vars)?
            };

            mutate(
                &mut rng,
                &mut child_vars,
                &current,
                problem,
                options.evolutionary.mutation_rate,
            );

            let child = evaluate_individual(
                model,
                problem,
                &child_vars,
                &mut constraint_values,
                options.evolutionary.penalty_weight,
                options.tolerance,
            )?;
            next.push(child);
        }

        population = next;
    }

    let status = if best_obj.is_nan() {
        SolveStatus::IterationLimit
    } else if problem.objective.kind == super::ObjectiveKind::Target
        && (best_obj - problem.objective.target_value).abs()
            <= problem.objective.target_tolerance.max(options.tolerance)
    {
        SolveStatus::Optimal
    } else {
        SolveStatus::Feasible
    };

    Ok(SolveOutcome {
        status,
        iterations: generations + 1,
        original_vars: Vec::new(),
        best_vars,
        best_objective: best_obj,
        max_constraint_violation: best_violation,
    })
}

fn evaluate_individual<M: SolverModel>(
    model: &mut M,
    problem: &SolverProblem,
    vars: &[f64],
    constraint_values: &mut [f64],
    penalty_weight: f64,
    feasibility_tol: f64,
) -> Result<Individual, SolverError> {
    model.set_vars(vars)?;
    model.recalc()?;
    let obj = model.objective();
    model.constraints(constraint_values);

    let violation = max_constraint_violation(constraint_values, &problem.constraints);
    let feasible = is_feasible(constraint_values, &problem.constraints, feasibility_tol);

    let mut merit = objective_merit(&problem.objective, obj);
    for c in &problem.constraints {
        let v = super::constraint_violation(constraint_values[c.index], c);
        merit += penalty_weight * v * v;
    }

    let fitness = -merit;

    let mut vars_vec: Vec<f64> = Vec::new();
    if vars_vec.try_reserve_exact(vars.len()).is_err() {
        debug_assert!(false, "solver allocation failed (vars={})", vars.len());
        return Err(SolverError::new("allocation failed"));
    }
    vars_vec.extend_from_slice(vars);

    Ok(Individual {
        vars: vars_vec,
        fitness,
        objective: obj,
        violation,
        feasible,
    })
}

fn better_objective(problem: &SolverProblem, a: f64, b: f64) -> bool {
    match problem.objective.kind {
        super::ObjectiveKind::Maximize => a > b,
        super::ObjectiveKind::Minimize => a < b,
        super::ObjectiveKind::Target => {
            (a - problem.objective.target_value).abs() < (b - problem.objective.target_value).abs()
        }
    }
}

fn random_vars(
    rng: &mut XorShift64,
    base: &[f64],
    problem: &SolverProblem,
) -> Result<Vec<f64>, SolverError> {
    let mut vars: Vec<f64> = Vec::new();
    if vars.try_reserve_exact(problem.variables.len()).is_err() {
        debug_assert!(
            false,
            "solver allocation failed (vars={})",
            problem.variables.len()
        );
        return Err(SolverError::new("allocation failed"));
    }
    for (j, spec) in problem.variables.iter().enumerate() {
        let (lo, hi) = finite_range(spec.lower, spec.upper, base[j]);
        let v = match spec.var_type {
            VarType::Binary => {
                if rng.next_f64() < 0.5 {
                    0.0
                } else {
                    1.0
                }
            }
            VarType::Integer => {
                let lo_i = lo.ceil();
                let hi_i = hi.floor();
                if hi_i < lo_i {
                    lo_i
                } else {
                    let span = (hi_i - lo_i + 1.0) as u64;
                    lo_i + (rng.next_u64() % span) as f64
                }
            }
            VarType::Continuous => lo + rng.next_f64() * (hi - lo),
        };
        vars.push(v);
    }
    clamp_vars(&mut vars, &problem.variables);
    Ok(vars)
}

fn finite_range(lower: f64, upper: f64, center: f64) -> (f64, f64) {
    let mut lo = if lower.is_finite() {
        lower
    } else {
        center - 10.0
    };
    let mut hi = if upper.is_finite() {
        upper
    } else {
        center + 10.0
    };
    if lo > hi {
        std::mem::swap(&mut lo, &mut hi);
    }
    if (hi - lo).abs() < 1e-9 {
        hi = lo + 1.0;
    }
    (lo, hi)
}

fn tournament<'a>(rng: &mut XorShift64, population: &'a [Individual]) -> &'a Individual {
    let k = 3usize.min(population.len());
    let mut best = &population[rng.next_usize(population.len())];
    for _ in 1..k {
        let cand = &population[rng.next_usize(population.len())];
        if cand.fitness > best.fitness {
            best = cand;
        }
    }
    best
}

fn try_clone_f64_slice(src: &[f64]) -> Result<Vec<f64>, SolverError> {
    let mut out: Vec<f64> = Vec::new();
    if out.try_reserve_exact(src.len()).is_err() {
        debug_assert!(false, "solver allocation failed (vars={})", src.len());
        return Err(SolverError::new("allocation failed"));
    }
    out.extend_from_slice(src);
    Ok(out)
}

fn crossover(rng: &mut XorShift64, a: &[f64], b: &[f64]) -> Result<Vec<f64>, SolverError> {
    debug_assert_eq!(a.len(), b.len());
    let mut child: Vec<f64> = Vec::new();
    if child.try_reserve_exact(a.len()).is_err() {
        debug_assert!(false, "solver allocation failed (vars={})", a.len());
        return Err(SolverError::new("allocation failed"));
    }
    for (x, y) in a.iter().zip(b.iter()) {
        child.push(if rng.next_f64() < 0.5 { *x } else { *y });
    }
    Ok(child)
}

fn mutate(
    rng: &mut XorShift64,
    vars: &mut [f64],
    base: &[f64],
    problem: &SolverProblem,
    rate: f64,
) {
    for (j, spec) in problem.variables.iter().enumerate() {
        if rng.next_f64() > rate {
            continue;
        }
        match spec.var_type {
            VarType::Binary => {
                vars[j] = if vars[j] >= 0.5 { 0.0 } else { 1.0 };
            }
            VarType::Integer => {
                let (lo, hi) = finite_range(spec.lower, spec.upper, base[j]);
                let step = (rng.next_u64() % 5) as i64 - 2; // [-2,2]
                let candidate = vars[j] + step as f64;
                vars[j] = candidate.max(lo.ceil()).min(hi.floor());
            }
            VarType::Continuous => {
                let (lo, hi) = finite_range(spec.lower, spec.upper, base[j]);
                let sigma = 0.1 * (hi - lo);
                let noise = rng.next_gaussian() * sigma;
                vars[j] = (vars[j] + noise).max(lo).min(hi);
            }
        }
    }
    clamp_vars(vars, &problem.variables);
}

struct XorShift64 {
    state: u64,
}

impl XorShift64 {
    fn new(seed: u64) -> Self {
        Self { state: seed.max(1) }
    }

    fn next_u64(&mut self) -> u64 {
        let mut x = self.state;
        x ^= x << 13;
        x ^= x >> 7;
        x ^= x << 17;
        self.state = x;
        x
    }

    fn next_f64(&mut self) -> f64 {
        // 53-bit precision float in [0, 1).
        let v = self.next_u64() >> 11;
        (v as f64) * (1.0 / ((1u64 << 53) as f64))
    }

    fn next_usize(&mut self, upper: usize) -> usize {
        if upper == 0 {
            return 0;
        }
        (self.next_u64() % (upper as u64)) as usize
    }

    fn next_gaussian(&mut self) -> f64 {
        // Box–Muller transform.
        let u1 = self.next_f64().max(1e-12);
        let u2 = self.next_f64();
        (-2.0 * u1.ln()).sqrt() * (2.0 * std::f64::consts::PI * u2).cos()
    }
}
