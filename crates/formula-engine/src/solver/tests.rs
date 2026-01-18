use super::*;
use crate::Engine;

struct FnModel<F>
where
    F: Fn(&[f64]) -> (f64, Vec<f64>),
{
    vars: Vec<f64>,
    objective: f64,
    constraints: Vec<f64>,
    f: F,
}

impl<F> FnModel<F>
where
    F: Fn(&[f64]) -> (f64, Vec<f64>),
{
    fn new(vars: Vec<f64>, f: F) -> Self {
        let (objective, constraints) = f(&vars);
        Self {
            vars,
            objective,
            constraints,
            f,
        }
    }
}

impl<F> SolverModel for FnModel<F>
where
    F: Fn(&[f64]) -> (f64, Vec<f64>),
{
    fn num_vars(&self) -> usize {
        self.vars.len()
    }

    fn num_constraints(&self) -> usize {
        self.constraints.len()
    }

    fn get_vars(&self, out: &mut [f64]) {
        out.copy_from_slice(&self.vars);
    }

    fn set_vars(&mut self, vars: &[f64]) -> Result<(), SolverError> {
        if vars.len() != self.vars.len() {
            return Err(SolverError::new("wrong var length"));
        }
        self.vars.copy_from_slice(vars);
        Ok(())
    }

    fn recalc(&mut self) -> Result<(), SolverError> {
        let (objective, constraints) = (self.f)(&self.vars);
        self.objective = objective;
        self.constraints = constraints;
        Ok(())
    }

    fn objective(&self) -> f64 {
        self.objective
    }

    fn constraints(&self, out: &mut [f64]) {
        out.copy_from_slice(&self.constraints);
    }
}

#[test]
fn simplex_solves_linear_lp() {
    // Maximize 3x + 2y
    // s.t. x + y <= 4
    //      x <= 2
    //      y <= 3
    //      x,y >= 0
    let model = FnModel::new(vec![0.0, 0.0], |vars| {
        let x = vars[0];
        let y = vars[1];
        let objective = 3.0 * x + 2.0 * y;
        let constraints = vec![x + y, x, y];
        (objective, constraints)
    });

    let mut model = model;
    let problem = SolverProblem {
        objective: Objective::maximize(),
        variables: vec![
            VarSpec::continuous(0.0, f64::INFINITY),
            VarSpec::continuous(0.0, f64::INFINITY),
        ],
        constraints: vec![
            Constraint::new(0, Relation::LessEqual, 4.0),
            Constraint::new(1, Relation::LessEqual, 2.0),
            Constraint::new(2, Relation::LessEqual, 3.0),
        ],
    };

    let mut options = SolveOptions::default();
    options.method = SolveMethod::Simplex;
    options.max_iterations = 100;
    options.tolerance = 1e-8;
    options.apply_solution = false;

    let outcome = Solver::solve(&mut model, &problem, options).expect("solve");
    assert!(matches!(
        outcome.status,
        SolveStatus::Optimal | SolveStatus::Feasible
    ));
    assert!(
        (outcome.best_vars[0] - 2.0).abs() < 1e-6,
        "x={}",
        outcome.best_vars[0]
    );
    assert!(
        (outcome.best_vars[1] - 2.0).abs() < 1e-6,
        "y={}",
        outcome.best_vars[1]
    );
    assert!((outcome.best_objective - 10.0).abs() < 1e-6);
    assert!(outcome.max_constraint_violation < 1e-6);
}

#[test]
fn simplex_hits_target_objective_value() {
    // Find x such that objective=x hits 5.0.
    let model = FnModel::new(vec![0.0], |vars| (vars[0], Vec::new()));
    let mut model = model;

    let problem = SolverProblem {
        objective: Objective::target(5.0, 1e-8),
        variables: vec![VarSpec::continuous(0.0, 10.0)],
        constraints: Vec::new(),
    };

    let mut options = SolveOptions::default();
    options.method = SolveMethod::Simplex;
    options.apply_solution = false;

    let outcome = Solver::solve(&mut model, &problem, options).expect("solve");
    assert_eq!(outcome.status, SolveStatus::Optimal);
    assert!((outcome.best_vars[0] - 5.0).abs() < 1e-6);
    assert!((outcome.best_objective - 5.0).abs() < 1e-6);
}

#[test]
fn simplex_solves_integer_program_via_branch_and_bound() {
    // Maximize x + y
    // s.t. 2x + y <= 4
    //      x <= 3
    //      y <= 3
    //      x,y integer >= 0
    //
    // LP relaxation chooses x=0.5, y=3 (objective=3.5). Integer optimum is 3.
    let model = FnModel::new(vec![0.0, 0.0], |vars| {
        let x = vars[0];
        let y = vars[1];
        let objective = x + y;
        let constraints = vec![2.0 * x + y, x, y];
        (objective, constraints)
    });

    let mut model = model;
    let problem = SolverProblem {
        objective: Objective::maximize(),
        variables: vec![VarSpec::integer(0.0, 3.0), VarSpec::integer(0.0, 3.0)],
        constraints: vec![
            Constraint::new(0, Relation::LessEqual, 4.0),
            Constraint::new(1, Relation::LessEqual, 3.0),
            Constraint::new(2, Relation::LessEqual, 3.0),
        ],
    };

    let mut options = SolveOptions::default();
    options.method = SolveMethod::Simplex;
    options.apply_solution = false;

    let outcome = Solver::solve(&mut model, &problem, options).expect("solve");
    assert!(matches!(
        outcome.status,
        SolveStatus::Optimal | SolveStatus::Feasible
    ));
    assert!((outcome.best_objective - 3.0).abs() < 1e-6);
    assert!((outcome.best_vars[0].fract()).abs() < 1e-6);
    assert!((outcome.best_vars[1].fract()).abs() < 1e-6);
    assert!(outcome.max_constraint_violation < 1e-6);
}

#[test]
fn simplex_can_be_cancelled_via_progress_callback() {
    let model = FnModel::new(vec![0.0], |vars| (vars[0], Vec::new()));
    let mut model = model;

    let problem = SolverProblem {
        objective: Objective::maximize(),
        variables: vec![VarSpec::continuous(0.0, 10.0)],
        constraints: Vec::new(),
    };

    let mut progress_calls = 0usize;
    let mut progress_cb = |_p: Progress| {
        progress_calls += 1;
        false
    };

    let mut options = SolveOptions::default();
    options.method = SolveMethod::Simplex;
    options.apply_solution = false;
    options.progress = Some(&mut progress_cb);

    let outcome = Solver::solve(&mut model, &problem, options).expect("solve");
    assert_eq!(outcome.status, SolveStatus::Cancelled);
    assert!(progress_calls >= 1);
    assert!(!outcome.best_vars.is_empty());
}

#[test]
fn grg_solves_nonlinear_constrained_problem() {
    // Minimize (x-1)^2 + (y-2)^2 subject to x + y = 3, 0<=x,y<=3.
    let model = FnModel::new(vec![0.5, 0.5], |vars| {
        let x = vars[0];
        let y = vars[1];
        let objective = (x - 1.0).powi(2) + (y - 2.0).powi(2);
        let constraints = vec![x + y];
        (objective, constraints)
    });

    let mut model = model;
    let problem = SolverProblem {
        objective: Objective::minimize(),
        variables: vec![VarSpec::continuous(0.0, 3.0), VarSpec::continuous(0.0, 3.0)],
        constraints: vec![Constraint::new(0, Relation::Equal, 3.0).with_tolerance(1e-6)],
    };

    let mut options = SolveOptions::default();
    options.method = SolveMethod::GrgNonlinear;
    options.max_iterations = 250;
    options.tolerance = 1e-6;
    options.apply_solution = false;
    options.grg.initial_step = 1.0;
    options.grg.penalty_weight = 100.0;

    let outcome = Solver::solve(&mut model, &problem, options).expect("solve");
    assert!(
        matches!(outcome.status, SolveStatus::Optimal | SolveStatus::Feasible),
        "{:?}",
        outcome.status
    );
    assert!(
        (outcome.best_vars[0] - 1.0).abs() < 1e-2,
        "x={}",
        outcome.best_vars[0]
    );
    assert!(
        (outcome.best_vars[1] - 2.0).abs() < 1e-2,
        "y={}",
        outcome.best_vars[1]
    );
    assert!(outcome.max_constraint_violation < 1e-3);
    assert!(outcome.best_objective < 1e-3);
}

#[test]
fn evolutionary_handles_nonsmooth_integer_problem() {
    // Minimize |x| with x integer in [-5,5].
    let model = FnModel::new(vec![5.0], |vars| {
        let x = vars[0];
        let objective = x.abs();
        (objective, Vec::new())
    });

    let mut model = model;
    let problem = SolverProblem {
        objective: Objective::minimize(),
        variables: vec![VarSpec::integer(-5.0, 5.0)],
        constraints: Vec::new(),
    };

    let mut options = SolveOptions::default();
    options.method = SolveMethod::Evolutionary;
    options.max_iterations = 80;
    options.tolerance = 1e-8;
    options.apply_solution = false;
    options.evolutionary.population_size = 30;
    options.evolutionary.elite_count = 3;
    options.evolutionary.seed = 123;

    let outcome = Solver::solve(&mut model, &problem, options).expect("solve");
    assert!(matches!(
        outcome.status,
        SolveStatus::Optimal | SolveStatus::Feasible
    ));
    assert_eq!(outcome.best_vars[0], 0.0);
    assert_eq!(outcome.best_objective, 0.0);
}

#[test]
fn evolutionary_handles_nan_objectives_without_panicking() {
    let model = FnModel::new(vec![0.0], |_vars| (f64::NAN, Vec::new()));

    let mut model = model;
    let problem = SolverProblem {
        objective: Objective::minimize(),
        variables: vec![VarSpec::continuous(-1.0, 1.0)],
        constraints: Vec::new(),
    };

    let mut options = SolveOptions::default();
    options.method = SolveMethod::Evolutionary;
    options.max_iterations = 10;
    options.tolerance = 1e-8;
    options.apply_solution = false;
    options.evolutionary.population_size = 20;
    options.evolutionary.elite_count = 2;
    options.evolutionary.seed = 123;

    let outcome = Solver::solve(&mut model, &problem, options).expect("solve");
    assert!(matches!(outcome.status, SolveStatus::IterationLimit));
    assert!(outcome.best_objective.is_nan());
}

#[test]
fn simplex_integrates_with_engine_recalc() {
    // Same LP as `simplex_solves_linear_lp`, but evaluated through the real
    // `Engine` formula + dependency graph.
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 0.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 0.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", "=3*A1+2*A2")
        .unwrap();
    engine.set_cell_formula("Sheet1", "C1", "=A1+A2").unwrap();
    engine.set_cell_formula("Sheet1", "C2", "=A1").unwrap();
    engine.set_cell_formula("Sheet1", "C3", "=A2").unwrap();
    engine.recalculate();

    let mut model = EngineSolverModel::new(
        &mut engine,
        "Sheet1",
        "B1",
        vec!["A1", "A2"],
        vec!["C1", "C2", "C3"],
    )
    .unwrap();

    let problem = SolverProblem {
        objective: Objective::maximize(),
        variables: vec![
            VarSpec::continuous(0.0, f64::INFINITY),
            VarSpec::continuous(0.0, f64::INFINITY),
        ],
        constraints: vec![
            Constraint::new(0, Relation::LessEqual, 4.0),
            Constraint::new(1, Relation::LessEqual, 2.0),
            Constraint::new(2, Relation::LessEqual, 3.0),
        ],
    };

    let mut options = SolveOptions::default();
    options.method = SolveMethod::Simplex;
    options.apply_solution = false;

    let outcome = Solver::solve(&mut model, &problem, options).expect("solve");
    assert!(matches!(
        outcome.status,
        SolveStatus::Optimal | SolveStatus::Feasible
    ));
    assert!((outcome.best_vars[0] - 2.0).abs() < 1e-6);
    assert!((outcome.best_vars[1] - 2.0).abs() < 1e-6);
    assert!((outcome.best_objective - 10.0).abs() < 1e-6);
    assert!(outcome.max_constraint_violation < 1e-6);
}

#[test]
fn grg_integrates_with_engine_recalc() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 0.5).unwrap();
    engine.set_cell_value("Sheet1", "A2", 0.5).unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", "=(A1-1)*(A1-1)+(A2-2)*(A2-2)")
        .unwrap();
    engine.set_cell_formula("Sheet1", "C1", "=A1+A2").unwrap();
    engine.recalculate();

    let mut model =
        EngineSolverModel::new(&mut engine, "Sheet1", "B1", vec!["A1", "A2"], vec!["C1"]).unwrap();

    let problem = SolverProblem {
        objective: Objective::minimize(),
        variables: vec![VarSpec::continuous(0.0, 3.0), VarSpec::continuous(0.0, 3.0)],
        constraints: vec![Constraint::new(0, Relation::Equal, 3.0).with_tolerance(1e-6)],
    };

    let mut options = SolveOptions::default();
    options.method = SolveMethod::GrgNonlinear;
    options.max_iterations = 300;
    options.tolerance = 1e-6;
    options.apply_solution = false;
    options.grg.penalty_weight = 100.0;

    let outcome = Solver::solve(&mut model, &problem, options).expect("solve");
    assert!(matches!(
        outcome.status,
        SolveStatus::Optimal | SolveStatus::Feasible
    ));
    assert!(
        (outcome.best_vars[0] - 1.0).abs() < 1e-2,
        "x={}",
        outcome.best_vars[0]
    );
    assert!(
        (outcome.best_vars[1] - 2.0).abs() < 1e-2,
        "y={}",
        outcome.best_vars[1]
    );
    assert!(outcome.max_constraint_violation < 1e-3);
    assert!(outcome.best_objective < 1e-3);
}
