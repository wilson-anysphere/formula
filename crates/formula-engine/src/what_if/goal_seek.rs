use crate::what_if::{CellRef, CellValue, WhatIfError, WhatIfModel};

/// Parameters for Goal Seek.
///
/// Mirrors the high-level design in `docs/07-power-features.md`.
#[derive(Clone, Debug)]
pub struct GoalSeekParams {
    /// Cell containing the formula we want to match.
    pub target_cell: CellRef,
    /// Desired output value for `target_cell`.
    pub target_value: f64,
    /// Cell to adjust while searching.
    pub changing_cell: CellRef,
    /// Maximum number of iterations to attempt.
    pub max_iterations: usize,
    /// Absolute tolerance on the target output.
    pub tolerance: f64,
    /// Derivative step used by finite differencing. If `None`, a value is
    /// chosen based on the current input (`abs(x)*0.001` or `0.001`).
    pub derivative_step: Option<f64>,
    /// Minimum absolute derivative before switching to bisection.
    pub min_derivative: f64,
    /// How aggressively to search for a sign-changing bracket when falling
    /// back to bisection.
    pub max_bracket_expansions: usize,
}

impl GoalSeekParams {
    pub fn new(
        target_cell: impl Into<CellRef>,
        target_value: f64,
        changing_cell: impl Into<CellRef>,
    ) -> Self {
        Self {
            target_cell: target_cell.into(),
            target_value,
            changing_cell: changing_cell.into(),
            max_iterations: 100,
            tolerance: 0.001,
            derivative_step: None,
            min_derivative: 1e-10,
            max_bracket_expansions: 50,
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum GoalSeekStatus {
    Converged,
    MaxIterationsReached,
    NoBracketFound,
    NumericalFailure,
}

#[derive(Clone, Debug)]
pub struct GoalSeekResult {
    pub status: GoalSeekStatus,
    pub solution: f64,
    pub iterations: usize,
    pub final_output: f64,
    pub final_error: f64,
}

impl GoalSeekResult {
    pub fn success(&self) -> bool {
        self.status == GoalSeekStatus::Converged
    }
}

#[derive(Clone, Copy, Debug)]
pub struct GoalSeekProgress {
    pub iteration: usize,
    pub input: f64,
    pub output: f64,
    pub error: f64,
}

pub struct GoalSeek;

impl GoalSeek {
    pub fn solve<M: WhatIfModel>(
        model: &mut M,
        params: GoalSeekParams,
    ) -> Result<GoalSeekResult, WhatIfError<M::Error>> {
        Self::solve_with_progress(model, params, |_| {})
    }

    pub fn solve_with_progress<M: WhatIfModel, F: FnMut(GoalSeekProgress)>(
        model: &mut M,
        params: GoalSeekParams,
        mut progress: F,
    ) -> Result<GoalSeekResult, WhatIfError<M::Error>> {
        if params.max_iterations == 0 {
            return Err(WhatIfError::InvalidParams("max_iterations must be > 0"));
        }
        if !(params.tolerance > 0.0) {
            return Err(WhatIfError::InvalidParams("tolerance must be > 0"));
        }
        if !(params.min_derivative > 0.0) {
            return Err(WhatIfError::InvalidParams("min_derivative must be > 0"));
        }

        // Ensure model outputs reflect the current state.
        model.recalculate()?;

        let mut current_input = get_number(model, &params.changing_cell)?;
        let mut current_output = get_number(model, &params.target_cell)?;
        let mut error = current_output - params.target_value;

        progress(GoalSeekProgress {
            iteration: 0,
            input: current_input,
            output: current_output,
            error,
        });

        if error.abs() < params.tolerance {
            return Ok(GoalSeekResult {
                status: GoalSeekStatus::Converged,
                solution: current_input,
                iterations: 0,
                final_output: current_output,
                final_error: error,
            });
        }

        for iter in 0..params.max_iterations {
            let delta = params
                .derivative_step
                .unwrap_or_else(|| (current_input.abs() * 0.001).max(0.001));

            let perturbed_output = eval_target(model, &params, current_input + delta)?;
            let derivative = (perturbed_output - current_output) / delta;

            if !derivative.is_finite() {
                return Ok(Self::bisection_fallback(
                    model,
                    &params,
                    iter,
                    current_input,
                    current_output,
                    error,
                    &mut progress,
                )?);
            }

            if derivative.abs() < params.min_derivative {
                return Ok(Self::bisection_fallback(
                    model,
                    &params,
                    iter,
                    current_input,
                    current_output,
                    error,
                    &mut progress,
                )?);
            }

            let next_input = current_input - error / derivative;
            if !next_input.is_finite() {
                return Ok(GoalSeekResult {
                    status: GoalSeekStatus::NumericalFailure,
                    solution: current_input,
                    iterations: iter,
                    final_output: current_output,
                    final_error: error,
                });
            }

            let next_output = eval_target(model, &params, next_input)?;
            let next_error = next_output - params.target_value;

            current_input = next_input;
            current_output = next_output;
            error = next_error;

            progress(GoalSeekProgress {
                iteration: iter + 1,
                input: current_input,
                output: current_output,
                error,
            });

            if error.abs() < params.tolerance {
                return Ok(GoalSeekResult {
                    status: GoalSeekStatus::Converged,
                    solution: current_input,
                    iterations: iter + 1,
                    final_output: current_output,
                    final_error: error,
                });
            }
        }

        Ok(GoalSeekResult {
            status: GoalSeekStatus::MaxIterationsReached,
            solution: current_input,
            iterations: params.max_iterations,
            final_output: current_output,
            final_error: error,
        })
    }

    fn bisection_fallback<M: WhatIfModel, F: FnMut(GoalSeekProgress)>(
        model: &mut M,
        params: &GoalSeekParams,
        iterations_used: usize,
        start_input: f64,
        start_output: f64,
        start_error: f64,
        progress: &mut F,
    ) -> Result<GoalSeekResult, WhatIfError<M::Error>> {
        // If we are already close enough, do not bother bracketing.
        if start_error.abs() < params.tolerance {
            return Ok(GoalSeekResult {
                status: GoalSeekStatus::Converged,
                solution: start_input,
                iterations: iterations_used,
                final_output: start_output,
                final_error: start_error,
            });
        }

        let Some((mut lo, mut hi, mut lo_err, mut hi_err)) =
            bracket_root(model, params, start_input, start_error)?
        else {
            // Can't safely bisect without a sign change.
            return Ok(GoalSeekResult {
                status: GoalSeekStatus::NoBracketFound,
                solution: start_input,
                iterations: iterations_used,
                final_output: start_output,
                final_error: start_error,
            });
        };

        let remaining = params.max_iterations.saturating_sub(iterations_used);
        let mut best_input = start_input;
        let mut best_output = start_output;
        let mut best_error = start_error;

        for i in 0..remaining {
            let mid = (lo + hi) * 0.5;
            let mid_output = eval_target(model, params, mid)?;
            let mid_err = mid_output - params.target_value;

            best_input = mid;
            best_output = mid_output;
            best_error = mid_err;

            progress(GoalSeekProgress {
                iteration: iterations_used + i + 1,
                input: best_input,
                output: best_output,
                error: best_error,
            });

            if best_error.abs() < params.tolerance {
                return Ok(GoalSeekResult {
                    status: GoalSeekStatus::Converged,
                    solution: best_input,
                    iterations: iterations_used + i + 1,
                    final_output: best_output,
                    final_error: best_error,
                });
            }

            if lo_err == 0.0 || hi_err == 0.0 {
                // Should have converged earlier, but protect against edge cases.
                break;
            }

            // Keep the sub-interval containing the sign change.
            if sign(mid_err) == sign(lo_err) {
                lo = mid;
                lo_err = mid_err;
            } else {
                hi = mid;
                hi_err = mid_err;
            }

            if (hi - lo).abs() <= f64::EPSILON * (lo.abs() + hi.abs() + 1.0) {
                // Interval collapse (numerical precision). Stop.
                break;
            }
        }

        Ok(GoalSeekResult {
            status: GoalSeekStatus::MaxIterationsReached,
            solution: best_input,
            iterations: params.max_iterations,
            final_output: best_output,
            final_error: best_error,
        })
    }
}

fn sign(value: f64) -> i8 {
    if value > 0.0 {
        1
    } else if value < 0.0 {
        -1
    } else {
        0
    }
}

fn get_number<M: WhatIfModel>(model: &M, cell: &CellRef) -> Result<f64, WhatIfError<M::Error>> {
    let value = model.get_cell_value(cell)?;
    value
        .as_number()
        .ok_or_else(|| WhatIfError::NonNumericCell {
            cell: cell.clone(),
            value,
        })
}

fn eval_target<M: WhatIfModel>(
    model: &mut M,
    params: &GoalSeekParams,
    input: f64,
) -> Result<f64, WhatIfError<M::Error>> {
    model.set_cell_value(&params.changing_cell, CellValue::Number(input))?;
    model.recalculate()?;
    get_number(model, &params.target_cell)
}

fn bracket_root<M: WhatIfModel>(
    model: &mut M,
    params: &GoalSeekParams,
    start_input: f64,
    start_error: f64,
) -> Result<Option<(f64, f64, f64, f64)>, WhatIfError<M::Error>> {
    // Pick an initial step scale. For tiny numbers we start at 1.0 to avoid a
    // pathological "zero-width" bracket search.
    let mut step = (start_input.abs() * 0.1).max(1.0);

    for _ in 0..params.max_bracket_expansions {
        let left = start_input - step;
        let left_output = eval_target(model, params, left)?;
        let left_error = left_output - params.target_value;
        if left_error.abs() < params.tolerance {
            return Ok(Some((left, left, left_error, left_error)));
        }
        if sign(left_error) != sign(start_error) {
            let (lo, hi, lo_err, hi_err) = if left < start_input {
                (left, start_input, left_error, start_error)
            } else {
                (start_input, left, start_error, left_error)
            };
            return Ok(Some((lo, hi, lo_err, hi_err)));
        }

        let right = start_input + step;
        let right_output = eval_target(model, params, right)?;
        let right_error = right_output - params.target_value;
        if right_error.abs() < params.tolerance {
            return Ok(Some((right, right, right_error, right_error)));
        }
        if sign(right_error) != sign(start_error) {
            let (lo, hi, lo_err, hi_err) = if start_input < right {
                (start_input, right, start_error, right_error)
            } else {
                (right, start_input, right_error, start_error)
            };
            return Ok(Some((lo, hi, lo_err, hi_err)));
        }

        step *= 2.0;
    }

    Ok(None)
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::collections::HashMap;

    struct FunctionModel<F> {
        changing: CellRef,
        target: CellRef,
        input: f64,
        values: HashMap<CellRef, CellValue>,
        formula: F,
    }

    impl<F> FunctionModel<F>
    where
        F: Fn(f64) -> f64,
    {
        fn new(
            changing: impl Into<CellRef>,
            target: impl Into<CellRef>,
            input: f64,
            formula: F,
        ) -> Self {
            Self {
                changing: changing.into(),
                target: target.into(),
                input,
                values: HashMap::new(),
                formula,
            }
        }
    }

    impl<F> WhatIfModel for FunctionModel<F>
    where
        F: Fn(f64) -> f64,
    {
        type Error = &'static str;

        fn get_cell_value(&self, cell: &CellRef) -> Result<CellValue, Self::Error> {
            if cell == &self.changing {
                return Ok(CellValue::Number(self.input));
            }
            if cell == &self.target {
                return Ok(self.values.get(cell).cloned().unwrap_or(CellValue::Blank));
            }
            Ok(self.values.get(cell).cloned().unwrap_or(CellValue::Blank))
        }

        fn set_cell_value(&mut self, cell: &CellRef, value: CellValue) -> Result<(), Self::Error> {
            if cell == &self.changing {
                self.input = value.as_number().ok_or("changing cell must be numeric")?;
                return Ok(());
            }
            self.values.insert(cell.clone(), value);
            Ok(())
        }

        fn recalculate(&mut self) -> Result<(), Self::Error> {
            let output = (self.formula)(self.input);
            self.values
                .insert(self.target.clone(), CellValue::Number(output));
            Ok(())
        }
    }

    #[test]
    fn goal_seek_converges_on_linear_function() {
        let mut model = FunctionModel::new("A1", "B1", 0.0, |x| 2.0 * x + 3.0);
        let mut params = GoalSeekParams::new("B1", 11.0, "A1");
        params.tolerance = 1e-9;

        let result = GoalSeek::solve(&mut model, params).unwrap();
        assert!(result.success(), "{result:?}");
        assert!((result.solution - 4.0).abs() < 1e-6);
    }

    #[test]
    fn goal_seek_converges_on_quadratic_function() {
        let mut model = FunctionModel::new("A1", "B1", 1.0, |x| x * x);
        let mut params = GoalSeekParams::new("B1", 9.0, "A1");
        params.tolerance = 1e-9;

        let result = GoalSeek::solve(&mut model, params).unwrap();
        assert!(result.success(), "{result:?}");
        assert!((result.solution - 3.0).abs() < 1e-6);
    }

    #[test]
    fn goal_seek_uses_bisection_when_derivative_is_tiny() {
        let mut model = FunctionModel::new("A1", "B1", 100.0, |x| 1.0 + 1e-12 * x);
        let mut params = GoalSeekParams::new("B1", 1.0, "A1");
        params.tolerance = 1e-15;
        params.min_derivative = 1e-10;

        let result = GoalSeek::solve(&mut model, params).unwrap();
        assert!(result.success(), "{result:?}");
        assert!(
            result.solution.abs() < 1e-3,
            "solution = {}",
            result.solution
        );
    }
}
