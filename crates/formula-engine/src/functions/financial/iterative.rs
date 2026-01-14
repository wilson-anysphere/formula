pub(crate) const EXCEL_ITERATION_TOLERANCE: f64 = 1.0e-7;

pub(crate) fn newton_raphson<F, DF>(guess: f64, max_iterations: usize, f: F, df: DF) -> Option<f64>
where
    F: Fn(f64) -> Option<f64>,
    DF: Fn(f64) -> Option<f64>,
{
    let mut x = guess;
    for _ in 0..max_iterations {
        let fx = f(x)?;
        let dfx = df(x)?;
        if dfx == 0.0 {
            return None;
        }

        let next = x - fx / dfx;
        if !next.is_finite() {
            return None;
        }

        if (next - x).abs() <= EXCEL_ITERATION_TOLERANCE {
            return Some(next);
        }
        x = next;
    }

    None
}

/// Robust root solver that combines Newton-Raphson with bracketing + bisection.
///
/// Excel's financial functions often rely on Newton iterations but fall back to safer
/// methods when the derivative is unstable or the update would jump out of bounds.
/// This helper approximates that behavior:
/// - First, attempt to *bracket* a root within `[lower_bound, upper_bound]`.
/// - Then, iterate using Newton steps when they stay inside the bracket; otherwise bisect.
///
/// Returns `None` when the root cannot be bracketed or when convergence fails.
pub(crate) fn solve_root_newton_bisection<F, DF>(
    guess: f64,
    lower_bound: f64,
    upper_bound: f64,
    max_iterations: usize,
    f: F,
    df: DF,
) -> Option<f64>
where
    F: Fn(f64) -> Option<f64>,
    DF: Fn(f64) -> Option<f64>,
{
    if !guess.is_finite() || !lower_bound.is_finite() || !upper_bound.is_finite() {
        return None;
    }
    if lower_bound >= upper_bound {
        return None;
    }

    let guess = guess.clamp(lower_bound, upper_bound);
    let f_guess = f(guess)?;
    if !f_guess.is_finite() {
        return None;
    }
    if f_guess == 0.0 {
        return Some(guess);
    }

    // ------------------------------------------------------------------
    // Bracket a sign change around the initial guess.
    // ------------------------------------------------------------------
    // Expand outward with an exponential step until f(a) and f(b) have opposite signs.
    let mut a = guess;
    let mut b = guess;
    let mut fa = f_guess;
    let mut fb = f_guess;

    let mut step = 0.1_f64.max(guess.abs() * 0.1);
    const MAX_BRACKET_STEPS: usize = 60;
    for _ in 0..MAX_BRACKET_STEPS {
        if fa.signum() != fb.signum() {
            break;
        }

        let next_a = (a - step).max(lower_bound);
        if next_a != a {
            a = next_a;
            fa = f(a)?;
            if !fa.is_finite() {
                return None;
            }
            if fa == 0.0 {
                return Some(a);
            }
        }

        let next_b = (b + step).min(upper_bound);
        if next_b != b {
            b = next_b;
            fb = f(b)?;
            if !fb.is_finite() {
                return None;
            }
            if fb == 0.0 {
                return Some(b);
            }
        }

        if a == lower_bound && b == upper_bound {
            break;
        }

        step *= 2.0;
    }

    if fa.signum() == fb.signum() {
        return None;
    }

    // ------------------------------------------------------------------
    // Safeguarded Newton iterations within the bracket.
    // ------------------------------------------------------------------
    // Keep the bracket ordered.
    if a > b {
        std::mem::swap(&mut a, &mut b);
        std::mem::swap(&mut fa, &mut fb);
    }

    // Start inside the bracket; prefer the original guess if it's in range.
    let mut x = if (a..=b).contains(&guess) {
        guess
    } else {
        (a + b) * 0.5
    };
    let mut fx = f_guess;
    if !(a..=b).contains(&x) {
        x = (a + b) * 0.5;
        fx = f(x)?;
    }

    for _ in 0..max_iterations {
        if fx == 0.0 {
            return Some(x);
        }
        if (b - a).abs() <= EXCEL_ITERATION_TOLERANCE {
            return Some(x);
        }

        let mut next = None;
        if let Some(dfx) = df(x) {
            if dfx.is_finite() && dfx != 0.0 {
                let candidate = x - fx / dfx;
                if candidate.is_finite() && candidate > a && candidate < b {
                    next = Some(candidate);
                }
            }
        }
        let next = next.unwrap_or_else(|| (a + b) * 0.5);
        let f_next = f(next)?;
        if !f_next.is_finite() {
            return None;
        }

        if f_next == 0.0 {
            return Some(next);
        }

        // Maintain the bracket invariant.
        if fa.signum() == f_next.signum() {
            a = next;
            fa = f_next;
        } else {
            b = next;
        }

        if (next - x).abs() <= EXCEL_ITERATION_TOLERANCE {
            return Some(next);
        }

        x = next;
        fx = f_next;
    }

    None
}
