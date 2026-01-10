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
