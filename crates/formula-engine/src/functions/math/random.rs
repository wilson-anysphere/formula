use crate::error::{ExcelError, ExcelResult};
use std::sync::atomic::{AtomicU64, Ordering};
use std::time::{SystemTime, UNIX_EPOCH};

static RAND_STATE: AtomicU64 = AtomicU64::new(0);

fn seed_state() -> u64 {
    let nanos = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap_or_default()
        .as_nanos();
    // Mix in the current address of the static for a tiny bit more entropy.
    let addr = (&RAND_STATE as *const AtomicU64 as usize) as u64;
    (nanos as u64) ^ addr.rotate_left(17) ^ 0x9e3779b97f4a7c15
}

fn next_u64() -> u64 {
    RAND_STATE
        .fetch_update(Ordering::SeqCst, Ordering::SeqCst, |state| {
            let mut s = state;
            if s == 0 {
                s = seed_state();
            }
            s = s.wrapping_mul(6364136223846793005).wrapping_add(1);
            Some(s)
        })
        .unwrap_or_else(|state| {
            let mut s = if state == 0 { seed_state() } else { state };
            s = s.wrapping_mul(6364136223846793005).wrapping_add(1);
            s
        })
}

/// RAND()
///
/// Returns a pseudorandom number in the range [0, 1).
pub fn rand() -> f64 {
    let bits = next_u64() >> 11; // 53 bits.
    (bits as f64) / ((1u64 << 53) as f64)
}

/// RANDBETWEEN(bottom, top)
///
/// Returns a pseudorandom integer in [bottom, top] (inclusive).
pub fn randbetween(bottom: f64, top: f64) -> ExcelResult<i64> {
    if !bottom.is_finite() || !top.is_finite() {
        return Err(ExcelError::Num);
    }

    let low = bottom.ceil() as i64;
    let high = top.floor() as i64;
    if low > high {
        return Err(ExcelError::Num);
    }

    let span = high - low + 1;
    if span <= 0 {
        return Err(ExcelError::Num);
    }

    let r = rand();
    let offset = (r * span as f64).floor() as i64;
    Ok(low + offset.min(span - 1))
}

