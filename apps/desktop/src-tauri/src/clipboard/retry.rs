use std::time::Duration;

/// Deterministic retry delays for Windows `OpenClipboard`.
///
/// The Windows clipboard is a globally shared resource. When another process temporarily holds the
/// clipboard lock, `OpenClipboard` can fail with a transient error. A short fixed retry window (e.g.
/// 100ms) is often insufficient in practice under real-world contention (rapid copy/paste between
/// apps, large clipboard payloads, etc).
///
/// We use an exponential backoff with a bounded total sleep budget (roughly ~1s) to improve
/// reliability without unbounded worst-case latency.
pub(crate) const OPEN_CLIPBOARD_RETRY_DELAYS: &[Duration] = &[
    Duration::from_millis(5),
    Duration::from_millis(10),
    Duration::from_millis(20),
    Duration::from_millis(40),
    Duration::from_millis(80),
    Duration::from_millis(160),
    Duration::from_millis(160),
    Duration::from_millis(160),
    Duration::from_millis(160),
    Duration::from_millis(160),
];

/// Retry `op` using the provided deterministic sleep schedule.
///
/// - The operation is attempted once immediately.
/// - After each failure, we sleep for the next delay and retry.
/// - When delays are exhausted, the final error is returned.
///
/// This is deliberately written to be unit-testable without platform APIs by injecting both the
/// operation and sleep functions.
pub(crate) fn retry_with_delays<T, E>(
    mut op: impl FnMut() -> Result<T, E>,
    delays: &[Duration],
    mut sleep: impl FnMut(Duration),
) -> Result<T, E> {
    for delay in delays {
        match op() {
            Ok(value) => return Ok(value),
            Err(_) => {
                sleep(*delay);
            }
        }
    }
    op()
}

pub(crate) fn total_delay(delays: &[Duration]) -> Duration {
    delays
        .iter()
        .copied()
        .fold(Duration::from_millis(0), |acc, d| acc + d)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn open_clipboard_retry_delays_have_reasonable_budget() {
        let total = total_delay(OPEN_CLIPBOARD_RETRY_DELAYS);
        assert!(
            total >= Duration::from_millis(500),
            "retry budget should be >= 500ms, got {total:?}"
        );
        assert!(
            total <= Duration::from_millis(1_000),
            "retry budget should be <= 1s, got {total:?}"
        );
    }

    #[test]
    fn retry_with_delays_sleeps_and_retries_until_success() {
        let delays = [
            Duration::from_millis(1),
            Duration::from_millis(2),
            Duration::from_millis(3),
        ];

        let mut attempts = 0usize;
        let mut sleeps = Vec::new();

        let result = retry_with_delays(
            || {
                attempts += 1;
                if attempts <= 3 {
                    Err("nope")
                } else {
                    Ok(42)
                }
            },
            &delays,
            |d| sleeps.push(d),
        );

        assert_eq!(result, Ok(42));
        assert_eq!(attempts, 4);
        assert_eq!(sleeps, delays);
    }

    #[test]
    fn retry_with_delays_returns_final_error_after_exhausting_delays() {
        let delays = [Duration::from_millis(1), Duration::from_millis(2)];

        let mut attempts = 0usize;
        let mut sleeps = Vec::new();

        let result: Result<(), usize> = retry_with_delays(
            || {
                attempts += 1;
                Err(attempts)
            },
            &delays,
            |d| sleeps.push(d),
        );

        // `delays.len() + 1` attempts: error from the final attempt should be returned.
        assert_eq!(attempts, delays.len() + 1);
        assert_eq!(result, Err(delays.len() + 1));
        assert_eq!(sleeps, delays);
    }

    #[test]
    fn retry_with_delays_attempts_once_when_delays_empty() {
        let mut attempts = 0usize;
        let mut slept = false;

        let result: Result<(), ()> = retry_with_delays(
            || {
                attempts += 1;
                Ok(())
            },
            &[],
            |_| slept = true,
        );

        assert_eq!(result, Ok(()));
        assert_eq!(attempts, 1);
        assert!(!slept);
    }
}
