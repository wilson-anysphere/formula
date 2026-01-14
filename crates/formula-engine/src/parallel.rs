#[cfg(all(feature = "parallel", not(target_arch = "wasm32")))]
use rayon::ThreadPool;
#[cfg(all(feature = "parallel", not(target_arch = "wasm32")))]
use std::sync::OnceLock;

/// Best-effort Rayon's thread pool for use inside the formula engine.
///
/// Rayon normally uses a **global** thread pool. Under extreme resource constraints (e.g. many test
/// binaries running concurrently on a multi-agent host), global pool initialization can fail and
/// Rayon will panic on first use.
///
/// To keep the engine resilient, we build and use a crate-local pool instead. If we can't create a
/// pool, callers should fall back to single-threaded execution.
#[cfg(all(feature = "parallel", not(target_arch = "wasm32")))]
static RAYON_POOL: OnceLock<Option<ThreadPool>> = OnceLock::new();

#[cfg(all(feature = "parallel", not(target_arch = "wasm32")))]
fn desired_rayon_threads() -> usize {
    let from_env = std::env::var("RAYON_NUM_THREADS")
        .ok()
        .and_then(|s| s.parse::<usize>().ok())
        .filter(|&n| n > 0);
    from_env.unwrap_or_else(|| {
        std::thread::available_parallelism()
            .map(|n| n.get())
            .unwrap_or(1)
    })
}

#[cfg(all(feature = "parallel", not(target_arch = "wasm32")))]
fn build_rayon_pool() -> Option<ThreadPool> {
    let requested = desired_rayon_threads().max(1);
    let try_build = |n| rayon::ThreadPoolBuilder::new().num_threads(n).build();

    match try_build(requested) {
        Ok(pool) => Some(pool),
        Err(_) if requested > 1 => try_build(1).ok(),
        Err(_) => None,
    }
}

/// Returns the crate-local Rayon thread pool, if one could be created.
#[cfg(all(feature = "parallel", not(target_arch = "wasm32")))]
pub(crate) fn rayon_pool() -> Option<&'static ThreadPool> {
    RAYON_POOL.get_or_init(build_rayon_pool).as_ref()
}
