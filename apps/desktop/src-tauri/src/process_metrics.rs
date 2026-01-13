use anyhow::{anyhow, Result};
use sysinfo::System;

#[derive(Clone, Debug)]
pub struct ProcessMetricsSnapshot {
    pub pid: u32,
    /// Resident set size in bytes (best-effort; platform-dependent).
    pub rss_bytes: u64,
    /// CPU usage percentage as reported by `sysinfo` (best-effort).
    pub cpu_usage_pct: Option<f32>,
}

pub fn snapshot_self() -> Result<ProcessMetricsSnapshot> {
    let pid = sysinfo::get_current_pid().map_err(|err| anyhow!(err))?;
    let mut system = System::new();
    system.refresh_processes();

    let process = system
        .process(pid)
        .ok_or_else(|| anyhow!("process {pid:?} not found in sysinfo snapshot"))?;

    // `sysinfo` reports process memory in kilobytes.
    let rss_bytes = process.memory().saturating_mul(1024);

    Ok(ProcessMetricsSnapshot {
        pid: pid.as_u32(),
        rss_bytes,
        cpu_usage_pct: Some(process.cpu_usage()),
    })
}

/// Print a single-line process metrics log to stdout.
///
/// Output format (stable):
/// `[metrics] rss_mb=<n> pid=<pid>`
pub fn log_process_metrics() {
    match snapshot_self() {
        Ok(snapshot) => {
            let rss_mb = snapshot.rss_bytes / (1024 * 1024);
            println!("[metrics] rss_mb={rss_mb} pid={}", snapshot.pid);
        }
        Err(err) => {
            let pid = std::process::id();
            eprintln!("[metrics] failed to collect process metrics: {err}");
            // Keep the stdout format stable for log parsers even on failure.
            println!("[metrics] rss_mb=0 pid={pid}");
        }
    }
}
