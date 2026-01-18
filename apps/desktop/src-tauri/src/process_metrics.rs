use anyhow::{anyhow, Result};
use sysinfo::System;

use crate::stdio::{stderrln, stdoutln};

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
    // Refresh only the current process to keep this cheap when called frequently.
    system.refresh_process(pid);

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
            stdoutln(format_args!(
                "[metrics] rss_mb={rss_mb} pid={}",
                snapshot.pid
            ));
        }
        Err(err) => {
            let pid = std::process::id();
            stderrln(format_args!(
                "[metrics] failed to collect process metrics: {err}"
            ));
            // Keep the stdout format stable for log parsers even on failure.
            stdoutln(format_args!("[metrics] rss_mb=0 pid={pid}"));
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn snapshot_self_reports_current_pid() {
        let snapshot = snapshot_self().expect("snapshot_self");
        assert_eq!(snapshot.pid, std::process::id());
    }
}
