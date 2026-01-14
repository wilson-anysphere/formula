use std::process::{Command, Stdio};

#[test]
fn cli_does_not_panic_on_broken_pipe() {
    // The benchmark runner can take a while. Closing stdout up-front should trigger a BrokenPipe
    // early and exit without panicking.
    let mut child = Command::new(env!("CARGO_BIN_EXE_perf_bench"))
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .expect("spawn perf_bench");

    // Closing the read end forces stdout writes to return EPIPE / BrokenPipe.
    drop(child.stdout.take());

    let output = child.wait_with_output().expect("wait for perf_bench");

    assert!(
        output.status.success(),
        "expected success even when stdout is closed\nstderr:\n{}",
        String::from_utf8_lossy(&output.stderr)
    );
}
