use std::process::{Command, Stdio};

#[test]
fn cli_does_not_panic_on_broken_pipe() {
    // Simulate piping the output to a downstream tool that exits early (e.g. `| head`).
    let mut child = Command::new(env!("CARGO_BIN_EXE_function_catalog"))
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .expect("spawn function_catalog");

    // Closing the read end forces stdout writes to return EPIPE / BrokenPipe.
    drop(child.stdout.take());

    let output = child
        .wait_with_output()
        .expect("wait for function_catalog to finish");

    assert!(
        output.status.success(),
        "expected success even when stdout is closed\nstderr:\n{}",
        String::from_utf8_lossy(&output.stderr)
    );
}
