//! Utilities for making example binaries behave nicely when their stdout is piped.
//!
//! Rust's `println!`/`print!` macros will panic if writing to stdout fails. This commonly happens
//! when piping output to tools like `head` (the downstream process exits early, closing the pipe),
//! which produces an `EPIPE` / "Broken pipe" error.
//!
//! These examples are primarily debugging tools; treating a broken pipe as a fatal error makes them
//! awkward to use. Install a panic hook that exits successfully when the panic is caused by a
//! broken pipe during stdout printing.

/// Exit successfully if a `println!`-triggered panic is caused by a broken pipe.
pub(crate) fn install() {
    let default_hook = std::panic::take_hook();
    std::panic::set_hook(Box::new(move |info| {
        if is_broken_pipe_panic(info) {
            // Exit without printing a backtrace or panic message. This matches common Unix CLI
            // behavior where SIGPIPE / EPIPE is treated as a normal early-termination condition.
            std::process::exit(0);
        }
        default_hook(info);
    }));
}

fn is_broken_pipe_panic(info: &std::panic::PanicHookInfo<'_>) -> bool {
    // The stdlib panic message for stdout printing failures is currently:
    // "failed printing to stdout: Broken pipe (os error 32)".
    //
    // We match on the substring rather than relying on a specific error kind, since the message
    // format is not a stable API and varies across platforms.
    let Some(msg) = panic_message(info) else {
        return false;
    };
    msg.contains("Broken pipe") || msg.contains("broken pipe") || msg.contains("os error 32")
}

fn panic_message(info: &std::panic::PanicHookInfo<'_>) -> Option<String> {
    if let Some(s) = info.payload().downcast_ref::<&str>() {
        return Some((*s).to_string());
    }
    if let Some(s) = info.payload().downcast_ref::<String>() {
        return Some(s.clone());
    }
    None
}
