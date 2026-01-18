use std::fmt;
use std::io::Write;

#[allow(dead_code)]
pub(crate) fn stdoutln(args: fmt::Arguments<'_>) {
    let mut out = std::io::stdout().lock();
    let _ = out.write_fmt(args);
    let _ = out.write_all(b"\n");
}

pub(crate) fn stderrln(args: fmt::Arguments<'_>) {
    let mut out = std::io::stderr().lock();
    let _ = out.write_fmt(args);
    let _ = out.write_all(b"\n");
}

