use formula_engine::run_benchmarks;
use std::io::{self, Write};

fn escape_json_string(input: &str) -> String {
    let mut out = String::new();
    let _ = out.try_reserve_exact(input.len().saturating_add(8));
    for ch in input.chars() {
        match ch {
            '"' => out.push_str("\\\""),
            '\\' => out.push_str("\\\\"),
            '\n' => out.push_str("\\n"),
            '\r' => out.push_str("\\r"),
            '\t' => out.push_str("\\t"),
            _ => out.push(ch),
        }
    }
    out
}

fn main() {
    let stdout = io::stdout();
    let mut out = io::BufWriter::new(stdout.lock());

    // Emit a stable JSON payload that the TypeScript harness can merge.
    //
    // Write the prefix + flush first so we can detect a closed stdout pipe before running
    // benchmarks (e.g. if the downstream consumer exits immediately).
    if let Err(err) = write!(&mut out, "{{\"benchmarks\":[") {
        if err.kind() == io::ErrorKind::BrokenPipe {
            return;
        }
        eprintln!("error: failed to write benchmark output: {err}");
        std::process::exit(1);
    }
    if let Err(err) = out.flush() {
        if err.kind() == io::ErrorKind::BrokenPipe {
            return;
        }
        eprintln!("error: failed to flush benchmark output: {err}");
        std::process::exit(1);
    }

    let results = run_benchmarks();

    for (i, r) in results.iter().enumerate() {
        if i > 0 {
            if let Err(err) = write!(&mut out, ",") {
                if err.kind() == io::ErrorKind::BrokenPipe {
                    return;
                }
                eprintln!("error: failed to write benchmark output: {err}");
                std::process::exit(1);
            }
        }
        if let Err(err) = write!(
            &mut out,
            "{{\"name\":\"{}\",\"iterations\":{},\"warmup\":{},\"unit\":\"{}\",\"mean\":{},\"median\":{},\"p95\":{},\"p99\":{},\"stdDev\":{},\"targetMs\":{},\"passed\":{}}}",
            escape_json_string(&r.name),
            r.iterations,
            r.warmup,
            r.unit,
            r.mean,
            r.median,
            r.p95,
            r.p99,
            r.std_dev,
            r.target_ms,
            if r.passed { "true" } else { "false" }
        ) {
            if err.kind() == io::ErrorKind::BrokenPipe {
                return;
            }
            eprintln!("error: failed to write benchmark output: {err}");
            std::process::exit(1);
        }
    }

    if let Err(err) = writeln!(&mut out, "]}}") {
        if err.kind() == io::ErrorKind::BrokenPipe {
            return;
        }
        eprintln!("error: failed to write benchmark output: {err}");
        std::process::exit(1);
    }
    if let Err(err) = out.flush() {
        if err.kind() == io::ErrorKind::BrokenPipe {
            return;
        }
        eprintln!("error: failed to flush benchmark output: {err}");
        std::process::exit(1);
    }
}
