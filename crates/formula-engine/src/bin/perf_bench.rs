use formula_engine::run_benchmarks;

fn escape_json_string(input: &str) -> String {
    let mut out = String::with_capacity(input.len() + 8);
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
    let results = run_benchmarks();

    // Emit a stable JSON payload that the TypeScript harness can merge.
    // Format:
    // { "benchmarks": [ { ... }, ... ] }
    print!("{{\"benchmarks\":[");

    for (i, r) in results.iter().enumerate() {
        if i > 0 {
            print!(",");
        }
        print!(
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
        );
    }

    println!("]}}");
}

