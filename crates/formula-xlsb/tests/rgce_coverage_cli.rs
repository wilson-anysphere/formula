use serde_json::Value;
use std::path::Path;
use std::process::Command;

#[test]
fn rgce_coverage_cli_reports_zero_failures_for_fixtures() {
    let manifest_dir = Path::new(env!("CARGO_MANIFEST_DIR"));
    let rgce_coverage_exe = env!("CARGO_BIN_EXE_rgce_coverage");
    let fixtures = [
        "tests/fixtures/simple.xlsb",
        "tests/fixtures/date1904.xlsb",
        "tests/fixtures/rich_shared_strings.xlsb",
        "tests/fixtures/udf.xlsb",
    ];

    for rel in fixtures {
        let path = manifest_dir.join(rel);
        let output = Command::new(rgce_coverage_exe)
            .arg(&path)
            .arg("--max")
            .arg("100000")
            .output()
            .expect("run rgce_coverage");
        assert!(
            output.status.success(),
            "rgce_coverage failed for fixture {rel}: status={:?}\nstdout:\n{}\nstderr:\n{}",
            output.status.code(),
            String::from_utf8_lossy(&output.stdout),
            String::from_utf8_lossy(&output.stderr)
        );
        let stdout = String::from_utf8(output.stdout).expect("stdout is utf-8");

        let mut last_nonempty = None;
        for (idx, line) in stdout.lines().enumerate() {
            let line = line.trim();
            if line.is_empty() {
                continue;
            }
            serde_json::from_str::<Value>(line).unwrap_or_else(|err| {
                panic!("fixture {rel}: invalid JSON on line {}: {err}: {line}", idx + 1)
            });
            last_nonempty = Some(line);
        }

        let summary_line = last_nonempty.expect("expected summary line");
        let summary: Value = serde_json::from_str(summary_line).expect("parse summary JSON");
        assert_eq!(summary["kind"], "summary", "fixture {rel}");
        assert_eq!(summary["decoded_failed"], 0, "fixture {rel}");
    }
}
