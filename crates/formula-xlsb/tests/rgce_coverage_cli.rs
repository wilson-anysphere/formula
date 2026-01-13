use serde_json::Value;
use std::ffi::OsStr;
use std::path::Path;
use std::process::Command;

use formula_xlsb::XlsbWorkbook;

mod common;

fn run_rgce_coverage(path: &Path, extra_args: &[&str]) -> Vec<Value> {
    let rgce_coverage_exe = env!("CARGO_BIN_EXE_rgce_coverage");
    let mut cmd = Command::new(rgce_coverage_exe);
    cmd.arg(path).arg("--max").arg("100000");
    cmd.args(extra_args);

    let output = cmd.output().expect("run rgce_coverage");
    let fixture_name = path
        .file_name()
        .and_then(OsStr::to_str)
        .unwrap_or("<unknown>");

    assert!(
        output.status.success(),
        "rgce_coverage failed for fixture {fixture_name}: status={:?}\nstdout:\n{}\nstderr:\n{}",
        output.status.code(),
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr)
    );

    let stdout = String::from_utf8(output.stdout).expect("stdout is utf-8");
    let mut values = Vec::new();
    for (idx, line) in stdout.lines().enumerate() {
        let line = line.trim();
        if line.is_empty() {
            continue;
        }
        let value: Value = serde_json::from_str(line).unwrap_or_else(|err| {
            panic!(
                "fixture {fixture_name}: invalid JSON on line {}: {err}: {line}",
                idx + 1
            )
        });
        values.push(value);
    }
    values
}

fn summary_from_output(values: &[Value]) -> &Value {
    let summary = values.last().expect("expected summary line");
    assert_eq!(summary["kind"], "summary");
    summary
}

#[test]
fn rgce_coverage_cli_reports_zero_failures_for_fixtures() {
    let manifest_dir = Path::new(env!("CARGO_MANIFEST_DIR"));
    let fixtures = [
        "tests/fixtures/simple.xlsb",
        "tests/fixtures/date1904.xlsb",
        "tests/fixtures/rich_shared_strings.xlsb",
        "tests/fixtures/udf.xlsb",
        "tests/fixtures_styles/date.xlsb",
    ];

    for rel in fixtures {
        let path = manifest_dir.join(rel);
        let output = run_rgce_coverage(&path, &[]);
        let summary = summary_from_output(&output);

        assert_eq!(summary["decoded_failed"], 0, "fixture {rel}");
        assert_eq!(
            summary["formulas_total"],
            summary["decoded_ok"].as_u64().unwrap_or(0)
                + summary["decoded_failed"].as_u64().unwrap_or(0),
            "fixture {rel}: totals should add up"
        );

        // Ensure selector options work and don't introduce decode failures.
        // (If the workbook has multiple sheets, totals may differ from the all-sheets run.)
        let wb = XlsbWorkbook::open(&path).expect("open xlsb fixture");
        let first_sheet_name = wb
            .sheet_metas()
            .first()
            .expect("fixture has at least one sheet")
            .name
            .clone();

        let sheet0 = run_rgce_coverage(&path, &["--sheet", "0"]);
        let sheet0_summary = summary_from_output(&sheet0);
        assert_eq!(
            sheet0_summary["decoded_failed"], 0,
            "fixture {rel} --sheet 0"
        );

        let by_name = run_rgce_coverage(&path, &["--sheet", &first_sheet_name]);
        let by_name_summary = summary_from_output(&by_name);
        assert_eq!(
            by_name_summary["decoded_failed"], 0,
            "fixture {rel} --sheet {first_sheet_name}"
        );

        let by_name_ci = run_rgce_coverage(&path, &["--sheet", &first_sheet_name.to_lowercase()]);
        let by_name_ci_summary = summary_from_output(&by_name_ci);
        assert_eq!(
            by_name_ci_summary["decoded_failed"],
            0,
            "fixture {rel} --sheet {} (case-insensitive)",
            first_sheet_name.to_lowercase()
        );

        // Max should cap formulas_total when formulas exist.
        let capped = run_rgce_coverage(&path, &["--max", "1"]);
        let capped_summary = summary_from_output(&capped);
        assert_eq!(capped_summary["decoded_failed"], 0, "fixture {rel} --max 1");
        assert!(
            capped_summary["formulas_total"].as_u64().unwrap_or(0) <= 1,
            "fixture {rel} --max 1 should cap formulas_total"
        );

        let zero = run_rgce_coverage(&path, &["--max", "0"]);
        let zero_summary = summary_from_output(&zero);
        assert_eq!(zero_summary["decoded_failed"], 0, "fixture {rel} --max 0");
        assert_eq!(
            zero_summary["formulas_total"].as_u64().unwrap_or(0),
            0,
            "fixture {rel} --max 0 should produce zero formulas"
        );
        assert_eq!(
            zero.len(),
            1,
            "fixture {rel} --max 0 should emit only the summary line"
        );
    }
}

#[test]
fn rgce_coverage_help_mentions_password_flag() {
    let rgce_coverage_exe = env!("CARGO_BIN_EXE_rgce_coverage");
    let output = Command::new(rgce_coverage_exe)
        .arg("--help")
        .output()
        .expect("run rgce_coverage --help");

    assert!(
        output.status.success(),
        "expected --help to exit successfully, status={:?}",
        output.status.code()
    );

    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(
        stdout.contains("--password"),
        "expected --password in help output, got:\n{stdout}"
    );
}

#[test]
fn rgce_coverage_errors_when_password_missing_for_encrypted_ooxml_wrapper() {
    use std::io::Cursor;

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    ole.create_stream("EncryptionInfo")
        .expect("create EncryptionInfo stream");
    ole.create_stream("EncryptedPackage")
        .expect("create EncryptedPackage stream");
    let bytes = ole.into_inner().into_inner();

    let tmp = tempfile::tempdir().expect("tempdir");
    let path = tmp.path().join("encrypted.xlsb");
    std::fs::write(&path, bytes).expect("write fixture");

    let rgce_coverage_exe = env!("CARGO_BIN_EXE_rgce_coverage");
    let output = Command::new(rgce_coverage_exe)
        .arg(&path)
        .output()
        .expect("run rgce_coverage");

    assert!(
        !output.status.success(),
        "expected rgce_coverage to fail when password is missing"
    );

    let stderr = String::from_utf8_lossy(&output.stderr).to_lowercase();
    assert!(
        stderr.contains("password"),
        "expected stderr to mention password, got:\n{stderr}"
    );
}

#[test]
fn rgce_coverage_opens_standard_encrypted_xlsb_with_password() {
    let plaintext_path = Path::new(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/tests/fixtures/simple.xlsb"
    ));
    let plaintext_bytes = std::fs::read(plaintext_path).expect("read xlsb fixture");

    let tmp = tempfile::tempdir().expect("tempdir");
    let password = "Password1234_";
    let encrypted = common::standard_encrypted_ooxml::build_standard_encrypted_ooxml_ole_bytes(
        &plaintext_bytes,
        password,
    );
    let encrypted_path = tmp.path().join("encrypted_standard.xlsb");
    std::fs::write(&encrypted_path, encrypted).expect("write encrypted fixture");

    // Use max=0 so we only exercise workbook open + selector plumbing (fast).
    let output = run_rgce_coverage(&encrypted_path, &["--password", password, "--max", "0"]);
    let summary = summary_from_output(&output);
    assert_eq!(summary["formulas_total"], 0);
}
