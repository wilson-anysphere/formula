use std::path::{Path, PathBuf};
use std::process::Command;

fn triage_bin() -> &'static str {
    // Cargo sets `CARGO_BIN_EXE_<name>` for integration tests. Binary names may contain `-`,
    // but some environments/tools normalize them to `_`, so accept either.
    option_env!("CARGO_BIN_EXE_formula-corpus-triage")
        .or(option_env!("CARGO_BIN_EXE_formula_corpus_triage"))
        .expect("formula-corpus-triage binary should be built for integration tests")
}

fn fixture_path(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/encrypted/ooxml")
        .join(name)
}

#[test]
fn cli_supports_empty_password_files_for_encrypted_workbooks() {
    let encrypted = fixture_path("agile-empty-password.xlsx");

    let tmp = tempfile::tempdir().expect("tempdir");
    let pw_path = tmp.path().join("password.txt");
    // A file containing only a newline represents the empty password.
    std::fs::write(&pw_path, "\r\n").expect("write password file");

    let output = Command::new(triage_bin())
        .args(["--input", encrypted.to_str().unwrap()])
        .args(["--password-file", pw_path.to_str().unwrap()])
        .output()
        .expect("run formula-corpus-triage");

    assert!(
        output.status.success(),
        "expected exit 0\nstdout:\n{}\nstderr:\n{}",
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr)
    );

    let stdout = String::from_utf8_lossy(&output.stdout);
    let json: serde_json::Value = serde_json::from_str(&stdout).expect("parse output json");
    assert_eq!(
        json.pointer("/result/open_ok").and_then(|v| v.as_bool()),
        Some(true),
        "expected triage to decrypt and open the workbook\nstdout:\n{stdout}"
    );
}

