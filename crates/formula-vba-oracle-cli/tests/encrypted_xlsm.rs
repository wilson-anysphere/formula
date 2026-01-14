use assert_cmd::prelude::*;
use std::process::Command;

#[test]
fn extract_supports_encrypted_xlsm_with_password() {
    let fixture_path =
        concat!(env!("CARGO_MANIFEST_DIR"), "/../../fixtures/encrypted/ooxml/basic-encrypted.xlsm");

    let assert = Command::new(assert_cmd::cargo::cargo_bin!("formula-vba-oracle-cli"))
        .args([
            "extract",
            "--input",
            fixture_path,
            "--format",
            "auto",
            "--password",
            "password",
        ])
        .assert()
        .success();

    let stdout = String::from_utf8_lossy(&assert.get_output().stdout);
    let json: serde_json::Value = serde_json::from_str(&stdout).expect("parse output json");

    assert_eq!(json["ok"], true, "stdout:\n{stdout}");
    let modules = json["workbook"]["vbaModules"]
        .as_array()
        .expect("vbaModules is array");
    assert!(
        modules.iter().any(|m| m["name"] == "Module1"),
        "stdout:\n{stdout}"
    );
    let procedures = json["procedures"]
        .as_array()
        .expect("procedures is array");
    assert!(
        procedures.iter().any(|p| p["name"] == "Hello"),
        "stdout:\n{stdout}"
    );
}

#[test]
fn extract_reports_wrong_password_for_encrypted_xlsm() {
    let fixture_path =
        concat!(env!("CARGO_MANIFEST_DIR"), "/../../fixtures/encrypted/ooxml/basic-encrypted.xlsm");

    let assert = Command::new(assert_cmd::cargo::cargo_bin!("formula-vba-oracle-cli"))
        .args([
            "extract",
            "--input",
            fixture_path,
            "--format",
            "auto",
            "--password",
            "wrongpw",
        ])
        .assert()
        .failure();

    let stdout = String::from_utf8_lossy(&assert.get_output().stdout);
    let json: serde_json::Value = serde_json::from_str(&stdout).expect("parse output json");
    assert_eq!(json["ok"], false, "stdout:\n{stdout}");
    let msg = json["error"].as_str().unwrap_or("").to_lowercase();
    assert!(
        msg.contains("password"),
        "expected password-related error, got: {msg}\nstdout:\n{stdout}"
    );
}
