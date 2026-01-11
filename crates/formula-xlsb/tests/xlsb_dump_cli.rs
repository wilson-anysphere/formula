use assert_cmd::prelude::*;
use std::process::Command;

#[test]
fn xlsb_dump_prints_sheet_and_formula() {
    let path = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/simple.xlsb");

    let assert = Command::new(assert_cmd::cargo::cargo_bin!("xlsb_dump"))
        .arg(path)
        .assert()
        .success();

    let stdout = String::from_utf8_lossy(&assert.get_output().stdout);
    assert!(stdout.contains("Sheet1"), "stdout:\n{stdout}");
    assert!(
        stdout.contains("B1*2") || stdout.contains("rgce="),
        "stdout:\n{stdout}"
    );
}
