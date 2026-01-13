use assert_cmd::prelude::*;
use std::process::Command;

mod common;

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

#[test]
fn xlsb_dump_prints_known_error_literals() {
    let fixture_path = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/simple.xlsb");
    let wb = formula_xlsb::XlsbWorkbook::open(fixture_path).expect("open xlsb fixture");

    // Patch A1 to a modern error code (#SPILL!) so the CLI should display the literal rather than
    // a raw hex code.
    let dir = tempfile::tempdir().expect("temp dir");
    let patched = dir.path().join("errors.xlsb");
    wb.save_with_cell_edits(
        &patched,
        0,
        &[formula_xlsb::CellEdit {
            row: 0,
            col: 0,
            new_value: formula_xlsb::CellValue::Error(0x2C), // #SPILL!
            new_formula: None,
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
            clear_formula: false,
        }],
    )
    .expect("save patched xlsb");

    let assert = Command::new(assert_cmd::cargo::cargo_bin!("xlsb_dump"))
        .arg(&patched)
        .arg("--sheet")
        .arg("0")
        .arg("--max")
        .arg("10")
        .assert()
        .success();

    let stdout = String::from_utf8_lossy(&assert.get_output().stdout);
    assert!(stdout.contains("#SPILL!"), "stdout:\n{stdout}");
}

#[test]
fn xlsb_dump_help_mentions_password_flag() {
    let assert = Command::new(assert_cmd::cargo::cargo_bin!("xlsb_dump"))
        .arg("--help")
        .assert()
        .success();

    let stdout = String::from_utf8_lossy(&assert.get_output().stdout);
    assert!(
        stdout.contains("--password"),
        "expected --password in help output, got:\n{stdout}"
    );
}

#[test]
fn xlsb_dump_errors_when_password_missing_for_encrypted_ooxml_wrapper() {
    use std::io::Cursor;

    // Synthetic OLE/CFB container that looks like Office-encrypted OOXML.
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

    let assert = Command::new(assert_cmd::cargo::cargo_bin!("xlsb_dump"))
        .arg(&path)
        .assert()
        .failure();

    let stderr = String::from_utf8_lossy(&assert.get_output().stderr).to_lowercase();
    assert!(
        stderr.contains("password"),
        "expected stderr to mention password, got:\n{stderr}"
    );
}

#[test]
fn xlsb_dump_opens_standard_encrypted_xlsb_with_password() {
    let plaintext_path = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/simple.xlsb");
    let plaintext_bytes = std::fs::read(plaintext_path).expect("read xlsb fixture");

    let tmp = tempfile::tempdir().expect("tempdir");
    let password = "Password1234_";
    let encrypted = common::standard_encrypted_ooxml::build_standard_encrypted_ooxml_ole_bytes(
        &plaintext_bytes,
        password,
    );
    let encrypted_path = tmp.path().join("encrypted_standard.xlsb");
    std::fs::write(&encrypted_path, encrypted).expect("write encrypted fixture");

    let assert = Command::new(assert_cmd::cargo::cargo_bin!("xlsb_dump"))
        .arg("--password")
        .arg(password)
        .arg(&encrypted_path)
        .assert()
        .success();

    let stdout = String::from_utf8_lossy(&assert.get_output().stdout);
    assert!(stdout.contains("Sheet1"), "stdout:\n{stdout}");
}
