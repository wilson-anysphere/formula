use std::io::{Cursor, Write};
use std::path::{Path, PathBuf};
use std::process::Command;

use zip::write::FileOptions;
use zip::{CompressionMethod, ZipWriter};

fn zip_bytes(parts: &[(&str, &[u8])]) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut writer = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(CompressionMethod::Stored);

    for (name, bytes) in parts {
        writer.start_file(*name, options).unwrap();
        writer.write_all(bytes).unwrap();
    }

    writer.finish().unwrap().into_inner()
}

fn encrypted_fixture_path(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/encrypted/ooxml")
        .join(name)
}

#[test]
fn cli_json_output_is_parseable_and_contains_expected_fields() {
    let expected_zip = zip_bytes(&[
        ("xl/theme/theme1.xml", br#"<a attr="1"/>"#),
        ("xl/theme/theme2.xml", br#"<a attr="3"/>"#),
    ]);
    let actual_zip = zip_bytes(&[
        ("xl/theme/theme1.xml", br#"<a attr="2"/>"#),
        ("xl/theme/theme2.xml", br#"<a attr="4"/>"#),
    ]);

    let tempdir = tempfile::tempdir().unwrap();
    let original_path = tempdir.path().join("original.xlsx");
    let modified_path = tempdir.path().join("modified.xlsx");
    std::fs::write(&original_path, expected_zip).unwrap();
    std::fs::write(&modified_path, actual_zip).unwrap();

    let output = Command::new(env!("CARGO_BIN_EXE_xlsx_diff"))
        .arg(&original_path)
        .arg(&modified_path)
        .arg("--format")
        .arg("json")
        .arg("--max-diffs")
        .arg("1")
        .arg("--ignore-part")
        .arg("foo/bar.xml")
        .arg("--ignore-glob")
        .arg("docProps/*")
        .arg("--ignore-path")
        .arg("some-noisy-path")
        .arg("--strict-calc-chain")
        .output()
        .unwrap();

    assert!(
        output.status.success(),
        "expected exit 0, got {:?}\nstdout:\n{}\nstderr:\n{}",
        output.status,
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr)
    );

    let json: serde_json::Value = serde_json::from_slice(&output.stdout).unwrap();
    assert_eq!(
        json["original"].as_str().unwrap(),
        original_path.to_string_lossy()
    );
    assert_eq!(
        json["modified"].as_str().unwrap(),
        modified_path.to_string_lossy()
    );
    assert_eq!(json["ignore_parts"], serde_json::json!(["foo/bar.xml"]));
    assert_eq!(json["ignore_globs"], serde_json::json!(["docProps/*"]));
    assert_eq!(
        json["ignore_paths"],
        serde_json::json!([{ "part": serde_json::Value::Null, "path_substring": "some-noisy-path", "kind": serde_json::Value::Null }])
    );
    assert_eq!(json["strict_calc_chain"], true);

    assert_eq!(json["counts"]["critical"].as_u64().unwrap(), 0);
    assert_eq!(json["counts"]["warning"].as_u64().unwrap(), 2);
    assert_eq!(json["counts"]["info"].as_u64().unwrap(), 0);

    let diffs = json["diffs"].as_array().unwrap();
    assert_eq!(diffs.len(), 1, "expected diffs list to be truncated");
    let diff = &diffs[0];
    assert_eq!(diff["severity"], "warning");
    assert_eq!(diff["part"], "xl/theme/theme1.xml");
    assert_eq!(diff["kind"], "attribute_changed");
    assert_eq!(diff["path"], "/@attr");
    assert_eq!(diff["expected"], "1");
    assert_eq!(diff["actual"], "2");
}

#[test]
fn cli_json_output_succeeds_for_encrypted_workbooks() {
    let plain = encrypted_fixture_path("plaintext.xlsx");
    let encrypted = encrypted_fixture_path("agile.xlsx");

    let output = Command::new(env!("CARGO_BIN_EXE_xlsx_diff"))
        .arg(&plain)
        .arg(&encrypted)
        .arg("--format")
        .arg("json")
        .arg("--modified-password")
        .arg("password")
        .output()
        .unwrap();

    assert!(
        output.status.success(),
        "expected exit 0, got {:?}\nstdout:\n{}\nstderr:\n{}",
        output.status,
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr)
    );

    let json: serde_json::Value = serde_json::from_slice(&output.stdout).unwrap();
    assert_eq!(json["original"].as_str().unwrap(), plain.to_string_lossy());
    assert_eq!(json["modified"].as_str().unwrap(), encrypted.to_string_lossy());
    assert_eq!(json["counts"]["critical"].as_u64().unwrap(), 0);
    assert_eq!(json["counts"]["warning"].as_u64().unwrap(), 0);
    assert_eq!(json["counts"]["info"].as_u64().unwrap(), 0);
    assert_eq!(json["diffs"].as_array().unwrap().len(), 0);
}
