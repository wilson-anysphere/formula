#![cfg(not(target_arch = "wasm32"))]

use std::io::Read;
use std::time::{SystemTime, UNIX_EPOCH};

#[path = "../examples/shared/vba_project_bin.rs"]
mod vba_project_bin;

#[test]
fn example_loader_handles_encrypted_ooxml_workbooks() {
    // Use the existing macro-enabled fixture so we can verify round-trip extraction of
    // `xl/vbaProject.bin` after encrypting/decrypting the workbook package.
    let fixture_path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/xlsx/macros/basic.xlsm"
    );
    let zip_bytes = std::fs::read(fixture_path).expect("read xlsm fixture");

    // Encrypt the workbook bytes into an OLE wrapper (EncryptionInfo + EncryptedPackage).
    let password = "password";
    let ole_bytes = formula_office_crypto::encrypt_package_to_ole(
        &zip_bytes,
        password,
        formula_office_crypto::EncryptOptions::default(),
    )
    .expect("encrypt to OLE");

    // Write the encrypted bytes to a temp file so we exercise the example loader's
    // path-based logic.
    let mut path = std::env::temp_dir();
    let uniq = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("clock set")
        .as_nanos();
    path.push(format!("formula-vba-encrypted-{uniq}.xlsm"));
    std::fs::write(&path, &ole_bytes).expect("write temp encrypted workbook");

    // Missing password should produce a clear, stable error string.
    let err = vba_project_bin::load_vba_project_bin(&path, None)
        .expect_err("expected missing password error");
    assert_eq!(err, "password required for encrypted workbook");

    // Correct password should decrypt and extract `xl/vbaProject.bin`.
    let (decrypted_vba_bin, source) =
        vba_project_bin::load_vba_project_bin(&path, Some(password)).expect("decrypt + extract");
    assert!(
        source.contains("encrypted workbook decrypted"),
        "unexpected source string: {source}"
    );

    // Verify the extracted `vbaProject.bin` matches the original ZIP's entry.
    let mut zip = zip::ZipArchive::new(std::io::Cursor::new(zip_bytes)).expect("valid zip");
    let mut entry = zip
        .by_name("xl/vbaProject.bin")
        .expect("expected vbaProject.bin in fixture");
    let mut expected = Vec::new();
    entry.read_to_end(&mut expected).expect("read vbaProject.bin");

    assert_eq!(decrypted_vba_bin, expected);

    let _ = std::fs::remove_file(&path);
}

