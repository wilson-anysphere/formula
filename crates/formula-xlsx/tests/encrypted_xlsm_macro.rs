use std::path::PathBuf;

fn fixture_path(rel: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures")
        .join(rel)
}

#[test]
fn decrypts_encrypted_xlsm_and_preserves_vba_project() {
    let encrypted =
        std::fs::read(fixture_path("encryption/encrypted_agile.xlsm")).expect("read fixture");
    let decrypted = formula_office_crypto::decrypt_encrypted_package(&encrypted, "password")
        .expect("decrypt encrypted_agile.xlsm");

    let package = formula_xlsx::XlsxPackage::from_bytes(&decrypted).expect("open decrypted xlsm");
    let vba = package
        .vba_project_bin()
        .expect("expected xl/vbaProject.bin to exist in xlsm package");

    // Ensure the macro project is structurally valid and can be parsed by `formula-vba`.
    formula_vba::VBAProject::parse(vba).expect("parse vbaProject.bin");
}
