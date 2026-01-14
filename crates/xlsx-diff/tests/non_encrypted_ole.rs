use std::path::Path;

use anyhow::Result;

#[test]
fn non_encrypted_ole_does_not_require_password() -> Result<()> {
    let fixture_path = Path::new(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../formula-xls/tests/fixtures/basic.xls"
    ));

    let err = match xlsx_diff::WorkbookArchive::open_with_password(fixture_path, None) {
        Ok(_) => panic!("expected .xls (OLE) file to be rejected"),
        Err(err) => err,
    };
    let msg = err.to_string().to_ascii_lowercase();

    assert!(
        msg.contains("ole compound file") && msg.contains("encryptedpackage"),
        "expected OLE/non-encryptedpackage error, got: {msg}"
    );
    assert!(
        !msg.contains("provide a password"),
        "should not request a password for non-encrypted OLE files, got: {msg}"
    );

    Ok(())
}
