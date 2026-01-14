use std::path::PathBuf;

use formula_io::Error;

#[test]
fn password_error_messages_are_actionable() {
    let path = PathBuf::from("book.xlsx");

    let required = Error::PasswordRequired { path: path.clone() };
    let required_msg = required.to_string().to_lowercase();
    assert!(
        required_msg.contains("password required"),
        "expected password-required error to mention that a password is required; got: {required_msg}"
    );
    assert!(
        required_msg.contains("open_workbook_with_password"),
        "expected password-required error to hint how to provide a password; got: {required_msg}"
    );
    assert!(
        !required_msg.contains("remove password protection"),
        "password-required error should not ask users to remove password protection; got: {required_msg}"
    );

    let invalid = Error::InvalidPassword { path };
    let invalid_msg = invalid.to_string().to_lowercase();
    assert!(
        invalid_msg.contains("invalid password"),
        "expected invalid-password error to mention that the password is invalid; got: {invalid_msg}"
    );
    assert!(
        !invalid_msg.contains("remove password protection"),
        "invalid-password error should not ask users to remove password protection; got: {invalid_msg}"
    );

    let xls = Error::EncryptedWorkbook {
        path: PathBuf::from("legacy.xls"),
    };
    let xls_msg = xls.to_string().to_lowercase();
    assert!(
        xls_msg.contains("password required"),
        "expected legacy encrypted-workbook error to mention that a password is required; got: {xls_msg}"
    );
    assert!(
        xls_msg.contains("open_workbook_with_password"),
        "expected legacy encrypted-workbook error to hint how to provide a password; got: {xls_msg}"
    );
    assert!(
        !xls_msg.contains("remove password protection"),
        "legacy encrypted-workbook error should not ask users to remove password protection; got: {xls_msg}"
    );
}
