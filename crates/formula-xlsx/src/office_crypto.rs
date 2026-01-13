pub(crate) fn encrypt_package_to_ole(
    package_bytes: &[u8],
    password: &str,
) -> Result<Vec<u8>, formula_office_crypto::OfficeCryptoError> {
    // Use our workspace Office crypto implementation so encryption output matches the
    // standard Office OLE wrapper format (`EncryptionInfo` + `EncryptedPackage`).
    let opts = formula_office_crypto::EncryptOptions::default();
    formula_office_crypto::encrypt_package_to_ole(package_bytes, password, opts)
}
