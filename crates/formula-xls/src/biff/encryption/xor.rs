//! BIFF XOR obfuscation (legacy) password verifier helpers.
//!
//! BIFF streams can use the XOR obfuscation scheme (the weakest "encryption" option). The
//! `FILEPASS` record stores a 16-bit password verifier value that is computed using the same
//! legacy Excel hash function used by worksheet/workbook protection.
//!
//! This module intentionally only exposes the deterministic verifier computation (not full stream
//! obfuscation), as it is useful for unit testing and for validating `FILEPASS` payloads.

/// Compute the legacy XOR password verifier (16-bit).
///
/// Note: This matches the algorithm used for Excel's legacy sheet/workbook protection password
/// hash. It is **not** cryptographically secure.
pub(crate) fn xor_password_verifier(password: &str) -> u16 {
    formula_model::hash_legacy_password(password)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn xor_password_verifier_matches_published_vectors() {
        // Published examples (Excel legacy XOR verifier).
        // See e.g. XlsxWriter/OpenXML docs.
        let cases = [
            ("password", 0x83AF),
            ("test", 0xCBEB),
            ("1234", 0xCC3D),
            ("", 0xCE4B),
        ];

        for (pw, expected) in cases {
            assert_eq!(xor_password_verifier(pw), expected, "pw={pw:?}");
        }
    }
}

