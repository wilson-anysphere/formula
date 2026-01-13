//! MS-OFFCRYPTO Agile `EncryptionInfo` XML parsing.
//!
//! Modern Excel "Encrypt with Password" workbooks embed an XML document in the `EncryptionInfo`
//! stream (version 4.4). That XML describes:
//! - the cipher + KDF parameters used to encrypt the package payload
//! - one or more `<keyEncryptor>` entries (password, certificate, ...)
//! - optional integrity metadata
//!
//! Excel can emit *multiple* key encryptors (for example, both password and certificate entries).
//! Formula currently supports only password-based key encryption, so this parser selects the first
//! `<keyEncryptor>` whose `@uri` matches the password schema.

use super::{OffCryptoError, Result};

/// Password key encryptor URI as used by MS-OFFCRYPTO Agile EncryptionInfo XML.
pub const KEY_ENCRYPTOR_URI_PASSWORD: &str =
    "http://schemas.microsoft.com/office/2006/keyEncryptor/password";
/// Certificate key encryptor URI as used by MS-OFFCRYPTO Agile EncryptionInfo XML.
pub const KEY_ENCRYPTOR_URI_CERTIFICATE: &str =
    "http://schemas.microsoft.com/office/2006/keyEncryptor/certificate";

/// Warnings produced while parsing `EncryptionInfo` XML.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum EncryptionInfoWarning {
    /// Multiple password `<keyEncryptor>` entries were present.
    ///
    /// Parsing is deterministic: the first password key encryptor wins.
    MultiplePasswordKeyEncryptors { count: usize },
}

/// Parsed key encryptor information for password-based encryption.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PasswordKeyEncryptor {
    /// The `uri` attribute of the selected `<keyEncryptor>` element.
    pub uri: String,
}

/// Parsed Agile `EncryptionInfo` XML (best-effort; currently focused on key encryptor selection).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AgileEncryptionInfo {
    /// The selected password-based key encryptor.
    pub password_key_encryptor: PasswordKeyEncryptor,
    /// Non-fatal parse warnings (deterministic; suitable for telemetry/corpus triage).
    pub warnings: Vec<EncryptionInfoWarning>,
}

/// Parse the XML payload of an Agile (4.4) `EncryptionInfo` stream.
///
/// The caller is responsible for reading the `EncryptionInfo` stream header and providing only the
/// XML bytes.
pub fn parse_agile_encryption_info_xml(xml: &[u8]) -> Result<AgileEncryptionInfo> {
    let xml = std::str::from_utf8(xml)?;
    let doc = roxmltree::Document::parse(xml)?;

    let root = doc.root_element();
    let key_encryptors = root
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "keyEncryptors")
        .ok_or_else(|| OffCryptoError::MissingRequiredElement {
            element: "keyEncryptors".to_string(),
        })?;

    let mut available_uris: Vec<String> = Vec::new();
    let mut password_uri_count = 0usize;
    let mut selected_password_uri: Option<String> = None;

    for enc in key_encryptors
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "keyEncryptor")
    {
        let uri = enc.attribute("uri").ok_or_else(|| OffCryptoError::MissingRequiredAttribute {
            element: "keyEncryptor".to_string(),
            attr: "uri".to_string(),
        })?;

        // Keep a deterministic list of URIs for error reporting. Prefer unique values but preserve
        // first-seen ordering.
        if !available_uris.iter().any(|u| u == uri) {
            available_uris.push(uri.to_string());
        }

        if uri == KEY_ENCRYPTOR_URI_PASSWORD {
            password_uri_count += 1;
            if selected_password_uri.is_none() {
                selected_password_uri = Some(uri.to_string());
            }
        }
    }

    let Some(uri) = selected_password_uri else {
        let mut msg = String::new();
        msg.push_str("unsupported key encryptor in Agile EncryptionInfo: ");
        msg.push_str("Formula currently supports only password-based encryption. ");

        if available_uris.is_empty() {
            msg.push_str("No `<keyEncryptor>` entries were found.");
        } else {
            msg.push_str("Found keyEncryptor URIs: ");
            msg.push_str(&available_uris.join(", "));
            msg.push('.');
        }

        if available_uris.iter().any(|u| u == KEY_ENCRYPTOR_URI_CERTIFICATE) {
            msg.push_str(" This file appears to be certificate-encrypted (public/private key) rather than password-encrypted. Re-save the workbook in Excel using “Encrypt with Password”.");
        } else {
            msg.push_str(" Re-save the workbook in Excel using “Encrypt with Password” (not certificate-based protection).");
        }

        return Err(OffCryptoError::UnsupportedKeyEncryptor {
            available_uris,
            message: msg,
        });
    };

    let mut warnings = Vec::new();
    if password_uri_count > 1 {
        warnings.push(EncryptionInfoWarning::MultiplePasswordKeyEncryptors {
            count: password_uri_count,
        });
    }

    Ok(AgileEncryptionInfo {
        password_key_encryptor: PasswordKeyEncryptor { uri },
        warnings,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn selects_password_key_encryptor_when_multiple_present() {
        let xml = r#"
            <encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
                        xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password"
                        xmlns:c="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate">
              <keyEncryptors>
                <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate">
                  <c:encryptedKey/>
                </keyEncryptor>
                <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
                  <p:encryptedKey spinCount="1"/>
                </keyEncryptor>
              </keyEncryptors>
            </encryption>
        "#;

        let info = parse_agile_encryption_info_xml(xml.as_bytes()).expect("parse should succeed");
        assert_eq!(info.password_key_encryptor.uri, KEY_ENCRYPTOR_URI_PASSWORD);
        assert!(info.warnings.is_empty());
    }

    #[test]
    fn errors_when_password_key_encryptor_missing() {
        let xml = r#"
            <encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
                        xmlns:c="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate">
              <keyEncryptors>
                <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate">
                  <c:encryptedKey/>
                </keyEncryptor>
              </keyEncryptors>
            </encryption>
        "#;

        let err = parse_agile_encryption_info_xml(xml.as_bytes()).expect_err("expected error");
        match &err {
            OffCryptoError::UnsupportedKeyEncryptor { available_uris, .. } => {
                assert!(
                    available_uris
                        .iter()
                        .any(|u| u == KEY_ENCRYPTOR_URI_CERTIFICATE),
                    "expected certificate URI to be reported, got {available_uris:?}"
                );
            }
            other => panic!("expected UnsupportedKeyEncryptor, got {other:?}"),
        }

        let msg = err.to_string();
        assert!(
            msg.contains(KEY_ENCRYPTOR_URI_CERTIFICATE) || msg.to_ascii_lowercase().contains("certificate"),
            "expected error message to mention certificate encryption; got: {msg}"
        );
    }

    #[test]
    fn warns_on_multiple_password_key_encryptors() {
        let xml = r#"
            <encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
                        xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
              <keyEncryptors>
                <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
                  <p:encryptedKey spinCount="1"/>
                </keyEncryptor>
                <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
                  <p:encryptedKey spinCount="2"/>
                </keyEncryptor>
              </keyEncryptors>
            </encryption>
        "#;

        let info = parse_agile_encryption_info_xml(xml.as_bytes()).expect("parse should succeed");
        assert_eq!(info.password_key_encryptor.uri, KEY_ENCRYPTOR_URI_PASSWORD);
        assert_eq!(
            info.warnings,
            vec![EncryptionInfoWarning::MultiplePasswordKeyEncryptors { count: 2 }]
        );
    }
}
