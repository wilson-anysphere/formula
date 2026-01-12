use serde_json::Value as JsonValue;

/// Maximum size of the workbook-backed `xl/formula/power-query.xml` payload (UTF-8 bytes).
///
/// This guards against a compromised webview (or buggy frontend) sending arbitrarily large XML
/// blobs that would otherwise be kept in memory and persisted into workbooks.
pub const MAX_POWER_QUERY_XML_BYTES: usize = 2 * 1024 * 1024; // 2 MiB

/// Maximum length of the Power Query credential "scope key".
///
/// Scope keys are persisted as document IDs in the encrypted credential store.
pub const MAX_CREDENTIAL_SCOPE_KEY_LEN: usize = 512;

/// Maximum size of the Power Query credential `secret` payload when serialized as JSON bytes.
pub const MAX_CREDENTIAL_SECRET_BYTES: usize = 64 * 1024; // 64 KiB

/// Maximum size of the Power Query refresh-state payload when serialized as JSON bytes.
pub const MAX_REFRESH_STATE_BYTES: usize = 1024 * 1024; // 1 MiB

/// Maximum number of top-level refresh-state entries (typically one per query / schedule).
pub const MAX_REFRESH_STATE_ENTRIES: usize = 10_000;

#[derive(Debug, thiserror::Error, PartialEq, Eq)]
pub enum PowerQueryValidationError {
    #[error("Payload too large")]
    PayloadTooLarge,
    #[error("Failed to serialize JSON")]
    JsonSerialize,
}

fn json_within_byte_limit(
    value: &JsonValue,
    max_bytes: usize,
) -> Result<(), PowerQueryValidationError> {
    struct ByteLimitWriter {
        remaining: usize,
    }

    impl std::io::Write for ByteLimitWriter {
        fn write(&mut self, buf: &[u8]) -> std::io::Result<usize> {
            if buf.len() > self.remaining {
                return Err(std::io::Error::new(
                    std::io::ErrorKind::Other,
                    "payload too large",
                ));
            }
            self.remaining -= buf.len();
            Ok(buf.len())
        }

        fn flush(&mut self) -> std::io::Result<()> {
            Ok(())
        }
    }

    // First, do a bounded serialization to avoid allocating a Vec<u8> for an untrusted payload.
    // If that passes, we still run `serde_json::to_vec` (as a deterministic byte-size check)
    // knowing it will only allocate up to `max_bytes`.
    let mut writer = ByteLimitWriter {
        remaining: max_bytes,
    };
    match serde_json::to_writer(&mut writer, value) {
        Ok(()) => {
            let bytes =
                serde_json::to_vec(value).map_err(|_| PowerQueryValidationError::JsonSerialize)?;
            if bytes.len() > max_bytes {
                return Err(PowerQueryValidationError::PayloadTooLarge);
            }
            Ok(())
        }
        Err(err) if err.is_io() => Err(PowerQueryValidationError::PayloadTooLarge),
        Err(_) => Err(PowerQueryValidationError::JsonSerialize),
    }
}

pub fn validate_power_query_xml_payload(xml: &str) -> Result<(), PowerQueryValidationError> {
    if xml.as_bytes().len() > MAX_POWER_QUERY_XML_BYTES {
        return Err(PowerQueryValidationError::PayloadTooLarge);
    }
    Ok(())
}

pub fn validate_power_query_credential_payload(
    scope_key: &str,
    secret: &JsonValue,
) -> Result<(), PowerQueryValidationError> {
    if scope_key.len() > MAX_CREDENTIAL_SCOPE_KEY_LEN {
        return Err(PowerQueryValidationError::PayloadTooLarge);
    }
    json_within_byte_limit(secret, MAX_CREDENTIAL_SECRET_BYTES)?;

    Ok(())
}

pub fn validate_power_query_refresh_state_payload(
    state: &JsonValue,
) -> Result<(), PowerQueryValidationError> {
    if let JsonValue::Object(map) = state {
        if map.len() > MAX_REFRESH_STATE_ENTRIES {
            return Err(PowerQueryValidationError::PayloadTooLarge);
        }
    }
    json_within_byte_limit(state, MAX_REFRESH_STATE_BYTES)?;

    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    use serde_json::json;

    #[test]
    fn power_query_xml_payload_rejects_oversized_payloads() {
        let ok = "a".repeat(MAX_POWER_QUERY_XML_BYTES);
        assert!(validate_power_query_xml_payload(&ok).is_ok());

        let too_big = "a".repeat(MAX_POWER_QUERY_XML_BYTES + 1);
        assert_eq!(
            validate_power_query_xml_payload(&too_big).unwrap_err(),
            PowerQueryValidationError::PayloadTooLarge
        );
    }

    #[test]
    fn credential_payload_rejects_long_scope_keys() {
        let scope_key = "a".repeat(MAX_CREDENTIAL_SCOPE_KEY_LEN + 1);
        let secret = json!({"password": "ok"});
        assert_eq!(
            validate_power_query_credential_payload(&scope_key, &secret).unwrap_err(),
            PowerQueryValidationError::PayloadTooLarge
        );
    }

    #[test]
    fn credential_payload_rejects_large_secrets() {
        // JSON serialization adds 2 bytes for the surrounding quotes.
        let ok_secret =
            JsonValue::String("a".repeat(MAX_CREDENTIAL_SECRET_BYTES.saturating_sub(2)));
        assert!(validate_power_query_credential_payload("scope", &ok_secret).is_ok());

        let too_big_secret =
            JsonValue::String("a".repeat(MAX_CREDENTIAL_SECRET_BYTES.saturating_sub(1)));
        assert_eq!(
            validate_power_query_credential_payload("scope", &too_big_secret).unwrap_err(),
            PowerQueryValidationError::PayloadTooLarge
        );
    }

    #[test]
    fn refresh_state_payload_rejects_too_many_entries() {
        let mut map = serde_json::Map::with_capacity(MAX_REFRESH_STATE_ENTRIES + 1);
        for i in 0..(MAX_REFRESH_STATE_ENTRIES + 1) {
            map.insert(
                format!("k{i}"),
                json!({"policy": {"type": "interval", "intervalMs": 1}}),
            );
        }
        let state = JsonValue::Object(map);
        assert_eq!(
            validate_power_query_refresh_state_payload(&state).unwrap_err(),
            PowerQueryValidationError::PayloadTooLarge
        );
    }

    #[test]
    fn refresh_state_payload_rejects_oversized_payloads() {
        let ok = JsonValue::String("a".repeat(MAX_REFRESH_STATE_BYTES.saturating_sub(2)));
        assert!(validate_power_query_refresh_state_payload(&ok).is_ok());

        let too_big = JsonValue::String("a".repeat(MAX_REFRESH_STATE_BYTES.saturating_sub(1)));
        assert_eq!(
            validate_power_query_refresh_state_payload(&too_big).unwrap_err(),
            PowerQueryValidationError::PayloadTooLarge
        );
    }
}
