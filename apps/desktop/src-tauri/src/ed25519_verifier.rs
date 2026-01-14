use base64::{engine::general_purpose::STANDARD, Engine as _};
use ed25519_dalek::{Signature, VerifyingKey};
use pkcs8::DecodePublicKey;
use serde::{de, Deserialize};
use std::fmt;

#[cfg(feature = "desktop")]
use crate::resource_limits::LimitedString;

// NOTE: Keep these limits in sync with the browser verifier in
// `shared/extension-package/v2-browser.mjs`.
const MAX_SIGNATURE_PAYLOAD_BYTES: usize = 5 * 1024 * 1024; // 5MB
const MAX_PUBLIC_KEY_PEM_BYTES: usize = 64 * 1024; // 64KB
const MAX_SIGNATURE_BASE64_BYTES: usize = 1024;

/// IPC-deserialized byte array with a maximum length enforced during deserialization.
///
/// This is a defense-in-depth guard to prevent a compromised webview from attempting to allocate an
/// unbounded `Vec<u8>` via a giant JSON array before we can validate the payload size.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct LimitedByteVec<const MAX: usize>(pub Vec<u8>);

impl<'de, const MAX: usize> Deserialize<'de> for LimitedByteVec<MAX> {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: serde::Deserializer<'de>,
    {
        struct LimitedByteVecVisitor<const MAX: usize>;

        impl<'de, const MAX: usize> de::Visitor<'de> for LimitedByteVecVisitor<MAX> {
            type Value = LimitedByteVec<MAX>;

            fn expecting(&self, formatter: &mut fmt::Formatter) -> fmt::Result {
                formatter.write_str("an array of bytes")
            }

            fn visit_seq<A>(self, mut seq: A) -> Result<Self::Value, A::Error>
            where
                A: de::SeqAccess<'de>,
            {
                let hint = seq.size_hint();
                if let Some(hint) = hint {
                    if hint > MAX {
                        return Err(de::Error::custom(format!(
                            "Payload is too large (max {MAX} bytes)"
                        )));
                    }
                }

                let mut out = match hint {
                    Some(hint) => Vec::with_capacity(hint.min(MAX)),
                    None => Vec::new(),
                };

                for _ in 0..MAX {
                    match seq.next_element::<u8>()? {
                        Some(v) => out.push(v),
                        None => return Ok(LimitedByteVec(out)),
                    }
                }

                // Detect overflow without allocating/parsing another typed element.
                if seq.next_element::<de::IgnoredAny>()?.is_some() {
                    return Err(de::Error::custom(format!(
                        "Payload is too large (max {MAX} bytes)"
                    )));
                }
                Ok(LimitedByteVec(out))
            }
        }

        deserializer.deserialize_seq(LimitedByteVecVisitor::<MAX>)
    }
}

pub fn verify_ed25519_signature_payload(
    payload: &[u8],
    signature_base64: &str,
    public_key_pem: &str,
) -> Result<bool, String> {
    if payload.len() > MAX_SIGNATURE_PAYLOAD_BYTES {
        return Err(format!(
            "Payload is too large (max {MAX_SIGNATURE_PAYLOAD_BYTES} bytes)"
        ));
    }

    if public_key_pem.len() > MAX_PUBLIC_KEY_PEM_BYTES {
        return Err(format!(
            "Public key PEM is too large (max {MAX_PUBLIC_KEY_PEM_BYTES} bytes)"
        ));
    }

    if signature_base64.len() > MAX_SIGNATURE_BASE64_BYTES {
        return Err(format!(
            "Signature base64 is too large (max {MAX_SIGNATURE_BASE64_BYTES} bytes)"
        ));
    }

    let signature_bytes = STANDARD
        .decode(signature_base64.trim())
        .map_err(|err| format!("Invalid signature base64: {err}"))?;

    if signature_bytes.len() != 64 {
        return Err(format!(
            "Invalid signature length: expected 64 bytes, got {}",
            signature_bytes.len()
        ));
    }

    let signature = {
        let sig: [u8; 64] = signature_bytes
            .as_slice()
            .try_into()
            .map_err(|_| "Invalid signature length".to_string())?;
        Signature::from_bytes(&sig)
    };

    let public_key = VerifyingKey::from_public_key_pem(public_key_pem.trim())
        .map_err(|err| format!("Invalid Ed25519 public key PEM: {err}"))?;

    match public_key.verify_strict(payload, &signature) {
        Ok(()) => Ok(true),
        Err(_) => Ok(false),
    }
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub fn verify_ed25519_signature(
    window: tauri::WebviewWindow,
    payload: LimitedByteVec<MAX_SIGNATURE_PAYLOAD_BYTES>,
    signature_base64: LimitedString<MAX_SIGNATURE_BASE64_BYTES>,
    public_key_pem: LimitedString<MAX_PUBLIC_KEY_PEM_BYTES>,
) -> Result<bool, String> {
    use crate::ipc_origin::Verb;

    let subject = "ed25519 verification";
    crate::ipc_origin::ensure_main_window(window.label(), subject, Verb::Is)?;
    crate::ipc_origin::ensure_stable_origin(&window, subject, Verb::Is)?;
    let url = window.url().map_err(|err| err.to_string())?;
    crate::ipc_origin::ensure_trusted_origin(&url, subject, Verb::Is)?;

    verify_ed25519_signature_payload(&payload.0, signature_base64.as_ref(), public_key_pem.as_ref())
}

#[cfg(test)]
mod tests {
    use super::*;
    use base64::engine::general_purpose::STANDARD;
    use ed25519_dalek::{Signer, SigningKey};
    use pkcs8::{EncodePublicKey, LineEnding};

    const BROWSER_VERIFIER_JS_PATH: &str = "shared/extension-package/v2-browser.mjs";
    const BROWSER_VERIFIER_JS_SOURCE: &str = include_str!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../../shared/extension-package/v2-browser.mjs"
    ));

    fn test_keypair() -> (SigningKey, String) {
        let signing_key = SigningKey::from_bytes(&[7u8; 32]);
        let public_key_pem = signing_key
            .verifying_key()
            .to_public_key_pem(LineEnding::LF)
            .expect("expected public key PEM encoding to succeed");
        (signing_key, public_key_pem)
    }

    #[test]
    fn valid_signature_returns_true() {
        let (signing_key, public_key_pem) = test_keypair();
        let payload = b"hello world".to_vec();
        let sig = signing_key.sign(&payload);
        let signature_base64 = STANDARD.encode(sig.to_bytes());

        let ok = verify_ed25519_signature_payload(&payload, &signature_base64, &public_key_pem)
            .expect("expected verification to succeed");
        assert!(ok);
    }

    #[test]
    fn tampered_payload_returns_false() {
        let (signing_key, public_key_pem) = test_keypair();
        let payload = b"hello world".to_vec();
        let sig = signing_key.sign(&payload);
        let signature_base64 = STANDARD.encode(sig.to_bytes());

        let mut tampered = payload.clone();
        tampered[0] ^= 0x01;

        let ok = verify_ed25519_signature_payload(&tampered, &signature_base64, &public_key_pem)
            .expect("expected verification to succeed");
        assert!(!ok);
    }

    #[test]
    fn invalid_base64_returns_err() {
        let (_signing_key, public_key_pem) = test_keypair();
        let payload = b"hello world".to_vec();
        let err = verify_ed25519_signature_payload(&payload, "not base64!!!", &public_key_pem)
            .expect_err("expected invalid base64 to return Err");
        assert!(
            err.to_ascii_lowercase().contains("base64"),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn invalid_pem_returns_err() {
        let (signing_key, _public_key_pem) = test_keypair();
        let payload = b"hello world".to_vec();
        let sig = signing_key.sign(&payload);
        let signature_base64 = STANDARD.encode(sig.to_bytes());

        let err = verify_ed25519_signature_payload(&payload, &signature_base64, "not a pem")
            .expect_err("expected invalid PEM to return Err");
        assert!(
            err.to_ascii_lowercase().contains("pem"),
            "unexpected error: {err}"
        );
    }

    struct SizeHintSeqDeserializer {
        len: usize,
    }

    struct SizeHintSeqAccess {
        len: usize,
    }

    impl<'de> de::SeqAccess<'de> for SizeHintSeqAccess {
        type Error = de::value::Error;

        fn next_element_seed<T>(&mut self, _seed: T) -> Result<Option<T::Value>, Self::Error>
        where
            T: de::DeserializeSeed<'de>,
        {
            panic!("unexpected element deserialization (size_hint guard should have failed first)");
        }

        fn size_hint(&self) -> Option<usize> {
            Some(self.len)
        }
    }

    impl<'de> serde::Deserializer<'de> for SizeHintSeqDeserializer {
        type Error = de::value::Error;

        fn deserialize_any<V>(self, visitor: V) -> Result<V::Value, Self::Error>
        where
            V: de::Visitor<'de>,
        {
            self.deserialize_seq(visitor)
        }

        fn deserialize_seq<V>(self, visitor: V) -> Result<V::Value, Self::Error>
        where
            V: de::Visitor<'de>,
        {
            visitor.visit_seq(SizeHintSeqAccess { len: self.len })
        }

        serde::forward_to_deserialize_any! {
            bool i8 i16 i32 i64 u8 u16 u32 u64 f32 f64 char str string bytes byte_buf option unit
            unit_struct newtype_struct tuple tuple_struct map struct enum identifier ignored_any
        }
    }

    #[test]
    fn limited_byte_vec_rejects_oversized_size_hint() {
        type SmallBytes = LimitedByteVec<4>;
        let err = <SmallBytes as Deserialize>::deserialize(SizeHintSeqDeserializer { len: 5 })
            .expect_err("expected size_hint guard to reject oversized byte vec")
            .to_string();
        assert!(
            err.contains("max") && err.contains("4"),
            "unexpected error message: {err}"
        );
    }

    #[test]
    fn limited_byte_vec_rejects_oversized_json_array() {
        type SmallBytes = LimitedByteVec<4>;
        let err = serde_json::from_str::<SmallBytes>("[0,1,2,3,4]")
            .expect_err("expected oversized byte array to be rejected")
            .to_string();
        assert!(
            err.contains("max") && err.contains("4"),
            "unexpected error message: {err}"
        );
    }

    #[test]
    fn invalid_signature_length_returns_err() {
        let (_signing_key, public_key_pem) = test_keypair();
        let payload = b"hello world".to_vec();
        let signature_base64 = STANDARD.encode([0u8; 10]);

        let err = verify_ed25519_signature_payload(&payload, &signature_base64, &public_key_pem)
            .expect_err("expected invalid signature length to return Err");
        assert!(
            err.to_ascii_lowercase().contains("signature length"),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn oversized_payload_returns_err() {
        let payload = vec![0u8; MAX_SIGNATURE_PAYLOAD_BYTES + 1];
        let err = verify_ed25519_signature_payload(&payload, "", "")
            .expect_err("expected oversized payload to return Err");
        assert!(
            err.to_ascii_lowercase().contains("payload is too large"),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn oversized_public_key_pem_returns_err() {
        let payload = b"hello world".to_vec();
        let oversized_pem = "A".repeat(MAX_PUBLIC_KEY_PEM_BYTES + 1);

        let err = verify_ed25519_signature_payload(&payload, "AA==", &oversized_pem)
            .expect_err("expected oversized public key PEM to return Err");
        assert!(
            err.to_ascii_lowercase().contains("public key pem is too large"),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn oversized_signature_base64_returns_err() {
        let payload = b"hello world".to_vec();
        let oversized_sig = "A".repeat(MAX_SIGNATURE_BASE64_BYTES + 1);

        let err = verify_ed25519_signature_payload(&payload, &oversized_sig, "")
            .expect_err("expected oversized signature base64 to return Err");
        assert!(
            err.to_ascii_lowercase().contains("signature base64 is too large"),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn browser_and_desktop_limits_stay_in_sync() {
        let js_payload = js_const_usize(BROWSER_VERIFIER_JS_SOURCE, "MAX_SIGNATURE_PAYLOAD_BYTES")
            .unwrap_or_else(|err| panic!("{err}"));
        let js_pem = js_const_usize(BROWSER_VERIFIER_JS_SOURCE, "MAX_PUBLIC_KEY_PEM_BYTES")
            .unwrap_or_else(|err| panic!("{err}"));
        let js_sig_b64 = js_const_usize(BROWSER_VERIFIER_JS_SOURCE, "MAX_SIGNATURE_BASE64_BYTES")
            .unwrap_or_else(|err| panic!("{err}"));

        assert_eq!(
            MAX_SIGNATURE_PAYLOAD_BYTES, js_payload,
            "Ed25519 verifier limit drift: MAX_SIGNATURE_PAYLOAD_BYTES differs between Rust ({MAX_SIGNATURE_PAYLOAD_BYTES}) \
             and {BROWSER_VERIFIER_JS_PATH} ({js_payload}). Update both sides to match."
        );
        assert_eq!(
            MAX_PUBLIC_KEY_PEM_BYTES, js_pem,
            "Ed25519 verifier limit drift: MAX_PUBLIC_KEY_PEM_BYTES differs between Rust ({MAX_PUBLIC_KEY_PEM_BYTES}) \
             and {BROWSER_VERIFIER_JS_PATH} ({js_pem}). Update both sides to match."
        );
        assert_eq!(
            MAX_SIGNATURE_BASE64_BYTES, js_sig_b64,
            "Ed25519 verifier limit drift: MAX_SIGNATURE_BASE64_BYTES differs between Rust ({MAX_SIGNATURE_BASE64_BYTES}) \
             and {BROWSER_VERIFIER_JS_PATH} ({js_sig_b64}). Update both sides to match."
        );
    }

    fn js_const_usize(source: &str, name: &str) -> Result<usize, String> {
        let expr = extract_js_const_expression(source, name)?;
        parse_js_usize_expr(expr).map_err(|err| {
            format!(
                "Failed to parse `{name}` in {BROWSER_VERIFIER_JS_PATH}.\n\
                 Found expression: `{expr}`\n\
                 Error: {err}\n\
                 Expected a top-level declaration like:\n\
                   const {name} = 5 * 1024 * 1024;\n\
                 (only integer literals, `*`, and parentheses are supported)."
            )
        })
    }

    fn extract_js_const_expression<'a>(source: &'a str, name: &str) -> Result<&'a str, String> {
        let bytes = source.as_bytes();
        let mut i = 0usize;
        while i < bytes.len() {
            skip_js_ws_and_comments(source, &mut i)?;
            if i >= bytes.len() {
                break;
            }

            match bytes[i] {
                b'\'' | b'"' => {
                    let quote = bytes[i];
                    skip_js_string(source, &mut i, quote)?;
                }
                b'`' => {
                    skip_js_template_literal(source, &mut i)?;
                }
                b if is_js_ident_start(b) => {
                    let ident = parse_js_ident(source, &mut i);
                    if ident != "const" {
                        continue;
                    }

                    skip_js_ws_and_comments(source, &mut i)?;
                    if i >= bytes.len() {
                        break;
                    }

                    // Skip destructuring `const { ... } = ...;` / `const [ ... ] = ...;`
                    if bytes[i] == b'{' || bytes[i] == b'[' {
                        let semi = find_js_statement_terminator(source, i)?;
                        i = semi + 1;
                        continue;
                    }

                    if !is_js_ident_start(bytes[i]) {
                        continue;
                    }
                    let var_name = parse_js_ident(source, &mut i);
                    if var_name != name {
                        // Skip the rest of this `const` statement so we don't accidentally treat
                        // a later identifier as a new declaration.
                        let semi = find_js_statement_terminator(source, i)?;
                        i = semi + 1;
                        continue;
                    }

                    skip_js_ws_and_comments(source, &mut i)?;
                    if i >= bytes.len() || bytes[i] != b'=' {
                        let line = 1 + bytes[..i].iter().filter(|b| **b == b'\n').count();
                        return Err(format!(
                            "Found `const {name}` in {BROWSER_VERIFIER_JS_PATH} (line {line}) but did not find `=`. \
                             Expected `const {name} = <expr>;`."
                        ));
                    }
                    i += 1; // '='

                    let expr_start = i;
                    let semi = find_js_statement_terminator(source, i)?;
                    let expr = source[expr_start..semi].trim();
                    if expr.is_empty() {
                        let line = 1 + bytes[..expr_start].iter().filter(|b| **b == b'\n').count();
                        return Err(format!(
                            "Found `const {name}` in {BROWSER_VERIFIER_JS_PATH} (line {line}) but the assigned expression was empty."
                        ));
                    }
                    return Ok(expr);
                }
                _ => {
                    i += 1;
                }
            }
        }

        Err(format!(
            "Failed to find `{name}` constant in {BROWSER_VERIFIER_JS_PATH}.\n\
             Expected a top-level declaration like:\n\
               const {name} = 5 * 1024 * 1024;\n\
             If the browser verifier was refactored, update this test parser accordingly so Rust and JS limits stay in sync."
        ))
    }

    fn skip_js_ws_and_comments(source: &str, i: &mut usize) -> Result<(), String> {
        let bytes = source.as_bytes();
        while *i < bytes.len() {
            match bytes[*i] {
                b' ' | b'\t' | b'\n' | b'\r' => *i += 1,
                b'/' if *i + 1 < bytes.len() && bytes[*i + 1] == b'/' => {
                    *i += 2;
                    while *i < bytes.len() && bytes[*i] != b'\n' {
                        *i += 1;
                    }
                }
                b'/' if *i + 1 < bytes.len() && bytes[*i + 1] == b'*' => {
                    *i += 2;
                    let mut found = false;
                    while *i + 1 < bytes.len() {
                        if bytes[*i] == b'*' && bytes[*i + 1] == b'/' {
                            *i += 2;
                            found = true;
                            break;
                        }
                        *i += 1;
                    }
                    if !found {
                        return Err(format!(
                            "Unterminated block comment while parsing {BROWSER_VERIFIER_JS_PATH}"
                        ));
                    }
                }
                _ => break,
            }
        }
        Ok(())
    }

    fn skip_js_string(source: &str, i: &mut usize, quote: u8) -> Result<(), String> {
        let bytes = source.as_bytes();
        debug_assert!(quote == b'\'' || quote == b'"');
        *i += 1; // opening quote
        while *i < bytes.len() {
            match bytes[*i] {
                b'\\' => *i = (*i + 2).min(bytes.len()),
                b if b == quote => {
                    *i += 1;
                    return Ok(());
                }
                _ => *i += 1,
            }
        }
        Err(format!(
            "Unterminated string literal while parsing {BROWSER_VERIFIER_JS_PATH}"
        ))
    }

    fn skip_js_template_literal(source: &str, i: &mut usize) -> Result<(), String> {
        let bytes = source.as_bytes();
        debug_assert!(bytes.get(*i) == Some(&b'`'));
        *i += 1; // opening backtick
        while *i < bytes.len() {
            match bytes[*i] {
                b'\\' => *i = (*i + 2).min(bytes.len()),
                b'`' => {
                    *i += 1;
                    return Ok(());
                }
                _ => *i += 1,
            }
        }
        Err(format!(
            "Unterminated template literal while parsing {BROWSER_VERIFIER_JS_PATH}"
        ))
    }

    fn find_js_statement_terminator(source: &str, mut i: usize) -> Result<usize, String> {
        let bytes = source.as_bytes();
        while i < bytes.len() {
            match bytes[i] {
                b';' => return Ok(i),
                b'\'' | b'"' => {
                    let quote = bytes[i];
                    skip_js_string(source, &mut i, quote)?;
                }
                b'`' => {
                    skip_js_template_literal(source, &mut i)?;
                }
                b'/' if i + 1 < bytes.len() && bytes[i + 1] == b'/' => {
                    i += 2;
                    while i < bytes.len() && bytes[i] != b'\n' {
                        i += 1;
                    }
                }
                b'/' if i + 1 < bytes.len() && bytes[i + 1] == b'*' => {
                    i += 2;
                    let mut found = false;
                    while i + 1 < bytes.len() {
                        if bytes[i] == b'*' && bytes[i + 1] == b'/' {
                            i += 2;
                            found = true;
                            break;
                        }
                        i += 1;
                    }
                    if !found {
                        return Err(format!(
                            "Unterminated block comment while parsing {BROWSER_VERIFIER_JS_PATH}"
                        ));
                    }
                }
                _ => i += 1,
            }
        }
        Err(format!(
            "Failed to find statement terminator `;` while parsing {BROWSER_VERIFIER_JS_PATH}"
        ))
    }

    fn is_js_ident_start(b: u8) -> bool {
        b.is_ascii_alphabetic() || b == b'_' || b == b'$'
    }

    fn is_js_ident_continue(b: u8) -> bool {
        is_js_ident_start(b) || b.is_ascii_digit()
    }

    fn parse_js_ident<'a>(source: &'a str, i: &mut usize) -> &'a str {
        let bytes = source.as_bytes();
        let start = *i;
        *i += 1;
        while *i < bytes.len() && is_js_ident_continue(bytes[*i]) {
            *i += 1;
        }
        &source[start..*i]
    }

    fn parse_js_usize_expr(expr: &str) -> Result<usize, String> {
        let expr = strip_js_comments(expr)?;
        let tokens = tokenize_js_mul_expr(&expr)?;
        let mut idx = 0usize;
        let value = parse_js_mul_expr(&tokens, &mut idx)?;
        if idx != tokens.len() {
            return Err(format!(
                "Unexpected trailing tokens after parsing expression: `{}`",
                &expr
            ));
        }
        usize::try_from(value).map_err(|_| "Value does not fit in usize".to_string())
    }

    fn strip_js_comments(input: &str) -> Result<String, String> {
        let bytes = input.as_bytes();
        let mut out = String::with_capacity(input.len());
        let mut i = 0usize;
        while i < bytes.len() {
            if bytes[i] == b'/' && i + 1 < bytes.len() && bytes[i + 1] == b'/' {
                i += 2;
                while i < bytes.len() && bytes[i] != b'\n' {
                    i += 1;
                }
                continue;
            }
            if bytes[i] == b'/' && i + 1 < bytes.len() && bytes[i + 1] == b'*' {
                i += 2;
                let mut found = false;
                while i + 1 < bytes.len() {
                    if bytes[i] == b'*' && bytes[i + 1] == b'/' {
                        i += 2;
                        found = true;
                        break;
                    }
                    i += 1;
                }
                if !found {
                    return Err("Unterminated block comment in expression".to_string());
                }
                continue;
            }
            out.push(bytes[i] as char);
            i += 1;
        }
        Ok(out)
    }

    #[derive(Debug, Clone, Copy)]
    enum JsTok {
        Num(u64),
        Star,
        LParen,
        RParen,
    }

    fn tokenize_js_mul_expr(expr: &str) -> Result<Vec<JsTok>, String> {
        let bytes = expr.as_bytes();
        let mut tokens = Vec::new();
        let mut i = 0usize;
        while i < bytes.len() {
            match bytes[i] {
                b' ' | b'\t' | b'\n' | b'\r' => {
                    i += 1;
                }
                b'(' => {
                    tokens.push(JsTok::LParen);
                    i += 1;
                }
                b')' => {
                    tokens.push(JsTok::RParen);
                    i += 1;
                }
                b'*' => {
                    if i + 1 < bytes.len() && bytes[i + 1] == b'*' {
                        return Err("Unsupported operator `**` (use explicit `*` multiplications)".to_string());
                    }
                    tokens.push(JsTok::Star);
                    i += 1;
                }
                b'0'..=b'9' => {
                    let start = i;
                    i += 1;
                    while i < bytes.len() {
                        match bytes[i] {
                            b'0'..=b'9' | b'_' => i += 1,
                            _ => break,
                        }
                    }
                    let raw = &expr[start..i];
                    let normalized: String = raw.chars().filter(|c| *c != '_').collect();
                    let value = normalized
                        .parse::<u64>()
                        .map_err(|err| format!("Invalid integer literal `{raw}`: {err}"))?;
                    tokens.push(JsTok::Num(value));
                }
                other => {
                    return Err(format!(
                        "Unexpected character `{}` in expression `{expr}`",
                        other as char
                    ));
                }
            }
        }
        Ok(tokens)
    }

    fn parse_js_mul_expr(tokens: &[JsTok], idx: &mut usize) -> Result<u64, String> {
        let mut value = parse_js_term(tokens, idx)?;
        while matches!(tokens.get(*idx), Some(JsTok::Star)) {
            *idx += 1; // '*'
            let rhs = parse_js_term(tokens, idx)?;
            value = value
                .checked_mul(rhs)
                .ok_or_else(|| "Expression overflowed u64".to_string())?;
        }
        Ok(value)
    }

    fn parse_js_term(tokens: &[JsTok], idx: &mut usize) -> Result<u64, String> {
        match tokens.get(*idx).copied() {
            Some(JsTok::Num(n)) => {
                *idx += 1;
                Ok(n)
            }
            Some(JsTok::LParen) => {
                *idx += 1;
                let value = parse_js_mul_expr(tokens, idx)?;
                match tokens.get(*idx) {
                    Some(JsTok::RParen) => {
                        *idx += 1;
                        Ok(value)
                    }
                    _ => Err("Expected `)`".to_string()),
                }
            }
            _ => Err("Expected an integer literal or parenthesized expression".to_string()),
        }
    }
}
