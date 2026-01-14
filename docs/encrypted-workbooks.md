# Encrypted / Password‑Protected Excel Workbooks

This page is a short entrypoint for Formula’s **file-level workbook encryption** (“Password to
open”) support.

The canonical, detailed documentation lives in:

- [`docs/21-encrypted-workbooks.md`](./21-encrypted-workbooks.md) — overview, support matrix, APIs,
  error semantics, security notes
- [`docs/21-offcrypto.md`](./21-offcrypto.md) — MS‑OFFCRYPTO “what the file looks like” + scheme detection + `formula-io` password APIs
- [`docs/22-ooxml-encryption.md`](./22-ooxml-encryption.md) — Agile (4.4) OOXML decryption details (HMAC target bytes, IV/salt gotchas)
- [`docs/office-encryption.md`](./office-encryption.md) — maintainer-level reference (supported parameter subsets, KDF nuances, writer defaults)

## Quickstart (Rust)

Prefer `formula-io` for format detection + password handling:

```rust
use formula_io::{open_workbook_with_options, Error, OpenOptions};

let path = "encrypted.xlsx";

match open_workbook_with_options(
    path,
    OpenOptions {
        password: Some("password".to_string()),
        ..Default::default()
    },
) {
    Ok(wb) => {
        let _ = wb;
    }
    Err(Error::PasswordRequired { .. }) => {
        // Prompt the user and retry with a password.
    }
    Err(Error::InvalidPassword { .. }) => {
        // Wrong password (or integrity mismatch for some OOXML schemes).
    }
    Err(err) => return Err(err),
}
```

## UX reminder (Desktop)

When the desktop app hits `PasswordRequired`, it should prompt the user for a password and retry
the open request. Passwords should not be persisted or logged.
