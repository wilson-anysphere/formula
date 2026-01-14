# Encrypted OOXML fixtures (`.xlsx` / `.xlsm`)

This directory contains **password-protected OOXML workbooks** (`.xlsx`, `.xlsm`) stored as
**OLE/CFB** (Compound File Binary) containers with `EncryptionInfo` + `EncryptedPackage` streams
(see MS-OFFCRYPTO / ECMA-376).

These fixtures intentionally live **outside** `fixtures/xlsx/` so they are not picked up by the
ZIP-based XLSX round-trip corpus (e.g. `xlsx-diff::collect_fixture_paths`).

## Passwords

- `agile.xlsx` / `standard.xlsx` / `agile-large.xlsx` / `standard-large.xlsx`: `password`
- `agile-empty-password.xlsx`: empty string (`""`)
- `agile-unicode.xlsx`: `pässwörd` (Unicode, NFC form)
- `agile-basic.xlsm` / `standard-basic.xlsm`: `password`

## Fixtures

- `plaintext.xlsx` – unencrypted ZIP-based workbook (starts with `PK`).
  - Copied from `fixtures/xlsx/basic/basic.xlsx`.
- `agile.xlsx` – Agile encrypted OOXML.
  - `EncryptionInfo` header version **Major 4 / Minor 4**
  - Decrypts to `plaintext.xlsx` with password `password`
- `standard.xlsx` – Standard encrypted OOXML.
  - `EncryptionInfo` header version **Major 3 / Minor 2**
  - Decrypts to `plaintext.xlsx` with password `password`
- `agile-empty-password.xlsx` – Agile encrypted OOXML with an **empty** open password.
  - `EncryptionInfo` header version **Major 4 / Minor 4**
  - Decrypts to `plaintext.xlsx` with password `""`
- `agile-unicode.xlsx` – Agile encrypted OOXML with a Unicode open password.
  - `EncryptionInfo` header version **Major 4 / Minor 4**
  - Decrypts to `plaintext.xlsx` with password `pässwörd` (Unicode, NFC form)
- `plaintext-large.xlsx` – unencrypted ZIP-based workbook, intentionally **> 4096 bytes**.
  - Copied from `fixtures/xlsx/basic/comments.xlsx`.
- `agile-large.xlsx` – Agile encrypted OOXML.
  - `EncryptionInfo` header version **Major 4 / Minor 4**
  - Decrypts to `plaintext-large.xlsx` with password `password`
- `standard-large.xlsx` – Standard encrypted OOXML.
  - `EncryptionInfo` header version **Major 3 / Minor 2**
  - Decrypts to `plaintext-large.xlsx` with password `password`

### Why the `*-large.xlsx` fixtures exist

Agile encryption processes the plaintext package in **4096-byte segments**. Since `plaintext.xlsx` is
< 4096 bytes, decrypting it only exercises the single-segment path. The `*-large.xlsx` fixtures make
sure we cover **multi-segment** decryption.

- `plaintext-basic.xlsm` – unencrypted ZIP-based macro-enabled workbook (starts with `PK`).
  - Copied from `fixtures/xlsx/macros/basic.xlsm`.
- `agile-basic.xlsm` – Agile encrypted macro-enabled workbook.
  - `EncryptionInfo` header version **Major 4 / Minor 4**
  - Decrypts exactly to `plaintext-basic.xlsm` with password `password`
- `standard-basic.xlsm` – Standard encrypted macro-enabled workbook.
  - `EncryptionInfo` header version **Major 3 / Minor 2**
  - Decrypts exactly to `plaintext-basic.xlsm` with password `password`

### Why the `*.xlsm` fixtures exist

The `.xlsm` fixtures exist to ensure the decryption + routing path preserves macros
(`xl/vbaProject.bin`) and correctly classifies the decrypted package as a macro-enabled workbook.

## Usage in tests

These fixtures are referenced explicitly by encryption-focused tests (they are not part of the
ZIP/OPC round-trip corpus under `fixtures/xlsx/`):

- `crates/formula-io/tests/encrypted_ooxml.rs` and `crates/formula-io/tests/encrypted_ooxml_fixtures.rs`:
  format/encryption detection (should surface `PasswordRequired`, including for the empty-password and
  Unicode-password fixtures).
- `crates/formula-io/tests/encrypted_ooxml_fixture_validation.rs`:
  sanity checks that the OLE container and `EncryptionInfo` headers match expectations.
- `crates/formula-io/tests/encrypted_ooxml_decrypt.rs` (behind `formula-io` feature `encrypted-workbooks`):
  end-to-end decryption for `agile.xlsx`, `agile-empty-password.xlsx`, and `agile-unicode.xlsx` against `plaintext.xlsx`,
  plus macro-enabled `.xlsm` fixture coverage (`agile-basic.xlsm` / `standard-basic.xlsm` against
  `plaintext-basic.xlsm`, validating `xl/vbaProject.bin` preservation and `.xlsm` format detection),
  plus on-the-fly Agile encryption/decryption (via `ms_offcrypto_writer`) for
  `open_workbook_with_password` / `open_workbook_model_with_password`.
  Includes coverage that a **missing** password is distinct from an **empty** password (`""`), and
  that Unicode password normalization matters (NFC vs NFD), and wrong-password coverage for
  `standard.xlsx`, `agile-unicode.xlsx`, and the macro-enabled `.xlsm` fixtures.
- `crates/formula-xlsx/tests/encrypted_ooxml_decrypt.rs`:
  end-to-end decryption for `agile-large.xlsx` + `standard-large.xlsx` against `plaintext-large.xlsx`
  (exercises multi-segment decryption).
- `crates/formula-xlsx/tests/encrypted_ooxml_empty_password.rs`:
  decrypts `agile-empty-password.xlsx` and asserts empty password `""` is distinct from a missing password.

## Inspecting encryption headers

You can inspect an encrypted OOXML container (and confirm Agile vs Standard) with:

```bash
bash scripts/cargo_agent.sh run -p formula-io --bin ooxml-encryption-info -- fixtures/encrypted/ooxml/agile.xlsx
bash scripts/cargo_agent.sh run -p formula-io --bin ooxml-encryption-info -- fixtures/encrypted/ooxml/standard.xlsx
bash scripts/cargo_agent.sh run -p formula-io --bin ooxml-encryption-info -- fixtures/encrypted/ooxml/agile-empty-password.xlsx
bash scripts/cargo_agent.sh run -p formula-io --bin ooxml-encryption-info -- fixtures/encrypted/ooxml/agile-unicode.xlsx
bash scripts/cargo_agent.sh run -p formula-io --bin ooxml-encryption-info -- fixtures/encrypted/ooxml/agile-large.xlsx
bash scripts/cargo_agent.sh run -p formula-io --bin ooxml-encryption-info -- fixtures/encrypted/ooxml/standard-large.xlsx
bash scripts/cargo_agent.sh run -p formula-io --bin ooxml-encryption-info -- fixtures/encrypted/ooxml/agile-basic.xlsm
bash scripts/cargo_agent.sh run -p formula-io --bin ooxml-encryption-info -- fixtures/encrypted/ooxml/standard-basic.xlsm
```

See `docs/21-encrypted-workbooks.md` for details on OOXML encryption containers (`EncryptionInfo` /
`EncryptedPackage`).

## Provenance

These fixtures are **synthetic** and **safe-to-ship**. They contain no proprietary or user data.

## Generation notes

The committed fixture binaries were generated using a mix of tooling:

- The baseline fixtures were generated using Python and
  [`msoffcrypto-tool`](https://github.com/nolze/msoffcrypto-tool) **5.4.2**.
- `agile-large.xlsx` was generated using the Rust
  [`ms-offcrypto-writer`](https://crates.io/crates/ms-offcrypto-writer) crate (with a deterministic
  RNG seed) so it includes `dataIntegrity` and is compatible with our strict Agile decryptor.

Implementation detail: `msoffcrypto-tool` includes a minimal OLE writer that does not correctly
handle an `EncryptedPackage` stream **≤ 4096 bytes**. Since `plaintext.xlsx` is tiny, the ciphertext
is padded so that the `EncryptedPackage` stream is **4104 bytes** (8-byte length prefix + 4096 bytes
ciphertext). The embedded unencrypted size prefix still points at the original plaintext length, so
decrypting produces identical bytes.

Alternative regeneration tooling also exists under `tools/encrypted-ooxml-fixtures/` (Apache POI
5.2.5), but was not used to generate the committed fixture bytes above.

## Regenerating with Apache POI (alternative)

To generate encrypted OOXML fixtures **without Excel**, you can use the Apache POI-based generator:

- Source: `tools/encrypted-ooxml-fixtures/GenerateEncryptedXlsx.java`
- Wrapper script (downloads pinned jars + verifies SHA-256): `tools/encrypted-ooxml-fixtures/generate.sh`
- Generator README: `tools/encrypted-ooxml-fixtures/README.md`

Example (from repo root):

```bash
PLAINTEXT=fixtures/encrypted/ooxml/plaintext.xlsx

tools/encrypted-ooxml-fixtures/generate.sh agile password "$PLAINTEXT" /tmp/agile.xlsx
tools/encrypted-ooxml-fixtures/generate.sh standard password "$PLAINTEXT" /tmp/standard.xlsx
tools/encrypted-ooxml-fixtures/generate.sh agile "" "$PLAINTEXT" /tmp/agile-empty-password.xlsx
tools/encrypted-ooxml-fixtures/generate.sh agile "pässwörd" "$PLAINTEXT" /tmp/agile-unicode.xlsx
```

Notes:
- The generated encrypted files are not expected to be byte-for-byte stable across runs (random
  salts/IVs).
- POI's `standard` mode currently emits an `EncryptionInfo` header version of `4.2`
  (still Standard/CryptoAPI encryption).
