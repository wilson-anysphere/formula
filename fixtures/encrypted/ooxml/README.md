# Encrypted OOXML XLSX fixtures

This directory contains **password-protected OOXML workbooks** (`.xlsx`) stored as **OLE/CFB**
(Compound File Binary) containers with `EncryptionInfo` + `EncryptedPackage` streams
(see MS-OFFCRYPTO / ECMA-376).

These fixtures intentionally live **outside** `fixtures/xlsx/` so they are not picked up by the
ZIP-based XLSX round-trip corpus (e.g. `xlsx-diff::collect_fixture_paths`).

## Passwords

- `agile.xlsx` / `standard.xlsx`: `password`
- `agile-empty-password.xlsx`: empty string (`""`)

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

## Inspecting encryption headers

You can inspect an encrypted OOXML container (and confirm Agile vs Standard) with:

```bash
bash scripts/cargo_agent.sh run -p formula-io --bin ooxml-encryption-info -- fixtures/encrypted/ooxml/agile.xlsx
bash scripts/cargo_agent.sh run -p formula-io --bin ooxml-encryption-info -- fixtures/encrypted/ooxml/standard.xlsx
bash scripts/cargo_agent.sh run -p formula-io --bin ooxml-encryption-info -- fixtures/encrypted/ooxml/agile-empty-password.xlsx
```

## Provenance

These fixtures are **synthetic** and **safe-to-ship**. They contain no proprietary or user data.

## Generation notes

The committed fixture binaries were generated using Python and
[`msoffcrypto-tool`](https://github.com/nolze/msoffcrypto-tool) **5.4.2**.

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

Example (from repo root):

```bash
PLAINTEXT=fixtures/encrypted/ooxml/plaintext.xlsx

tools/encrypted-ooxml-fixtures/generate.sh agile password "$PLAINTEXT" /tmp/agile.xlsx
tools/encrypted-ooxml-fixtures/generate.sh standard password "$PLAINTEXT" /tmp/standard.xlsx
tools/encrypted-ooxml-fixtures/generate.sh agile "" "$PLAINTEXT" /tmp/agile-empty-password.xlsx
```

Notes:
- The generated encrypted files are not expected to be byte-for-byte stable across runs (random
  salts/IVs).
- POI's `standard` mode currently emits an `EncryptionInfo` header version of `4.2`
  (still Standard/CryptoAPI encryption).
