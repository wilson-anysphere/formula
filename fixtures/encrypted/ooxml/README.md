# Encrypted OOXML fixtures (`.xlsx` / `.xlsm`)

This directory contains **password-protected OOXML workbooks** (`.xlsx`, `.xlsm`) stored as
**OLE/CFB** (Compound File Binary) containers with `EncryptionInfo` + `EncryptedPackage` streams
(see MS-OFFCRYPTO / ECMA-376).

These fixtures intentionally live **outside** `fixtures/xlsx/` so they are not picked up by the
ZIP-based XLSX round-trip corpus (e.g. `xlsx-diff::collect_fixture_paths`).

Note: this directory currently vendors encrypted `.xlsx`/`.xlsm` samples only. For an encrypted
`.xlsb` fixture, see `fixtures/encrypted/encrypted.xlsb` (real-world Apache POI corpus; password
`tika`).

## Passwords

- `agile.xlsx` / `standard.xlsx` / `standard-4.2.xlsx` / `standard-rc4.xlsx` / `agile-large.xlsx` / `standard-large.xlsx`: `password`
- `agile.xlsm` / `standard.xlsm` / `agile-basic.xlsm` / `standard-basic.xlsm` / `basic-password.xlsm`: `password`
- `agile-empty-password.xlsx`: empty string (`""`)
- `agile-unicode.xlsx`: `pässwörd` (Unicode, NFC form)
- `agile-unicode-excel.xlsx`: `pässwörd🔒` (Unicode, NFC form, includes non-BMP emoji)
- `standard-unicode.xlsx`: `pässwörd🔒` (Unicode, NFC form, includes non-BMP emoji)

End-to-end fixtures that pair an encrypted file with its decrypted plaintext (and thus need a
stable known password) can live under `fixtures/xlsx/encrypted/`, which is explicitly skipped by
`xlsx-diff::collect_fixture_paths`.

## Fixtures

- `plaintext.xlsx` – unencrypted ZIP-based workbook (starts with `PK`).
  - Copied from `fixtures/xlsx/basic/basic.xlsx`.
- `plaintext-excel.xlsx` – unencrypted ZIP-based workbook produced by Microsoft Excel (starts with `PK`).
  - Copied from `crates/formula-offcrypto/tests/fixtures/outputs/example.xlsx`.
- `plaintext.xlsm` – unencrypted ZIP-based macro-enabled workbook (starts with `PK`).
  - Copied from `fixtures/xlsx/macros/basic.xlsm` (includes `xl/vbaProject.bin`). (Identical to
    `plaintext-basic.xlsm`; kept for naming symmetry with `plaintext.xlsx`.)
- `agile.xlsx` – Agile encrypted OOXML.
  - `EncryptionInfo` header version **Major 4 / Minor 4**
  - Decrypts to `plaintext.xlsx` with password `password`
- `standard.xlsx` – Standard encrypted OOXML.
   - `EncryptionInfo` header version **Major 3 / Minor 2**
   - `EncryptionInfo` flags: `fCryptoAPI` + `fAES` (`0x00000024`)
   - ECMA-376/MS-OFFCRYPTO Standard: 50,000 password-hash iterations + AES-ECB
   - Decrypts to `plaintext.xlsx` with password `password`
- `standard-4.2.xlsx` – Standard encrypted OOXML (Apache POI output).
   - `EncryptionInfo` header version **Major 4 / Minor 2**
   - Decrypts to `plaintext.xlsx` with password `password`
- `standard-unicode.xlsx` – Standard encrypted OOXML with a Unicode open password (Apache POI output).
   - `EncryptionInfo` header version **Major 4 / Minor 2**
   - Decrypts to `plaintext.xlsx` with password `pässwörd🔒` (Unicode, NFC form, includes non-BMP emoji)
- `standard-rc4.xlsx` – Standard encrypted OOXML (RC4 CryptoAPI).
  - `EncryptionInfo` header version **Major 3 / Minor 2**
  - `EncryptionHeader.algId` = `CALG_RC4` (`0x00006801`)
  - Decrypts to `plaintext.xlsx` with password `password`
- `agile.xlsm` – Agile encrypted OOXML (macro-enabled workbook package).
  - `EncryptionInfo` header version **Major 4 / Minor 4**
  - Decrypts to `plaintext.xlsm` with password `password`
- `standard.xlsm` – Standard encrypted OOXML (macro-enabled workbook package).
  - `EncryptionInfo` header version **Major 3 / Minor 2**
  - Decrypts to `plaintext.xlsm` with password `password`
- `agile-empty-password.xlsx` – Agile encrypted OOXML with an **empty** open password.
    - `EncryptionInfo` header version **Major 4 / Minor 4**
    - Decrypts to `plaintext.xlsx` with password `""`
- `agile-unicode.xlsx` – Agile encrypted OOXML with a Unicode open password.
  - `EncryptionInfo` header version **Major 4 / Minor 4**
  - Decrypts to `plaintext.xlsx` with password `pässwörd` (Unicode, NFC form)
- `agile-unicode-excel.xlsx` – Agile encrypted OOXML with a Unicode open password.
  - `EncryptionInfo` header version **Major 4 / Minor 4**
  - Decrypts to `plaintext-excel.xlsx` with password `pässwörd🔒` (Unicode, NFC form, includes non-BMP emoji)
- `plaintext-large.xlsx` – unencrypted ZIP-based workbook, intentionally **> 4096 bytes**.
   - Copied from `fixtures/xlsx/basic/comments.xlsx`.
- `agile-large.xlsx` – Agile encrypted OOXML.
    - `EncryptionInfo` header version **Major 4 / Minor 4**
   - Decrypts to `plaintext-large.xlsx` with password `password`
- `standard-large.xlsx` – Standard encrypted OOXML.
   - `EncryptionInfo` header version **Major 3 / Minor 2**
   - `EncryptionInfo` flags: `fCryptoAPI` + `fAES` (`0x00000024`)
   - Decrypts to `plaintext-large.xlsx` with password `password`

### Why the `*-large.xlsx` fixtures exist

Agile encryption processes the plaintext package in **4096-byte segments**. Since `plaintext.xlsx` is
< 4096 bytes, decrypting it only exercises the single-segment path. The `*-large.xlsx` fixtures make
sure we cover **multi-segment** decryption.

Note: ECMA-376/MS-OFFCRYPTO Standard encryption uses AES-ECB (no IV), so multi-segment decryption is
not a meaningful distinction for the Standard algorithm itself. `standard-large.xlsx` still exists
as a regression fixture to ensure we cover Standard decryption against a **larger** package (and
therefore more realistic ciphertext sizes + truncation behavior).

- `plaintext-basic.xlsm` – unencrypted ZIP-based macro-enabled workbook (starts with `PK`).
  - Copied from `fixtures/xlsx/macros/basic.xlsm`.
- `agile-basic.xlsm` – Agile encrypted macro-enabled workbook.
  - `EncryptionInfo` header version **Major 4 / Minor 4**
  - Decrypts exactly to `plaintext-basic.xlsm` with password `password`
- `standard-basic.xlsm` – Standard encrypted macro-enabled workbook.
  - `EncryptionInfo` header version **Major 3 / Minor 2**
  - `EncryptionInfo` flags: `fCryptoAPI` + `fAES` (`0x00000024`)
  - Decrypts exactly to `plaintext-basic.xlsm` with password `password`
- `basic-password.xlsm` – Agile encrypted macro-enabled workbook with a minimal VBA project.
  - `EncryptionInfo` header version **Major 4 / Minor 4**
  - Decrypts with password `password`
  - Source workbook (unencrypted): `fixtures/xlsx/macros/basic.xlsm`

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
- `crates/formula-io/tests/agile_encryption_info.rs`:
  pins the committed Agile `EncryptionInfo` XML parameters (`spinCount` / algorithms / key size /
  salt size) and keeps the table in this README in sync (to prevent silent fixture regeneration
  drift and to keep decryption CI performance predictable).
- `crates/formula-io/tests/encrypted_ooxml_decrypt.rs` (behind `formula-io` feature `encrypted-workbooks`):
  end-to-end decryption for Standard + Agile fixtures (including `standard.xlsx` / `standard-4.2.xlsx` / `standard-rc4.xlsx` / `standard-unicode.xlsx` and `agile.xlsx` / `agile-empty-password.xlsx` / `agile-unicode.xlsx`) against `plaintext.xlsx`, plus `agile-unicode-excel.xlsx` against `plaintext-excel.xlsx`,
  plus large-package coverage (`agile-large.xlsx` / `standard-large.xlsx` against `plaintext-large.xlsx`, exercising multi-segment decryption for Agile),
  plus macro-enabled `.xlsm` fixture coverage (`agile-basic.xlsm` / `standard-basic.xlsm` against `plaintext-basic.xlsm`, and `agile.xlsm` / `standard.xlsm` against `plaintext.xlsm`), validating `xl/vbaProject.bin` preservation and `.xlsm` format detection,
  plus on-the-fly Agile encryption/decryption (via `ms_offcrypto_writer`) for
  `open_workbook_with_password` / `open_workbook_model_with_password`.
  Includes coverage that a **missing** password is distinct from an **empty** password (`""`), and
  that Unicode password normalization matters (NFC vs NFD), and wrong-password coverage for
  these fixtures.
- `crates/formula-io/tests/open_encrypted_xlsm_with_password.rs` (behind `formula-io` feature `encrypted-workbooks`):
  opens `basic-password.xlsm` and asserts the VBA parts survive decryption.
- `crates/formula-xlsx/tests/encrypted_ooxml_decrypt.rs`:
  end-to-end decryption for `agile-large.xlsx` + `standard-large.xlsx` against `plaintext-large.xlsx`
  (exercises multi-segment decryption for Agile, and larger-package coverage for Standard), plus
  macro-enabled `.xlsm` fixtures (including `agile.xlsm` / `standard.xlsm`).
- `crates/formula-xlsx/tests/encrypted_ooxml_empty_password.rs`:
  decrypts `agile-empty-password.xlsx` and asserts empty password `""` is distinct from a missing password.
- `crates/xlsx-diff/tests/encrypted_ooxml_diff.rs`:
  asserts the OPC-level diff tool can decrypt and compare encrypted `.xlsx`/`.xlsm` fixtures against the
  corresponding plaintext packages (and also exercises an on-the-fly encrypted `.xlsb` wrapper).
- `crates/formula-offcrypto/tests/standard_encryptedpackage_mode.rs`:
  canary asserting the committed Standard AES fixtures use the **ECB** `EncryptedPackage` mode (and
  will loudly fail if fixture regeneration drifts to segmented CBC).

## Agile `EncryptionInfo` parameters (pinned)

Agile encryption stores configuration as an XML `<encryption>` document inside the `EncryptionInfo`
stream. Fixture regeneration tooling (e.g. `msoffcrypto-tool` / Apache POI version bumps) can
silently change defaults like `spinCount`, hash algorithm, or cipher key size; those changes can
also impact CI runtime (large `spinCount` values make password-based key derivation expensive).

The table below is the **canonical** source of truth for committed Agile fixtures, and is asserted
in `crates/formula-io/tests/agile_encryption_info.rs` to prevent silent drift (including an upper
bound on `spinCount` to keep CI runtime predictable).

| fixture | spinCount | cipherAlgorithm | cipherChaining | keyBits | hashAlgorithm | saltSize |
| --- | ---: | --- | --- | ---: | --- | ---: |
| agile.xlsx | 100000 | AES | ChainingModeCBC | 256 | SHA512 | 16 |
| agile-large.xlsx | 100000 | AES | ChainingModeCBC | 256 | SHA512 | 16 |
| agile-unicode.xlsx | 100000 | AES | ChainingModeCBC | 256 | SHA512 | 16 |
| agile-unicode-excel.xlsx | 100000 | AES | ChainingModeCBC | 256 | SHA512 | 16 |
| agile-basic.xlsm | 100000 | AES | ChainingModeCBC | 256 | SHA512 | 16 |
| basic-password.xlsm | 100000 | AES | ChainingModeCBC | 256 | SHA512 | 16 |
| agile-empty-password.xlsx | 1000 | AES | ChainingModeCBC | 128 | SHA256 | 16 |

## Inspecting encryption headers

You can inspect an encrypted OOXML container (and confirm Agile vs Standard) with:

```bash
bash scripts/cargo_agent.sh run -p formula-io --bin ooxml-encryption-info -- fixtures/encrypted/ooxml/agile.xlsx
bash scripts/cargo_agent.sh run -p formula-io --bin ooxml-encryption-info -- fixtures/encrypted/ooxml/standard.xlsx
bash scripts/cargo_agent.sh run -p formula-io --bin ooxml-encryption-info -- fixtures/encrypted/ooxml/standard-rc4.xlsx
bash scripts/cargo_agent.sh run -p formula-io --bin ooxml-encryption-info -- fixtures/encrypted/ooxml/agile.xlsm
bash scripts/cargo_agent.sh run -p formula-io --bin ooxml-encryption-info -- fixtures/encrypted/ooxml/standard.xlsm
bash scripts/cargo_agent.sh run -p formula-io --bin ooxml-encryption-info -- fixtures/encrypted/ooxml/agile-empty-password.xlsx
bash scripts/cargo_agent.sh run -p formula-io --bin ooxml-encryption-info -- fixtures/encrypted/ooxml/agile-unicode.xlsx
bash scripts/cargo_agent.sh run -p formula-io --bin ooxml-encryption-info -- fixtures/encrypted/ooxml/agile-unicode-excel.xlsx
bash scripts/cargo_agent.sh run -p formula-io --bin ooxml-encryption-info -- fixtures/encrypted/ooxml/standard-unicode.xlsx
bash scripts/cargo_agent.sh run -p formula-io --bin ooxml-encryption-info -- fixtures/encrypted/ooxml/agile-large.xlsx
bash scripts/cargo_agent.sh run -p formula-io --bin ooxml-encryption-info -- fixtures/encrypted/ooxml/standard-large.xlsx
bash scripts/cargo_agent.sh run -p formula-io --bin ooxml-encryption-info -- fixtures/encrypted/ooxml/agile-basic.xlsm
bash scripts/cargo_agent.sh run -p formula-io --bin ooxml-encryption-info -- fixtures/encrypted/ooxml/standard-basic.xlsm
bash scripts/cargo_agent.sh run -p formula-io --bin ooxml-encryption-info -- fixtures/encrypted/ooxml/basic-password.xlsm
```

See `docs/21-encrypted-workbooks.md` for details on OOXML encryption containers (`EncryptionInfo` /
`EncryptedPackage`).

## Provenance

These fixtures are **synthetic** and **safe-to-ship**. They contain no proprietary or user data.

## Generation notes

The committed fixture binaries were generated using a mix of tooling:

- `agile.xlsx` and `standard.xlsx` were generated using Python and
  [`msoffcrypto-tool`](https://github.com/nolze/msoffcrypto-tool) **5.4.2**.
- `agile.xlsm` and `standard.xlsm` were generated using Python and
  [`msoffcrypto-tool`](https://github.com/nolze/msoffcrypto-tool) **5.4.2** (from `plaintext.xlsm`).
- `standard-large.xlsx` is a synthetic fixture aligned with `crates/formula-xlsx::offcrypto`’s
  Standard decrypt behavior (see `docs/office-encryption.md`).
- `basic-password.xlsm` was generated in Excel (starting from `fixtures/xlsx/macros/basic.xlsm`)
  and saved with open password `password`.
- `agile-large.xlsx` was generated using the Rust
  [`ms-offcrypto-writer`](https://crates.io/crates/ms-offcrypto-writer) crate (with a deterministic
  RNG seed) so it includes `dataIntegrity` and is compatible with our strict Agile decryptor.
- `standard-4.2.xlsx` was generated using **Apache POI 5.2.5** via
  `tools/encrypted-ooxml-fixtures/generate.sh standard`.
- `standard-unicode.xlsx` was generated using **Apache POI 5.2.5** via
  `tools/encrypted-ooxml-fixtures/generate.sh standard`.

Implementation detail: `msoffcrypto-tool` includes a minimal OLE writer that does not correctly
handle an `EncryptedPackage` stream **≤ 4096 bytes**. Since `plaintext.xlsx` is tiny, the ciphertext
is padded so that the `EncryptedPackage` stream is **4104 bytes** (8-byte length prefix + 4096 bytes
ciphertext). The embedded unencrypted size prefix still points at the original plaintext length, so
decrypting produces identical bytes.

Alternative regeneration tooling also exists under `tools/encrypted-ooxml-fixtures/` (Apache POI
5.2.5) and is used to generate `standard-4.2.xlsx`.

`standard-rc4.xlsx` was generated using the in-repo Rust example
`crates/formula-io/examples/generate_standard_rc4_ooxml_fixture.rs` (deterministic output):

```bash
cargo run -p formula-io --example generate_standard_rc4_ooxml_fixture -- \
  fixtures/encrypted/ooxml/plaintext.xlsx \
  fixtures/encrypted/ooxml/standard-rc4.xlsx
```

## Regenerating Standard (CryptoAPI) fixtures (deterministic)

From the repo root:

```bash
bash scripts/cargo_agent.sh run -p formula-xlsx --example regen_standard_fixtures
```

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

## `EncryptionInfo` parameters pinned by tests

These values are asserted by `crates/formula-io/tests/encrypted_ooxml_fixtures.rs` to prevent silent
fixture drift (cipher/keysize/provider/flags).

### Standard / CryptoAPI AES fixtures (`standard.xlsx`, `standard-large.xlsx`, `standard-basic.xlsm`, `standard-4.2.xlsx`, `standard-unicode.xlsx`)

`EncryptionVersionInfo` (stream prefix):

- `major`:
  - `3`: `standard.xlsx`, `standard-large.xlsx`, `standard-basic.xlsm`
  - `4`: `standard-4.2.xlsx`, `standard-unicode.xlsx` (Apache POI output)
- `minor`: `2`
- `flags`: `0x00000024` (`fCryptoAPI` + `fAES`)

`EncryptionHeader` (CryptoAPI):

- `AlgID`: `0x0000660E` (CALG_AES_128)
- `AlgIDHash`: `0x00008004` (CALG_SHA1)
- `KeySize`: `128` (bits)
- `ProviderType`: `24` (PROV_RSA_AES)
- `CSPName`:
  - `"Microsoft Enhanced RSA and AES Cryptographic Provider"`

`EncryptionVerifier`:

- `saltSize`: `16`
