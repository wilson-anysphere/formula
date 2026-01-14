# Encrypted OOXML fixture generator (Apache POI)

This directory contains a small, standalone generator for producing **Office-encrypted OOXML**
spreadsheets (`.xlsx`/`.xlsm`/`.xlsb`) without Excel.

It is an **alternative** regeneration tool for the encrypted OOXML fixtures under
`fixtures/encrypted/ooxml/` (or to create new ones).

Note: most committed fixture bytes in `fixtures/encrypted/ooxml/` are generated via Python +
[`msoffcrypto-tool`](https://github.com/nolze/msoffcrypto-tool), but some fixtures are generated via
Apache POI (for example `standard-4.2.xlsx`, `standard-unicode.xlsx`). See
`fixtures/encrypted/ooxml/README.md` for the canonical per-fixture provenance + regeneration notes.

## What it generates

Excel password-encrypted OOXML workbooks are stored on disk as an **OLE/CFB container** (not a ZIP),
containing:

- `EncryptionInfo`
- `EncryptedPackage` (encrypted bytes of the underlying `.xlsx`/`.xlsm`/`.xlsb` ZIP payload)

The generator wraps a plaintext OOXML ZIP (`.xlsx`/`.xlsm`/`.xlsb`) into this container using Apache POI.

## Usage

From the repo root:

```bash
tools/encrypted-ooxml-fixtures/generate.sh agile password fixtures/encrypted/ooxml/plaintext.xlsx /tmp/agile.xlsx
tools/encrypted-ooxml-fixtures/generate.sh standard password fixtures/encrypted/ooxml/plaintext.xlsx /tmp/standard.xlsx

# Empty password (note the explicit empty string argument):
tools/encrypted-ooxml-fixtures/generate.sh agile "" fixtures/encrypted/ooxml/plaintext.xlsx /tmp/agile-empty-password.xlsx

# Unicode password (NFC normalization form):
tools/encrypted-ooxml-fixtures/generate.sh agile "pässwörd" fixtures/encrypted/ooxml/plaintext.xlsx /tmp/agile-unicode.xlsx

# Unicode password (NFC, includes non-BMP emoji):
tools/encrypted-ooxml-fixtures/generate.sh standard "pässwörd🔒" fixtures/encrypted/ooxml/plaintext.xlsx /tmp/standard-unicode.xlsx

# Macro-enabled `.xlsm` fixtures:
tools/encrypted-ooxml-fixtures/generate.sh agile password fixtures/encrypted/ooxml/plaintext-basic.xlsm /tmp/agile-basic.xlsm
tools/encrypted-ooxml-fixtures/generate.sh standard password fixtures/encrypted/ooxml/plaintext-basic.xlsm /tmp/standard-basic.xlsm

# `.xlsb` fixtures (OOXML-in-OLE encrypted wrappers):
tools/encrypted-ooxml-fixtures/generate.sh agile password crates/formula-xlsb/tests/fixtures/simple.xlsb /tmp/agile.xlsb
```

## Reproducibility / supply-chain safety

`generate.sh`:

- downloads a **pinned** set of jars from Maven Central into `tools/encrypted-ooxml-fixtures/.cache/`
- verifies each jar by **SHA-256 checksum**
- caches downloaded jars + compiled classes in `tools/encrypted-ooxml-fixtures/.cache/` (gitignored)

Apache POI version: **5.2.5** (see `generate.sh` for the full pinned jar list + checksums).

You can override the cache location by setting:

```bash
export ENCRYPTED_OOXML_FIXTURES_CACHE_DIR="$HOME/.cache/formula/encrypted-ooxml-fixtures"
```

## Notes

- The encrypted output is **not expected to be byte-for-byte stable** across runs (encryption uses
  random salts/IVs).
- POI's `standard` mode currently emits an `EncryptionInfo` header version of `4.2`
  (still Standard/CryptoAPI encryption).
