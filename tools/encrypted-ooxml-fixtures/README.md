# Encrypted OOXML fixture generator (Apache POI)

This directory contains a small, standalone generator for producing **Office-encrypted OOXML**
workbooks (`.xlsx`) without Excel.

It is an **alternative** regeneration tool for the encrypted OOXML fixtures under
`fixtures/encrypted/ooxml/` (or to create new ones).

Note: the **committed fixture bytes** in `fixtures/encrypted/ooxml/` are generated via Python +
[`msoffcrypto-tool`](https://github.com/nolze/msoffcrypto-tool) (see
`fixtures/encrypted/ooxml/README.md` for the canonical recipe, tool versions, and padding notes).
This Apache POI generator is provided as a convenient cross-platform option, but it is **not** used
to produce the committed fixture binaries.

## What it generates

Excel password-encrypted OOXML workbooks are stored on disk as an **OLE/CFB container** (not a ZIP),
containing:

- `EncryptionInfo`
- `EncryptedPackage` (encrypted bytes of the underlying `.xlsx` ZIP payload)

The generator wraps a plaintext OOXML ZIP (`.xlsx`) into this container using Apache POI.

## Usage

From the repo root:

```bash
tools/encrypted-ooxml-fixtures/generate.sh agile password fixtures/encrypted/ooxml/plaintext.xlsx /tmp/agile.xlsx
tools/encrypted-ooxml-fixtures/generate.sh standard password fixtures/encrypted/ooxml/plaintext.xlsx /tmp/standard.xlsx

# Empty password (note the explicit empty string argument):
tools/encrypted-ooxml-fixtures/generate.sh agile "" fixtures/encrypted/ooxml/plaintext.xlsx /tmp/agile-empty-password.xlsx
```

## Reproducibility / supply-chain safety

`generate.sh`:

- downloads a **pinned** set of jars from Maven Central into `tools/encrypted-ooxml-fixtures/.cache/`
- verifies each jar by **SHA-256 checksum**
- caches downloaded jars + compiled classes in `tools/encrypted-ooxml-fixtures/.cache/` (gitignored)

Apache POI version: **5.2.5** (see `generate.sh` for the full pinned jar list + checksums).

## Notes

- The encrypted output is **not expected to be byte-for-byte stable** across runs (encryption uses
  random salts/IVs).
- POI's `standard` mode currently emits an `EncryptionInfo` header version of `4.2`
  (still Standard/CryptoAPI encryption).
