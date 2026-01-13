# Encrypted XLSX fixtures

This directory contains **password-to-open** Excel workbooks that use an **OLE/CFB wrapper**
(`EncryptionInfo` + `EncryptedPackage`, per **MS-OFFCRYPTO**). These files are intentionally **not**
part of the ZIP/OPC round-trip fixture corpus even though they use the `.xlsx` extension.

Why this lives under `fixtures/xlsx/`:

- The ZIP/OPC round-trip harness (`crates/xlsx-diff`) enumerates `fixtures/xlsx/**` and then opens
  each file as a ZIP archive.
- Encrypted OOXML workbooks are **not ZIP files on disk**, so `xlsx-diff::collect_fixture_paths`
  intentionally skips the `fixtures/xlsx/encrypted/` subtree.

Canonical encrypted fixture inventory (including more formats/passwords) lives under:

- `fixtures/encrypted/` (real-world encrypted `.xlsx`/`.xlsb`/`.xls`)
- `fixtures/encrypted/ooxml/` (vendored encrypted `.xlsx`/`.xlsm` corpus + passwords)

Password for all fixtures in this directory: `Password1234_`

## Fixtures

Generated from `fixtures/xlsx/basic/basic.xlsx` (1 sheet with `A1=1`, `B1="Hello"`):

- `agile-password.xlsx` — ECMA-376 **Agile** encryption
- `standard-password.xlsx` — ECMA-376 **Standard** encryption

Apache POI–generated Standard fixture with a pinned plaintext payload:

- `standard_password.xlsx` – **Standard** (CryptoAPI) encrypted OOXML workbook
- `standard_password_plain.xlsx` – the decrypted plaintext `.xlsx` ZIP (what you get after
  decrypting `standard_password.xlsx`)

The plaintext workbook is intentionally minimal:

- `Sheet1!A1` = `"hello"`

## How `standard_password.xlsx` was produced (one-time)

The encrypted fixture was generated once (and the resulting bytes were committed) using **Apache
POI** with `EncryptionMode.standard`.

Pseudo-code:

```java
POIFSFileSystem fs = new POIFSFileSystem();
EncryptionInfo info = new EncryptionInfo(EncryptionMode.standard);
Encryptor enc = info.getEncryptor();
enc.confirmPassword("Password1234_");

// Encrypt the plaintext ZIP bytes directly.
try (InputStream is = new FileInputStream("standard_password_plain.xlsx");
     OutputStream os = enc.getDataStream(fs)) {
  is.transferTo(os);
}

try (FileOutputStream out = new FileOutputStream("standard_password.xlsx")) {
  fs.writeFilesystem(out);
}
```

Tests must **not** depend on Java/POI at runtime; they use the committed bytes directly.
