# Encrypted OOXML (`.xlsx`) fixtures

Files in this directory are Excel workbooks that require a password to open.

Even though they use the `.xlsx` extension, encrypted OOXML workbooks are **not ZIP/OPC
packages**. Excel stores them as OLE/CFB (Compound File Binary) containers with
`EncryptionInfo` + `EncryptedPackage` streams (see MS-OFFCRYPTO).

These fixtures live outside `fixtures/xlsx/` so they are not picked up by
`xlsx-diff::collect_fixture_paths` (which drives the ZIP-based round-trip test corpus).
