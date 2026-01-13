# Encrypted OOXML (`.xlsx`) fixtures

This directory is the home for Excel workbooks that require a password to open.

Even though they use the `.xlsx` extension, encrypted OOXML workbooks are **not ZIP/OPC
packages**. Excel stores them as OLE/CFB (Compound File Binary) containers with
`EncryptionInfo` + `EncryptedPackage` streams (see MS-OFFCRYPTO).

These fixtures live outside `fixtures/xlsx/` so they are not picked up by
`xlsx-diff::collect_fixture_paths` (which drives the ZIP-based round-trip test corpus).

If you add a real encrypted workbook fixture, document the expected password in the adjacent README
(or encode it in the filename) so tests can open it deterministically.
