use std::fs::File;
use std::io::Read;
use std::path::Path;

/// Guard against accidentally adding OLE/CFB "encrypted OOXML" workbooks into the ZIP fixture
/// corpus.
///
/// The `fixtures/xlsx/**` round-trip harness assumes each `*.xlsx` / `*.xlsm` / `*.xlsb` file is a
/// ZIP/Open Packaging Convention container. Password-protected OOXML files are *not* ZIPs: Excel
/// stores them as OLE/CFB compound documents (`D0 CF 11 E0 A1 B1 1A E1` header).
#[test]
fn xlsx_fixtures_are_zip_archives() -> Result<(), Box<dyn std::error::Error>> {
    let fixtures_root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx");
    let fixtures = xlsx_diff::collect_fixture_paths(&fixtures_root)?;
    assert!(
        !fixtures.is_empty(),
        "no fixtures found under {}",
        fixtures_root.display()
    );

    let ole_magic: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

    for fixture in fixtures {
        let mut file = File::open(&fixture)?;

        // Read up to 8 bytes so we can detect both ZIP (`PK..`) and OLE/CFB (`D0 CF 11 E0 ...`)
        // headers.
        let mut prefix = [0u8; 8];
        let mut read = 0usize;
        while read < prefix.len() {
            let n = file.read(&mut prefix[read..])?;
            if n == 0 {
                break;
            }
            read += n;
        }

        if read < 2 {
            panic!(
                "Fixture {} is too small to be a ZIP/OPC workbook (read {read} bytes). \
                 Fixtures under fixtures/xlsx/ must be ZIP-based OOXML workbooks; \
                 encrypted/password-protected fixtures belong under fixtures/encrypted/...",
                fixture.display()
            );
        }

        // `PK\x03\x04` is the local file header signature, but `PK` is enough to distinguish ZIP
        // from OLE/CFB for our purposes.
        if !prefix.starts_with(b"PK") {
            let prefix_hex = prefix[..read]
                .iter()
                .map(|b| format!("{b:02X}"))
                .collect::<Vec<_>>()
                .join(" ");

            if read == 8 && prefix == ole_magic {
                panic!(
                    "Fixture {} is an OLE/CFB compound file (header {prefix_hex}), not a ZIP/OPC workbook. \
                     This is typical of encrypted/password-protected OOXML files and will break round-trip tests. \
                     Move it under fixtures/encrypted/....",
                    fixture.display()
                );
            }

            panic!(
                "Fixture {} does not look like a ZIP/OPC workbook: expected it to start with ZIP magic 'PK', got {prefix_hex}. \
                 If this workbook is encrypted/password-protected, it is likely an OLE/CFB container and must live under fixtures/encrypted/....",
                fixture.display()
            );
        }
    }

    Ok(())
}

