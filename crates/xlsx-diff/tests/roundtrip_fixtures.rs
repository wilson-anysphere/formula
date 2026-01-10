use std::path::Path;

use anyhow::{Context, Result};

#[test]
fn roundtrip_fixtures_no_critical_diffs() -> Result<()> {
    let fixtures_root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx");
    let fixtures = xlsx_diff::collect_fixture_paths(&fixtures_root)?;
    assert!(
        !fixtures.is_empty(),
        "no fixtures found under {}",
        fixtures_root.display()
    );

    for fixture in fixtures {
        let tmpdir = tempfile::tempdir()?;
        let extension = fixture.extension().and_then(|ext| ext.to_str()).unwrap_or("xlsx");
        let roundtripped = tmpdir
            .path()
            .join(format!("roundtripped.{extension}"));

        // Round-trip via the current XLSX package implementation (OPC-level).
        let original_bytes =
            std::fs::read(&fixture).with_context(|| format!("read fixture {}", fixture.display()))?;
        let pkg = formula_xlsx::XlsxPackage::from_bytes(&original_bytes)
            .with_context(|| format!("parse fixture {}", fixture.display()))?;
        let written_bytes = pkg
            .write_to_bytes()
            .with_context(|| format!("write roundtripped fixture {}", fixture.display()))?;
        std::fs::write(&roundtripped, written_bytes)
            .with_context(|| format!("write temp file {}", roundtripped.display()))?;

        let report = xlsx_diff::diff_workbooks(&fixture, &roundtripped)?;
        if report.has_at_least(xlsx_diff::Severity::Critical) {
            eprintln!("Critical diffs detected for fixture {}", fixture.display());
            for diff in report
                .differences
                .iter()
                .filter(|d| d.severity == xlsx_diff::Severity::Critical)
            {
                eprintln!("{diff}");
            }
            panic!("fixture {} did not round-trip cleanly", fixture.display());
        }
    }

    Ok(())
}
