use std::path::Path;

use anyhow::Result;

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
        let roundtripped = tmpdir.path().join("roundtripped.xlsx");

        xlsx_diff::roundtrip_zip_copy(&fixture, &roundtripped)?;

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
