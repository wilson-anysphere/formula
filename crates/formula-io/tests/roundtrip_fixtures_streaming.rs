use std::path::Path;

use formula_io::{open_workbook, save_workbook, Error};
use xlsx_diff::Severity;

#[test]
fn roundtrip_xlsx_fixtures_open_save_no_critical_diffs() {
    let fixtures_root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx");
    let fixtures =
        xlsx_diff::collect_fixture_paths(&fixtures_root).expect("collect xlsx fixtures");
    assert!(
        !fixtures.is_empty(),
        "no fixtures found under {}",
        fixtures_root.display()
    );

    let tmpdir = tempfile::tempdir().expect("temp dir");

    let mut processed = 0usize;
    for (idx, fixture) in fixtures.iter().enumerate() {
        let wb = match open_workbook(fixture) {
            Ok(wb) => wb,
            Err(Error::PasswordRequired { .. } | Error::EncryptedWorkbook { .. }) => {
                // Roundtrip tests are meant for plaintext fixtures. Encrypted OOXML/BIFF files must
                // be opened via the password APIs and are covered by dedicated encrypted fixture
                // tests.
                eprintln!("skipping encrypted fixture {}", fixture.display());
                continue;
            }
            Err(err) => {
                panic!("open_workbook failed for {}: {err}", fixture.display());
            }
        };
        processed += 1;

        let extension = fixture
            .extension()
            .and_then(|ext| ext.to_str())
            .unwrap_or("xlsx");
        let out_path = tmpdir.path().join(format!("roundtrip_{idx}.{extension}"));

        save_workbook(&wb, &out_path).unwrap_or_else(|err| {
            panic!(
                "save_workbook failed for {} -> {}: {err}",
                fixture.display(),
                out_path.display()
            )
        });

        let report = xlsx_diff::diff_workbooks(fixture, &out_path).unwrap_or_else(|err| {
            panic!(
                "diff_workbooks failed for {} vs {}: {err}",
                fixture.display(),
                out_path.display()
            )
        });

        if report.has_at_least(Severity::Critical) {
            eprintln!("Critical diffs detected for fixture {}", fixture.display());
            for diff in report
                .differences
                .iter()
                .filter(|d| d.severity == Severity::Critical)
            {
                eprintln!("{diff}");
            }
            panic!("fixture {} did not round-trip cleanly", fixture.display());
        }
    }

    assert!(processed > 0, "no plaintext fixtures were round-tripped");
}
