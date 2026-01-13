use std::path::Path;

use formula_model::CellValue;
use formula_xlsx::{load_from_bytes, load_from_path};

#[test]
fn reads_rich_text_shared_string_cells_as_rich_text() -> Result<(), Box<dyn std::error::Error>> {
    let fixture_path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/xlsx/styles/rich-text-shared-strings.xlsx");

    let doc = load_from_path(&fixture_path)?;
    let sheet_id = doc.workbook.sheets[0].id;
    let sheet = doc.workbook.sheet(sheet_id).ok_or("missing sheet")?;

    let value = sheet.value_a1("A1")?;
    match value {
        CellValue::RichText(rich) => {
            assert_eq!(rich.text, "Hello Bold Italic");
            assert!(
                !rich.runs.is_empty(),
                "expected rich text runs to be preserved for shared string"
            );
        }
        other => panic!("expected A1 to be CellValue::RichText, got {other:?}"),
    }

    Ok(())
}

#[test]
fn noop_roundtrip_does_not_rewrite_rich_text_shared_string_cells(
) -> Result<(), Box<dyn std::error::Error>> {
    let fixture_path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/xlsx/styles/rich-text-shared-strings.xlsx");
    let bytes = std::fs::read(&fixture_path)?;

    let doc = load_from_bytes(&bytes)?;
    let saved = doc.save_to_vec()?;

    let tmpdir = tempfile::tempdir()?;
    let out_path = tmpdir.path().join("roundtripped.xlsx");
    std::fs::write(&out_path, &saved)?;

    let report = xlsx_diff::diff_workbooks(&fixture_path, &out_path)?;
    if report.has_at_least(xlsx_diff::Severity::Critical) {
        eprintln!(
            "Critical diffs detected for fixture {}",
            fixture_path.display()
        );
        for diff in report
            .differences
            .iter()
            .filter(|d| d.severity == xlsx_diff::Severity::Critical)
        {
            eprintln!("{diff}");
        }
        panic!(
            "fixture {} did not round-trip cleanly",
            fixture_path.display()
        );
    }

    Ok(())
}

