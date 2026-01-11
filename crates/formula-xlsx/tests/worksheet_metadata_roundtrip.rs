use std::path::{Path, PathBuf};

use formula_model::{CellRef, Hyperlink, HyperlinkTarget, Range, SheetVisibility, TabColor, Workbook};
use formula_xlsx::load_from_bytes;
use pretty_assertions::assert_eq;

fn assert_no_critical_diffs(expected: &Path, actual_bytes: &[u8]) {
    let tmpdir = tempfile::tempdir().expect("tempdir");
    let out = tmpdir.path().join("roundtripped.xlsx");
    std::fs::write(&out, actual_bytes).expect("write roundtrip");

    let report = xlsx_diff::diff_workbooks(expected, &out).expect("xlsx diff");
    if report.has_at_least(xlsx_diff::Severity::Critical) {
        eprintln!("Critical diffs detected for fixture {}", expected.display());
        for diff in report
            .differences
            .iter()
            .filter(|d| d.severity == xlsx_diff::Severity::Critical)
        {
            eprintln!("{diff}");
        }
        panic!("fixture {} did not round-trip cleanly", expected.display());
    }
}

fn fixture_path(rel: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join(rel)
}

#[test]
fn hyperlinks_are_loaded_into_model_and_roundtrip_preserves_parts() {
    let fixture = fixture_path("../../fixtures/xlsx/hyperlinks/hyperlinks.xlsx");
    let bytes = std::fs::read(&fixture).expect("fixture exists");

    let doc = load_from_bytes(&bytes).expect("load");
    let sheet1 = doc.workbook.sheet_by_name("Sheet1").expect("Sheet1 present");

    assert_eq!(
        sheet1.hyperlinks,
        vec![
            Hyperlink {
                range: Range::new(CellRef::new(0, 0), CellRef::new(0, 0)),
                target: HyperlinkTarget::ExternalUrl {
                    uri: "https://example.com".to_string()
                },
                display: Some("Example".to_string()),
                tooltip: Some("Go to example".to_string()),
                rel_id: Some("rId1".to_string()),
            },
            Hyperlink {
                range: Range::new(CellRef::new(1, 0), CellRef::new(1, 0)),
                target: HyperlinkTarget::Internal {
                    sheet: "Sheet2".to_string(),
                    cell: CellRef::new(1, 1),
                },
                display: Some("Jump".to_string()),
                tooltip: None,
                rel_id: None,
            },
            Hyperlink {
                range: Range::new(CellRef::new(2, 0), CellRef::new(2, 0)),
                target: HyperlinkTarget::Email {
                    uri: "mailto:test@example.com".to_string(),
                },
                display: None,
                tooltip: None,
                rel_id: Some("rId2".to_string()),
            },
        ]
    );

    let saved = doc.save_to_vec().expect("save");
    assert_no_critical_diffs(&fixture, &saved);
}

#[test]
fn merged_cells_are_loaded_into_model_and_roundtrip_preserves_parts() {
    let fixture = fixture_path("tests/fixtures/merged-cells.xlsx");
    let bytes = std::fs::read(&fixture).expect("fixture exists");

    let doc = load_from_bytes(&bytes).expect("load");
    let sheet1 = doc.workbook.sheet_by_name("Sheet1").expect("Sheet1 present");

    assert_eq!(sheet1.merged_regions.region_count(), 1);
    assert_eq!(
        sheet1.merged_regions.iter().next().unwrap().range,
        Range::new(CellRef::new(0, 0), CellRef::new(1, 1))
    );

    let saved = doc.save_to_vec().expect("save");
    assert_no_critical_diffs(&fixture, &saved);
}

#[test]
fn sheet_visibility_ids_and_tab_color_are_loaded_into_model() {
    let fixture = fixture_path("tests/fixtures/sheet-metadata.xlsx");
    let bytes = std::fs::read(&fixture).expect("fixture exists");

    let doc = load_from_bytes(&bytes).expect("load");
    assert_eq!(doc.workbook.sheets.len(), 3);

    let visible = doc.workbook.sheet_by_name("Visible").expect("Visible");
    assert_eq!(visible.visibility, SheetVisibility::Visible);
    assert_eq!(visible.xlsx_sheet_id, Some(1));
    assert_eq!(visible.xlsx_rel_id.as_deref(), Some("rId1"));
    assert_eq!(visible.tab_color, Some(TabColor::rgb("FFFF0000")));

    let hidden = doc.workbook.sheet_by_name("Hidden").expect("Hidden");
    assert_eq!(hidden.visibility, SheetVisibility::Hidden);
    assert_eq!(hidden.xlsx_sheet_id, Some(2));
    assert_eq!(hidden.xlsx_rel_id.as_deref(), Some("rId2"));

    let very_hidden = doc
        .workbook
        .sheet_by_name("VeryHidden")
        .expect("VeryHidden");
    assert_eq!(very_hidden.visibility, SheetVisibility::VeryHidden);
    assert_eq!(very_hidden.xlsx_sheet_id, Some(3));
    assert_eq!(very_hidden.xlsx_rel_id.as_deref(), Some("rId3"));

    let saved = doc.save_to_vec().expect("save");
    assert_no_critical_diffs(&fixture, &saved);
}

#[test]
fn editing_merges_updates_merge_cells_block() {
    let fixture = fixture_path("tests/fixtures/merged-cells.xlsx");
    let bytes = std::fs::read(&fixture).expect("fixture exists");
    let mut doc = load_from_bytes(&bytes).expect("load");

    let sheet_id = doc.workbook.sheet_by_name("Sheet1").unwrap().id;
    let sheet = doc.workbook.sheet_mut(sheet_id).unwrap();
    sheet
        .merge_range(Range::new(CellRef::new(2, 2), CellRef::new(3, 3)))
        .expect("merge ok");

    let saved = doc.save_to_vec().expect("save");
    let reloaded = load_from_bytes(&saved).expect("reload");
    let sheet1 = reloaded.workbook.sheet_by_name("Sheet1").unwrap();
    assert_eq!(sheet1.merged_regions.region_count(), 2);
}

#[test]
fn editing_hyperlinks_updates_xml_and_relationships() {
    let fixture = fixture_path("../../fixtures/xlsx/hyperlinks/hyperlinks.xlsx");
    let bytes = std::fs::read(&fixture).expect("fixture exists");
    let mut doc = load_from_bytes(&bytes).expect("load");

    let sheet_id = doc.workbook.sheet_by_name("Sheet1").unwrap().id;
    let sheet = doc.workbook.sheet_mut(sheet_id).unwrap();
    assert_eq!(sheet.hyperlinks.len(), 3);

    match &mut sheet.hyperlinks[0].target {
        HyperlinkTarget::ExternalUrl { uri } => {
            *uri = "https://example.org".to_string();
        }
        other => panic!("unexpected target: {other:?}"),
    }

    let saved = doc.save_to_vec().expect("save");
    let reloaded = load_from_bytes(&saved).expect("reload");
    let sheet1 = reloaded.workbook.sheet_by_name("Sheet1").unwrap();
    assert_eq!(sheet1.hyperlinks.len(), 3);
    match &sheet1.hyperlinks[0].target {
        HyperlinkTarget::ExternalUrl { uri } => assert_eq!(uri, "https://example.org"),
        other => panic!("unexpected target: {other:?}"),
    }
    assert_eq!(sheet1.hyperlinks[0].rel_id.as_deref(), Some("rId1"));
}

#[test]
fn new_documents_write_worksheet_metadata() {
    let mut workbook = Workbook::new();
    let sheet1_id = workbook.add_sheet("Sheet1").unwrap();
    let sheet2_id = workbook.add_sheet("HiddenSheet").unwrap();

    {
        let sheet1 = workbook.sheet_mut(sheet1_id).unwrap();
        sheet1.tab_color = Some(TabColor::rgb("FF00FF00"));
        sheet1
            .merge_range(Range::new(CellRef::new(0, 0), CellRef::new(1, 1)))
            .expect("merge ok");
        sheet1.hyperlinks.push(Hyperlink {
            range: Range::new(CellRef::new(0, 0), CellRef::new(0, 0)),
            target: HyperlinkTarget::ExternalUrl {
                uri: "https://example.com".to_string(),
            },
            display: None,
            tooltip: None,
            rel_id: Some("rId1".to_string()),
        });
    }

    {
        let sheet2 = workbook.sheet_mut(sheet2_id).unwrap();
        sheet2.visibility = SheetVisibility::Hidden;
    }

    let doc = formula_xlsx::XlsxDocument::new(workbook);
    let bytes = doc.save_to_vec().expect("save");

    let loaded = load_from_bytes(&bytes).expect("reload");
    let sheet1 = loaded.workbook.sheet_by_name("Sheet1").unwrap();
    assert_eq!(sheet1.tab_color, Some(TabColor::rgb("FF00FF00")));
    assert_eq!(sheet1.merged_regions.region_count(), 1);
    assert_eq!(sheet1.hyperlinks.len(), 1);

    let hidden = loaded.workbook.sheet_by_name("HiddenSheet").unwrap();
    assert_eq!(hidden.visibility, SheetVisibility::Hidden);
}
