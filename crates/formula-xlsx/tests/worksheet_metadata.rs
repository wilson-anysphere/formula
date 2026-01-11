use std::io::Read;
use std::path::Path;

use formula_model::{CellRef, CellValue, Hyperlink, HyperlinkTarget, Range, TabColor, Workbook};
use formula_xlsx::{load_from_bytes, load_from_path, XlsxDocument};
use zip::ZipArchive;

#[test]
fn imports_row_and_col_properties_into_model() {
    let fixture = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/xlsx/metadata/row-col-properties.xlsx");
    let doc = load_from_path(&fixture).expect("load fixture");
    let sheet = doc.workbook.sheet_by_name("Sheet1").expect("sheet exists");

    let col1 = sheet.col_properties(1).expect("col 2 should have props");
    assert_eq!(col1.width, Some(25.0));
    assert!(!col1.hidden);

    let col2 = sheet.col_properties(2).expect("col 3 should have props");
    assert_eq!(col2.width, None);
    assert!(col2.hidden);

    let row1 = sheet.row_properties(1).expect("row 2 should have props");
    assert_eq!(row1.height, Some(30.0));
    assert!(!row1.hidden);

    let row2 = sheet.row_properties(2).expect("row 3 should have props");
    assert_eq!(row2.height, None);
    assert!(row2.hidden);
}

#[test]
fn imports_merge_cells_into_model() {
    let fixture = Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/merged-cells.xlsx");
    let bytes = std::fs::read(&fixture).expect("read fixture");
    let doc = load_from_bytes(&bytes).expect("load fixture");
    let sheet = doc.workbook.sheet_by_name("Sheet1").expect("sheet exists");

    let expected = Range::from_a1("A1:B2").expect("range");
    assert!(
        sheet.merged_regions.iter().any(|r| r.range == expected),
        "expected merged range {expected} to be present"
    );
}

#[test]
fn imports_hyperlinks_into_model() {
    let fixture = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/xlsx/hyperlinks/hyperlinks.xlsx");
    let doc = load_from_path(&fixture).expect("load fixture");
    let sheet = doc.workbook.sheet_by_name("Sheet1").expect("sheet exists");

    let expected = vec![
        Hyperlink {
            range: Range::from_a1("A1").expect("range"),
            target: HyperlinkTarget::ExternalUrl {
                uri: "https://example.com".to_string(),
            },
            display: Some("Example".to_string()),
            tooltip: Some("Go to example".to_string()),
            rel_id: Some("rId1".to_string()),
        },
        Hyperlink {
            range: Range::from_a1("A2").expect("range"),
            target: HyperlinkTarget::Internal {
                sheet: "Sheet2".to_string(),
                cell: CellRef::from_a1("B2").expect("cell"),
            },
            display: Some("Jump".to_string()),
            tooltip: None,
            rel_id: None,
        },
        Hyperlink {
            range: Range::from_a1("A3").expect("range"),
            target: HyperlinkTarget::Email {
                uri: "mailto:test@example.com".to_string(),
            },
            display: None,
            tooltip: None,
            rel_id: Some("rId2".to_string()),
        },
    ];

    assert_eq!(sheet.hyperlinks, expected);
}

#[test]
fn new_sheet_emits_metadata() {
    let mut workbook = Workbook::new();
    let sheet_id = workbook.add_sheet("Sheet1").expect("add sheet");

    {
        let sheet = workbook.sheet_mut(sheet_id).expect("sheet exists");
        sheet.set_value(CellRef::from_a1("A1").unwrap(), CellValue::String("Merged".to_string()));
        sheet.merge_range(Range::from_a1("A1:B2").unwrap())
            .expect("merge");

        sheet.set_col_width(1, Some(25.0));
        sheet.set_col_hidden(2, true);
        sheet.set_row_height(1, Some(30.0));
        sheet.set_row_hidden(2, true);

        sheet.zoom = 1.25;
        sheet.frozen_rows = 1;
        sheet.frozen_cols = 2;
        sheet.tab_color = Some(TabColor::rgb("FFFF0000"));

        let mut link = Hyperlink::for_cell(
            CellRef::from_a1("A1").unwrap(),
            HyperlinkTarget::ExternalUrl {
                uri: "https://example.com".to_string(),
            },
        );
        link.display = Some("Example".to_string());
        sheet.hyperlinks.push(link);
    }

    let doc = XlsxDocument::new(workbook);
    let bytes = doc.save_to_vec().expect("save");

    let cursor = std::io::Cursor::new(&bytes);
    let mut archive = ZipArchive::new(cursor).expect("zip open");

    let mut sheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")
        .expect("sheet xml")
        .read_to_string(&mut sheet_xml)
        .expect("read sheet xml");

    assert!(sheet_xml.contains("<sheetPr"), "missing <sheetPr>");
    assert!(
        sheet_xml.contains("tabColor") && sheet_xml.contains("FFFF0000"),
        "missing tabColor"
    );
    assert!(sheet_xml.contains("<sheetViews"), "missing <sheetViews>");
    assert!(sheet_xml.contains("zoomScale=\"125\""), "missing zoomScale");
    assert!(sheet_xml.contains("state=\"frozen\""), "missing frozen pane");
    assert!(sheet_xml.contains("<cols"), "missing <cols>");
    assert!(
        sheet_xml.contains("width=\"25\"") && sheet_xml.contains("customWidth=\"1\""),
        "missing column width"
    );
    assert!(
        sheet_xml.contains("ht=\"30\"") && sheet_xml.contains("customHeight=\"1\""),
        "missing row height"
    );
    assert!(sheet_xml.contains("hidden=\"1\""), "missing hidden col/row");
    assert!(sheet_xml.contains("<mergeCells"), "missing <mergeCells>");
    assert!(sheet_xml.contains("<hyperlinks"), "missing <hyperlinks>");

    let mut rels_xml = String::new();
    archive
        .by_name("xl/worksheets/_rels/sheet1.xml.rels")
        .expect("sheet rels")
        .read_to_string(&mut rels_xml)
        .expect("read sheet rels xml");
    assert!(rels_xml.contains("relationships/hyperlink"));
    assert!(rels_xml.contains("https://example.com"));
    assert!(rels_xml.contains("Id=\"rId1\""));
}
