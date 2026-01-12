use std::io::Write;

use formula_model::{
    BorderStyle, CellRef, Color, FillPattern, HorizontalAlignment, VerticalAlignment,
};

mod common;

use common::xls_fixture_builder;

fn import_fixture(bytes: &[u8]) -> formula_xls::XlsImportResult {
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(bytes).expect("write xls bytes");
    formula_xls::import_xls_path(tmp.path()).expect("import xls")
}

#[test]
fn imports_rich_biff_cell_styles() {
    let bytes = xls_fixture_builder::build_rich_styles_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("Styles")
        .expect("Styles sheet missing");

    let a1 = CellRef::from_a1("A1").unwrap();
    let cell = sheet.cell(a1).expect("A1 missing");
    assert_ne!(cell.style_id, 0, "expected a non-default style id");

    let style = result
        .workbook
        .styles
        .get(cell.style_id)
        .expect("style missing");

    // Font.
    let font = style.font.as_ref().expect("expected font");
    assert_eq!(font.name.as_deref(), Some("Courier New"));
    assert_eq!(font.size_100pt, Some(1000));
    assert!(font.bold);
    assert!(font.italic);
    assert!(font.underline);
    assert!(font.strike);
    assert_eq!(font.color, Some(Color::Argb(0xFFFF_0000)));

    // Fill.
    let fill = style.fill.as_ref().expect("expected fill");
    assert!(matches!(fill.pattern, FillPattern::Solid));
    assert_eq!(fill.fg_color, Some(Color::Argb(0xFFFF_0000)));
    assert_eq!(fill.bg_color, Some(Color::Argb(0xFF00_FF00)));

    // Border.
    let border = style.border.as_ref().expect("expected border");
    assert_eq!(border.left.style, BorderStyle::Thin);
    assert_eq!(border.left.color, Some(Color::Argb(0xFF00_FF00)));
    assert!(border.diagonal_up);
    assert!(!border.diagonal_down);
    assert_eq!(border.diagonal.style, BorderStyle::Thin);

    // Alignment.
    let alignment = style.alignment.as_ref().expect("expected alignment");
    assert_eq!(alignment.horizontal, Some(HorizontalAlignment::Center));
    assert_eq!(alignment.vertical, Some(VerticalAlignment::Top));
    assert!(alignment.wrap_text);
    assert_eq!(alignment.rotation, Some(45));
    assert_eq!(alignment.indent, Some(2));

    // Protection.
    let protection = style.protection.as_ref().expect("expected protection");
    assert!(!protection.locked);
    assert!(protection.hidden);

    // Number format.
    assert_eq!(style.number_format.as_deref(), Some("0.00%"));

    // Non-solid fill pattern should map to a valid OOXML patternType token so XLSX round-tripping
    // does not emit invalid XML.
    let b1 = CellRef::from_a1("B1").unwrap();
    let cell_b1 = sheet.cell(b1).expect("B1 missing");
    assert_ne!(cell_b1.style_id, 0, "expected a non-default style id");
    let style_b1 = result
        .workbook
        .styles
        .get(cell_b1.style_id)
        .expect("B1 style missing");
    let fill_b1 = style_b1.fill.as_ref().expect("expected B1 fill");
    assert!(
        matches!(fill_b1.pattern, FillPattern::Other(ref v) if v == "mediumGray"),
        "unexpected fill pattern: {:?}",
        fill_b1.pattern
    );
}
