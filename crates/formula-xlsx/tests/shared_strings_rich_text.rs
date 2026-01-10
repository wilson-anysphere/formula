use formula_model::rich_text::{RichText, RichTextRunStyle, Underline};
use formula_model::Color;
use formula_xlsx::shared_strings::read_shared_strings_from_xlsx;
use formula_xlsx::shared_strings::{
    parse_shared_strings_xml, write_shared_strings_to_xlsx, write_shared_strings_xml,
};

#[test]
fn parses_rich_text_runs_from_shared_strings_xml() {
    let xml = include_str!("fixtures/sharedStrings_rich.xml");
    let shared = parse_shared_strings_xml(xml).expect("parse sharedStrings.xml");

    assert_eq!(shared.items.len(), 2);

    let rich = &shared.items[0];
    assert_eq!(rich.text, "Hello Bold Italic");
    assert_eq!(rich.runs.len(), 3);

    assert_eq!(rich.slice_run_text(&rich.runs[0]), "Hello ");
    assert_eq!(rich.slice_run_text(&rich.runs[1]), "Bold");
    assert_eq!(rich.slice_run_text(&rich.runs[2]), " Italic");

    assert_eq!(rich.runs[0].style.font.as_deref(), Some("Calibri"));
    assert_eq!(rich.runs[0].style.size_100pt, Some(1100));
    assert_eq!(rich.runs[0].style.color, Some(Color::new_argb(0xFF000000)));

    assert_eq!(rich.runs[1].style.bold, Some(true));
    assert_eq!(rich.runs[1].style.color, Some(Color::new_argb(0xFFFF0000)));

    assert_eq!(rich.runs[2].style.italic, Some(true));
    assert_eq!(rich.runs[2].style.underline, Some(Underline::Single));
}

#[test]
fn shared_strings_xml_round_trip_preserves_rich_runs() {
    let xml = include_str!("fixtures/sharedStrings_rich.xml");
    let shared = parse_shared_strings_xml(xml).expect("parse sharedStrings.xml");

    let written = write_shared_strings_xml(&shared).expect("write sharedStrings.xml");
    let reparsed = parse_shared_strings_xml(&written).expect("re-parse written xml");

    assert_eq!(shared, reparsed);
    assert!(written.contains("<b"));
    assert!(written.contains("<u"));
}

#[test]
fn xlsx_round_trip_preserves_shared_strings() {
    let fixture_path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/tests/fixtures/rich_text_shared_strings.xlsx"
    );

    let shared =
        read_shared_strings_from_xlsx(fixture_path).expect("read shared strings from xlsx");

    // Write to a temp output workbook and re-read.
    let dir = tempfile::tempdir().expect("tempdir");
    let out_path = dir.path().join("out.xlsx");
    write_shared_strings_to_xlsx(fixture_path, &out_path, &shared).expect("write xlsx");

    let reloaded = read_shared_strings_from_xlsx(&out_path).expect("re-read shared strings");
    assert_eq!(shared, reloaded);
}

#[test]
fn editing_preservation_contract_example() {
    // This models the UI MVP rule: if a rich value is edited as plain text and the
    // user does not change the text, rich runs should be preserved.
    let original = RichText::from_segments(vec![
        ("Hello ".to_string(), RichTextRunStyle::default()),
        (
            "Bold".to_string(),
            RichTextRunStyle {
                bold: Some(true),
                ..Default::default()
            },
        ),
    ]);

    let edited_same = apply_plain_text_edit(Some(original.clone()), "Hello Bold");
    assert_eq!(edited_same, Some(original.clone()));

    let edited_changed = apply_plain_text_edit(Some(original), "Hello Changed");
    assert_eq!(edited_changed, Some(RichText::new("Hello Changed")));

    let edited_from_none = apply_plain_text_edit(None, "X");
    assert_eq!(edited_from_none, Some(RichText::new("X")));
}

fn apply_plain_text_edit(original: Option<RichText>, edited: &str) -> Option<RichText> {
    match original {
        Some(rich) if rich.text == edited => Some(rich),
        _ => Some(RichText::new(edited)),
    }
}
