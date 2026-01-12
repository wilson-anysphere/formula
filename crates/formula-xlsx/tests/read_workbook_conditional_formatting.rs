use formula_model::{CfRuleKind, CfRuleSchema};
use formula_xlsx::{parse_worksheet_conditional_formatting_streaming, read_workbook};

fn fixture(path: &str) -> String {
    let dir = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join(path);
    dir.to_string_lossy().to_string()
}

#[test]
fn read_workbook_populates_conditional_formatting_office2007() {
    let wb = read_workbook(fixture("conditional_formatting_2007.xlsx")).expect("read workbook");
    let sheet = wb.sheets.first().expect("workbook has at least one sheet");
    assert_eq!(sheet.conditional_formatting_rules.len(), 4);
    assert_eq!(sheet.conditional_formatting_dxfs.len(), 1);
}

#[test]
fn read_workbook_populates_conditional_formatting_x14() {
    let wb = read_workbook(fixture("conditional_formatting_x14.xlsx")).expect("read workbook");
    let sheet = wb.sheets.first().expect("workbook has at least one sheet");
    assert_eq!(sheet.conditional_formatting_rules.len(), 1);

    let rule = &sheet.conditional_formatting_rules[0];
    assert_eq!(rule.schema, CfRuleSchema::X14);
    match &rule.kind {
        CfRuleKind::DataBar(db) => {
            assert_eq!(
                format!("{:08X}", db.color.unwrap().argb().unwrap_or(0)),
                "FF638EC6"
            );
            assert_eq!(db.min_length, Some(0));
            assert_eq!(db.max_length, Some(100));
            assert_eq!(db.gradient, Some(false));
        }
        other => panic!("expected DataBar rule, got {other:?}"),
    }
}

#[test]
fn streaming_extractor_supports_prefixed_worksheet_root() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <x:conditionalFormatting sqref="A1:A1">
    <x:cfRule type="expression" priority="1">
      <x:formula>A1&gt;0</x:formula>
    </x:cfRule>
  </x:conditionalFormatting>
</x:worksheet>"#;

    let parsed = parse_worksheet_conditional_formatting_streaming(xml).expect("parse cf");
    assert_eq!(parsed.rules.len(), 1);
}

