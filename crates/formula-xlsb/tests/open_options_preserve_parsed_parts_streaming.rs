use std::path::Path;

use formula_xlsb::{OpenOptions, XlsbWorkbook};
use pretty_assertions::assert_eq;

fn fixture_path() -> String {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures/simple.xlsb")
        .to_string_lossy()
        .into_owned()
}

#[test]
fn preserve_parsed_parts_controls_preserved_bytes_without_changing_parsed_results() {
    let path = fixture_path();

    let opts_preserve = OpenOptions {
        preserve_unknown_parts: false,
        preserve_parsed_parts: true,
        preserve_worksheets: false,
        decode_formulas: true,
    };
    let opts_stream = OpenOptions {
        preserve_unknown_parts: false,
        preserve_parsed_parts: false,
        preserve_worksheets: false,
        decode_formulas: true,
    };

    let wb_preserve = XlsbWorkbook::open_with_options(&path, opts_preserve).expect("open xlsb");
    let wb_stream = XlsbWorkbook::open_with_options(&path, opts_stream).expect("open xlsb");

    // Ensure both paths produce identical derived workbook metadata.
    assert_eq!(wb_preserve.sheet_metas(), wb_stream.sheet_metas());
    assert_eq!(wb_preserve.workbook_properties(), wb_stream.workbook_properties());
    assert_eq!(wb_preserve.shared_strings(), wb_stream.shared_strings());
    assert_eq!(
        wb_preserve.shared_strings_table().len(),
        wb_stream.shared_strings_table().len()
    );
    assert_eq!(wb_preserve.defined_names(), wb_stream.defined_names());

    // Raw bytes for parsed parts should only be kept when requested.
    assert!(wb_preserve.preserved_parts().contains_key("xl/workbook.bin"));
    assert!(wb_preserve
        .preserved_parts()
        .contains_key("xl/sharedStrings.bin"));

    assert!(!wb_stream.preserved_parts().contains_key("xl/workbook.bin"));
    assert!(!wb_stream
        .preserved_parts()
        .contains_key("xl/sharedStrings.bin"));
}

