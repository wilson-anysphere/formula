use formula_model::import::{import_csv_to_worksheet, CsvOptions};
use formula_model::{CellRef, CellValue};
use std::io::Cursor;

#[test]
fn csv_auto_detects_semicolon_delimiter() {
    let csv = "a;b;c\n1;2;3\n4;5;6\n";
    let sheet = import_csv_to_worksheet(
        1,
        "Data",
        Cursor::new(csv.as_bytes()),
        CsvOptions::default(),
    )
    .unwrap();

    assert_eq!(sheet.value(CellRef::new(0, 0)), CellValue::Number(1.0));
    assert_eq!(sheet.value(CellRef::new(0, 1)), CellValue::Number(2.0));
    assert_eq!(sheet.value(CellRef::new(0, 2)), CellValue::Number(3.0));
}

#[test]
fn csv_auto_detects_tab_delimiter() {
    let csv = "a\tb\tc\n1\t2\t3\n";
    let sheet = import_csv_to_worksheet(
        1,
        "Data",
        Cursor::new(csv.as_bytes()),
        CsvOptions::default(),
    )
    .unwrap();

    assert_eq!(sheet.value(CellRef::new(0, 0)), CellValue::Number(1.0));
    assert_eq!(sheet.value(CellRef::new(0, 1)), CellValue::Number(2.0));
    assert_eq!(sheet.value(CellRef::new(0, 2)), CellValue::Number(3.0));
}

#[test]
fn csv_auto_detects_pipe_delimiter() {
    let csv = "a|b|c\n1|2|3\n";
    let sheet = import_csv_to_worksheet(
        1,
        "Data",
        Cursor::new(csv.as_bytes()),
        CsvOptions::default(),
    )
    .unwrap();

    assert_eq!(sheet.value(CellRef::new(0, 0)), CellValue::Number(1.0));
    assert_eq!(sheet.value(CellRef::new(0, 1)), CellValue::Number(2.0));
    assert_eq!(sheet.value(CellRef::new(0, 2)), CellValue::Number(3.0));
}

#[test]
fn csv_auto_detect_ignores_commas_inside_quoted_strings() {
    // If delimiter detection naively counts commas inside quotes, it may incorrectly pick `,`
    // over the real `;` delimiter.
    let csv = "id;text\n1;\"a,b\"\n2;\"c,d\"\n";
    let sheet = import_csv_to_worksheet(
        1,
        "Data",
        Cursor::new(csv.as_bytes()),
        CsvOptions::default(),
    )
    .unwrap();

    assert_eq!(sheet.value(CellRef::new(0, 0)), CellValue::Number(1.0));
    assert_eq!(
        sheet.value(CellRef::new(0, 1)),
        CellValue::String("a,b".to_string())
    );
    assert_eq!(sheet.value(CellRef::new(1, 0)), CellValue::Number(2.0));
    assert_eq!(
        sheet.value(CellRef::new(1, 1)),
        CellValue::String("c,d".to_string())
    );
}

#[test]
fn csv_auto_detect_respects_excel_sep_directive() {
    // Excel supports `sep=<delimiter>` as a special first line that chooses the delimiter and is
    // not treated as a header/data row.
    let csv = "sep=;\na;b\n1;2\n";
    let sheet = import_csv_to_worksheet(
        1,
        "Data",
        Cursor::new(csv.as_bytes()),
        CsvOptions::default(),
    )
    .unwrap();

    assert_eq!(sheet.value(CellRef::new(0, 0)), CellValue::Number(1.0));
    assert_eq!(sheet.value(CellRef::new(0, 1)), CellValue::Number(2.0));
}

#[test]
fn csv_auto_detect_respects_excel_sep_directive_tab() {
    // Excel also supports `sep=<delimiter>` with tab as the delimiter. This should override the
    // content-based delimiter sniffing even if the remaining rows contain commas.
    let csv = "sep=\t\ncol1,col2\nhello,world\n";
    let sheet = import_csv_to_worksheet(
        1,
        "Data",
        Cursor::new(csv.as_bytes()),
        CsvOptions::default(),
    )
    .unwrap();

    assert_eq!(
        sheet.value(CellRef::new(0, 0)),
        CellValue::String("hello,world".to_string())
    );
    assert_eq!(sheet.value(CellRef::new(0, 1)), CellValue::Empty);
}

#[test]
fn csv_auto_detect_respects_excel_sep_directive_with_space_after_equals() {
    // Be tolerant of a single extra space after `sep=` (seen in some CSV generators).
    let csv = "sep= ;\na;b\n1;2\n";
    let sheet = import_csv_to_worksheet(
        1,
        "Data",
        Cursor::new(csv.as_bytes()),
        CsvOptions::default(),
    )
    .unwrap();

    assert_eq!(sheet.value(CellRef::new(0, 0)), CellValue::Number(1.0));
    assert_eq!(sheet.value(CellRef::new(0, 1)), CellValue::Number(2.0));
}

#[test]
fn csv_auto_detect_respects_excel_sep_directive_with_utf8_bom() {
    // Excel-exported CSVs can include a UTF-8 BOM; we should still detect and honor the `sep=`
    // directive.
    let csv = "\u{FEFF}sep=;\na;b\n1;2\n";
    let sheet = import_csv_to_worksheet(
        1,
        "Data",
        Cursor::new(csv.as_bytes()),
        CsvOptions::default(),
    )
    .unwrap();

    assert_eq!(sheet.value(CellRef::new(0, 0)), CellValue::Number(1.0));
    assert_eq!(sheet.value(CellRef::new(0, 1)), CellValue::Number(2.0));
}

#[test]
fn csv_auto_detect_prefers_semicolon_when_decimal_separator_is_comma_and_no_header() {
    // In locales where `,` is the decimal separator, semicolon-delimited CSVs often contain many
    // commas inside numbers. When delimiter detection is ambiguous (both `;` and `,` yield a
    // consistent column count), prefer `;` when `decimal_separator` is `,`.
    let csv = "1,23;4,56\n7,89;0,12\n";
    let options = CsvOptions {
        has_header: false,
        decimal_separator: ',',
        ..CsvOptions::default()
    };
    let sheet = import_csv_to_worksheet(1, "Data", Cursor::new(csv.as_bytes()), options).unwrap();

    assert_eq!(sheet.value(CellRef::new(0, 0)), CellValue::Number(1.23));
    assert_eq!(sheet.value(CellRef::new(0, 1)), CellValue::Number(4.56));
    assert_eq!(sheet.value(CellRef::new(1, 0)), CellValue::Number(7.89));
    assert_eq!(sheet.value(CellRef::new(1, 1)), CellValue::Number(0.12));
}
