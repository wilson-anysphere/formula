use formula_columnar::ColumnType as ColumnarType;
use formula_model::import::{
    import_csv_to_columnar_table, import_csv_to_worksheet, CsvDateOrder, CsvOptions,
    CsvTextEncoding, CsvTimestampTzPolicy,
};
use formula_model::{CellRef, CellValue};
use std::io::Cursor;

#[test]
fn csv_import_streams_into_columnar_backed_worksheet() {
    let csv = concat!(
        "id,amount,ratio,flag,ts,category\n",
        "1,$12.34,50%,true,1970-01-02,A\n",
        "2,$0.01,12.5%,false,1970-01-03,B\n",
    );

    let sheet = import_csv_to_worksheet(
        1,
        "Data",
        Cursor::new(csv.as_bytes()),
        CsvOptions::default(),
    )
    .unwrap();

    assert_eq!(sheet.value(CellRef::new(0, 0)), CellValue::Number(1.0));
    assert_eq!(sheet.value(CellRef::new(0, 1)), CellValue::Number(12.34));
    assert_eq!(sheet.value(CellRef::new(0, 2)), CellValue::Number(0.5));
    assert_eq!(sheet.value(CellRef::new(0, 3)), CellValue::Boolean(true));
    assert_eq!(
        sheet.value(CellRef::new(0, 4)),
        CellValue::Number(86_400_000.0)
    );
    assert_eq!(
        sheet.value(CellRef::new(0, 5)),
        CellValue::String("A".to_string())
    );

    assert_eq!(sheet.value(CellRef::new(1, 0)), CellValue::Number(2.0));
    assert_eq!(sheet.value(CellRef::new(1, 1)), CellValue::Number(0.01));
    assert_eq!(sheet.value(CellRef::new(1, 2)), CellValue::Number(0.125));
    assert_eq!(sheet.value(CellRef::new(1, 3)), CellValue::Boolean(false));
    assert_eq!(
        sheet.value(CellRef::new(1, 4)),
        CellValue::Number(172_800_000.0)
    );
    assert_eq!(
        sheet.value(CellRef::new(1, 5)),
        CellValue::String("B".to_string())
    );
}

#[test]
fn csv_import_rfc4180_quotes_newlines_and_crlf() {
    let csv = concat!(
        "id,text\r\n",
        "1,\"hello, world\"\r\n",
        "2,\"line1\r\nline2\"\r\n",
        "3,\"he said \"\"hi\"\"\"\r\n",
    );

    let sheet = import_csv_to_worksheet(
        1,
        "Data",
        Cursor::new(csv.as_bytes()),
        CsvOptions::default(),
    )
    .unwrap();

    assert_eq!(
        sheet.value(CellRef::new(0, 1)),
        CellValue::String("hello, world".to_string())
    );
    assert_eq!(
        sheet.value(CellRef::new(1, 1)),
        CellValue::String("line1\r\nline2".to_string())
    );
    assert_eq!(
        sheet.value(CellRef::new(2, 1)),
        CellValue::String("he said \"hi\"".to_string())
    );
}

#[test]
fn csv_import_supports_cr_only_line_endings() {
    let csv = "id,text\r1,hello\r2,world\r";
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
        CellValue::String("hello".to_string())
    );
    assert_eq!(sheet.value(CellRef::new(1, 0)), CellValue::Number(2.0));
    assert_eq!(
        sheet.value(CellRef::new(1, 1)),
        CellValue::String("world".to_string())
    );
}

#[test]
fn csv_import_trailing_delimiter_produces_empty_field() {
    let csv = "a,b,c\n1,2,\n";
    let sheet = import_csv_to_worksheet(
        1,
        "Data",
        Cursor::new(csv.as_bytes()),
        CsvOptions::default(),
    )
    .unwrap();

    // Trailing delimiter => third field exists but is empty.
    assert_eq!(sheet.value(CellRef::new(0, 2)), CellValue::Empty);
}

#[test]
fn csv_import_handles_utf8_inside_quoted_fields() {
    let csv = "id,text\n1,\"こんにちは,世界\"\n";
    let sheet = import_csv_to_worksheet(
        1,
        "Data",
        Cursor::new(csv.as_bytes()),
        CsvOptions::default(),
    )
    .unwrap();
    assert_eq!(
        sheet.value(CellRef::new(0, 1)),
        CellValue::String("こんにちは,世界".to_string())
    );
}

#[test]
fn csv_import_preserves_whitespace_in_string_values() {
    let csv = "text\n  hello  \n";
    let sheet = import_csv_to_worksheet(
        1,
        "Data",
        Cursor::new(csv.as_bytes()),
        CsvOptions::default(),
    )
    .unwrap();
    assert_eq!(
        sheet.value(CellRef::new(0, 0)),
        CellValue::String("  hello  ".to_string())
    );
}

#[test]
fn csv_import_header_only_defaults_columns_to_string() {
    let csv = "a,b,c";
    let table = import_csv_to_columnar_table(Cursor::new(csv.as_bytes()), CsvOptions::default())
        .expect("import header-only csv");
    assert_eq!(table.row_count(), 0);
    assert_eq!(table.column_count(), 3);
    assert_eq!(table.schema()[0].column_type, ColumnarType::String);
    assert_eq!(table.schema()[1].column_type, ColumnarType::String);
    assert_eq!(table.schema()[2].column_type, ColumnarType::String);
}

#[test]
fn csv_import_infers_string_type_for_all_empty_column() {
    let csv = "a,b\n1,\n2,\n";
    let table = import_csv_to_columnar_table(Cursor::new(csv.as_bytes()), CsvOptions::default())
        .expect("import csv");
    assert_eq!(table.schema()[0].column_type, ColumnarType::Number);
    assert_eq!(table.schema()[1].column_type, ColumnarType::String);
}

#[test]
fn csv_import_supports_utf16le_bom_tab_delimited_text() {
    // Excel's "Unicode Text" export is UTF-16LE with a BOM and (typically) tab-delimited.
    let tsv = "id\ttext\r\n1\thello\r\n2\tworld\r\n";
    let mut bytes = vec![0xFF, 0xFE];
    for unit in tsv.encode_utf16() {
        bytes.extend_from_slice(&unit.to_le_bytes());
    }

    let sheet = import_csv_to_worksheet(1, "Data", Cursor::new(bytes), CsvOptions::default())
        .expect("import utf16 tsv");

    assert_eq!(sheet.value(CellRef::new(0, 0)), CellValue::Number(1.0));
    assert_eq!(
        sheet.value(CellRef::new(0, 1)),
        CellValue::String("hello".to_string())
    );
    assert_eq!(sheet.value(CellRef::new(1, 0)), CellValue::Number(2.0));
    assert_eq!(
        sheet.value(CellRef::new(1, 1)),
        CellValue::String("world".to_string())
    );
}

#[test]
fn csv_import_supports_utf16be_bom_tab_delimited_text() {
    let tsv = "id\ttext\r\n1\thello\r\n2\tworld\r\n";
    let mut bytes = vec![0xFE, 0xFF];
    for unit in tsv.encode_utf16() {
        bytes.extend_from_slice(&unit.to_be_bytes());
    }

    let sheet = import_csv_to_worksheet(1, "Data", Cursor::new(bytes), CsvOptions::default())
        .expect("import utf16be tsv");

    assert_eq!(sheet.value(CellRef::new(0, 0)), CellValue::Number(1.0));
    assert_eq!(
        sheet.value(CellRef::new(0, 1)),
        CellValue::String("hello".to_string())
    );
    assert_eq!(sheet.value(CellRef::new(1, 0)), CellValue::Number(2.0));
    assert_eq!(
        sheet.value(CellRef::new(1, 1)),
        CellValue::String("world".to_string())
    );
}

#[test]
fn csv_import_utf16le_honors_excel_sep_directive() {
    // Excel supports a special first line `sep=<delimiter>` which explicitly specifies the CSV
    // delimiter and is not treated as a header/data row.
    //
    // Make the sample ambiguous for delimiter sniffing (commas appear inside unquoted fields).
    // Without honoring `sep=;`, the sniffer would prefer `,` over `;` and parse the header/rows
    // incorrectly.
    let csv = "sep=;\r\na;b\r\n1,hello;world\r\n2,foo;bar\r\n";
    let mut bytes = vec![0xFF, 0xFE];
    for unit in csv.encode_utf16() {
        bytes.extend_from_slice(&unit.to_le_bytes());
    }

    let table =
        import_csv_to_columnar_table(Cursor::new(bytes.clone()), CsvOptions::default()).unwrap();
    assert_eq!(table.schema()[0].name, "a");
    assert_eq!(table.schema()[1].name, "b");
    assert_eq!(table.row_count(), 2);

    let sheet = import_csv_to_worksheet(1, "Data", Cursor::new(bytes), CsvOptions::default())
        .expect("import utf16 csv with sep directive");
    assert_eq!(
        sheet.value(CellRef::new(0, 0)),
        CellValue::String("1,hello".to_string())
    );
    assert_eq!(
        sheet.value(CellRef::new(0, 1)),
        CellValue::String("world".to_string())
    );
    assert_eq!(
        sheet.value(CellRef::new(1, 0)),
        CellValue::String("2,foo".to_string())
    );
    assert_eq!(
        sheet.value(CellRef::new(1, 1)),
        CellValue::String("bar".to_string())
    );
}

#[test]
fn csv_import_supports_utf16le_tab_delimited_text_without_bom() {
    let tsv = "id\ttext\r\n1\thello\r\n2\tworld\r\n";
    let mut bytes = Vec::new();
    for unit in tsv.encode_utf16() {
        bytes.extend_from_slice(&unit.to_le_bytes());
    }

    let sheet = import_csv_to_worksheet(1, "Data", Cursor::new(bytes), CsvOptions::default())
        .expect("import utf16le tsv without bom");

    assert_eq!(sheet.value(CellRef::new(0, 0)), CellValue::Number(1.0));
    assert_eq!(
        sheet.value(CellRef::new(0, 1)),
        CellValue::String("hello".to_string())
    );
    assert_eq!(sheet.value(CellRef::new(1, 0)), CellValue::Number(2.0));
    assert_eq!(
        sheet.value(CellRef::new(1, 1)),
        CellValue::String("world".to_string())
    );
}

#[test]
fn csv_import_supports_utf16be_tab_delimited_text_without_bom() {
    let tsv = "id\ttext\r\n1\thello\r\n2\tworld\r\n";
    let mut bytes = Vec::new();
    for unit in tsv.encode_utf16() {
        bytes.extend_from_slice(&unit.to_be_bytes());
    }

    let sheet = import_csv_to_worksheet(1, "Data", Cursor::new(bytes), CsvOptions::default())
        .expect("import utf16be tsv without bom");

    assert_eq!(sheet.value(CellRef::new(0, 0)), CellValue::Number(1.0));
    assert_eq!(
        sheet.value(CellRef::new(0, 1)),
        CellValue::String("hello".to_string())
    );
    assert_eq!(sheet.value(CellRef::new(1, 0)), CellValue::Number(2.0));
    assert_eq!(
        sheet.value(CellRef::new(1, 1)),
        CellValue::String("world".to_string())
    );
}

#[test]
fn csv_import_supports_utf16le_tab_delimited_text_without_bom_mostly_non_ascii() {
    // Regression test for BOM-less UTF-16 detection: if the content is mostly non-ASCII, the NUL
    // byte ratio can be much lower than the typical ~50% seen in ASCII-heavy UTF-16.
    let left = "あ".repeat(200);
    let right = "い".repeat(200);
    let tsv = format!("{left}\t{right}\r\n");
    let mut bytes = Vec::new();
    for unit in tsv.encode_utf16() {
        bytes.extend_from_slice(&unit.to_le_bytes());
    }

    let sheet = import_csv_to_worksheet(
        1,
        "Data",
        Cursor::new(bytes),
        CsvOptions {
            has_header: false,
            ..CsvOptions::default()
        },
    )
    .expect("import utf16le tsv without bom");

    assert_eq!(sheet.value(CellRef::new(0, 0)), CellValue::String(left));
    assert_eq!(sheet.value(CellRef::new(0, 1)), CellValue::String(right));
}

#[test]
fn csv_import_supports_utf16be_tab_delimited_text_without_bom_mostly_non_ascii() {
    let left = "あ".repeat(200);
    let right = "い".repeat(200);
    let tsv = format!("{left}\t{right}\r\n");
    let mut bytes = Vec::new();
    for unit in tsv.encode_utf16() {
        bytes.extend_from_slice(&unit.to_be_bytes());
    }

    let sheet = import_csv_to_worksheet(
        1,
        "Data",
        Cursor::new(bytes),
        CsvOptions {
            has_header: false,
            ..CsvOptions::default()
        },
    )
    .expect("import utf16be tsv without bom");

    assert_eq!(sheet.value(CellRef::new(0, 0)), CellValue::String(left));
    assert_eq!(sheet.value(CellRef::new(0, 1)), CellValue::String(right));
}

#[test]
fn csv_import_handles_wide_rows() {
    let cols = 200usize;
    let header = (0..cols)
        .map(|i| format!("c{}", i + 1))
        .collect::<Vec<_>>()
        .join(",");
    let row = (0..cols)
        .map(|i| (i + 1).to_string())
        .collect::<Vec<_>>()
        .join(",");
    let csv = format!("{header}\n{row}\n");

    let sheet = import_csv_to_worksheet(
        1,
        "Data",
        Cursor::new(csv.as_bytes()),
        CsvOptions::default(),
    )
    .unwrap();

    assert_eq!(sheet.value(CellRef::new(0, 0)), CellValue::Number(1.0));
    assert_eq!(
        sheet.value(CellRef::new(0, (cols - 1) as u32)),
        CellValue::Number(cols as f64)
    );
}

#[test]
fn csv_import_type_inference_respects_locale_options() {
    let csv = "amount;ratio;date\n€1.234,50;12,5%;31/12/1970\n";
    let options = CsvOptions {
        delimiter: b';',
        decimal_separator: ',',
        date_order: CsvDateOrder::Dmy,
        ..CsvOptions::default()
    };
    let sheet = import_csv_to_worksheet(1, "Data", Cursor::new(csv.as_bytes()), options).unwrap();

    assert_eq!(sheet.value(CellRef::new(0, 0)), CellValue::Number(1234.5));
    assert_eq!(sheet.value(CellRef::new(0, 1)), CellValue::Number(0.125));
    assert_eq!(
        sheet.value(CellRef::new(0, 2)),
        CellValue::Number(31_449_600_000.0)
    );
}

#[test]
fn csv_import_autodetects_decimal_comma_locale_for_semicolon_csv() {
    // A common Excel locale configuration uses:
    // - `;` as the list separator / CSV delimiter
    // - `,` as the decimal separator
    //
    // When exporting CSV, Excel often omits the `sep=` directive, relying on locale defaults.
    // Import should infer that `,` is a decimal separator (so delimiter sniffing does not mistake
    // decimal commas for field separators) and parse values accordingly.
    let csv = "amount;ratio\n€1.234,50;12,5%\n";
    let sheet = import_csv_to_worksheet(
        1,
        "Data",
        Cursor::new(csv.as_bytes()),
        CsvOptions::default(),
    )
    .unwrap();

    assert_eq!(sheet.value(CellRef::new(0, 0)), CellValue::Number(1234.5));
    assert_eq!(sheet.value(CellRef::new(0, 1)), CellValue::Number(0.125));
}

#[test]
fn csv_import_autodetects_tab_delimiter_for_decimal_comma_values() {
    // Without a header row, delimiter sniffing can be ambiguous: decimal commas can make the input
    // look comma-delimited. Ensure we still pick the real delimiter (tab) when the sample suggests
    // a decimal-comma locale.
    let tsv = "1,23\t4,56\n7,89\t0,12\n";
    let sheet = import_csv_to_worksheet(
        1,
        "Data",
        Cursor::new(tsv.as_bytes()),
        CsvOptions {
            has_header: false,
            ..CsvOptions::default()
        },
    )
    .unwrap();

    assert_eq!(sheet.value(CellRef::new(0, 0)), CellValue::Number(1.23));
    assert_eq!(sheet.value(CellRef::new(0, 1)), CellValue::Number(4.56));
    assert_eq!(sheet.value(CellRef::new(1, 0)), CellValue::Number(7.89));
    assert_eq!(sheet.value(CellRef::new(1, 1)), CellValue::Number(0.12));
}

#[test]
fn csv_import_date_order_preference_changes_ambiguous_dates() {
    let csv = "d\n01/02/1970\n";

    let mdy = CsvOptions {
        date_order: CsvDateOrder::Mdy,
        ..CsvOptions::default()
    };
    let sheet = import_csv_to_worksheet(1, "Data", Cursor::new(csv.as_bytes()), mdy).unwrap();
    assert_eq!(
        sheet.value(CellRef::new(0, 0)),
        CellValue::Number(86_400_000.0)
    );

    let dmy = CsvOptions {
        date_order: CsvDateOrder::Dmy,
        ..CsvOptions::default()
    };
    let sheet = import_csv_to_worksheet(1, "Data", Cursor::new(csv.as_bytes()), dmy).unwrap();
    assert_eq!(
        sheet.value(CellRef::new(0, 0)),
        CellValue::Number(2_678_400_000.0)
    );
}

#[test]
fn csv_import_supports_timezone_offset_policy() {
    let csv = "ts\n1970-01-01T00:00:00-01:00\n";
    let options = CsvOptions {
        timestamp_tz_policy: CsvTimestampTzPolicy::ConvertToUtc,
        ..CsvOptions::default()
    };
    let sheet = import_csv_to_worksheet(1, "Data", Cursor::new(csv.as_bytes()), options).unwrap();
    assert_eq!(
        sheet.value(CellRef::new(0, 0)),
        CellValue::Number(3_600_000.0)
    );
}

#[test]
fn csv_import_supports_additional_date_formats() {
    let csv = "d1,d2\n1970/01/02,19700102\n";
    let sheet = import_csv_to_worksheet(
        1,
        "Data",
        Cursor::new(csv.as_bytes()),
        CsvOptions::default(),
    )
    .unwrap();

    assert_eq!(
        sheet.value(CellRef::new(0, 0)),
        CellValue::Number(86_400_000.0)
    );
    assert_eq!(
        sheet.value(CellRef::new(0, 1)),
        CellValue::Number(86_400_000.0)
    );
}

#[test]
fn csv_import_parses_parentheses_negative_numbers_with_grouping() {
    let csv = "n\n\"(1,234.50)\"\n";
    let sheet = import_csv_to_worksheet(
        1,
        "Data",
        Cursor::new(csv.as_bytes()),
        CsvOptions::default(),
    )
    .unwrap();

    assert_eq!(sheet.value(CellRef::new(0, 0)), CellValue::Number(-1234.5));
}

#[test]
fn csv_import_reports_invalid_utf8_with_row_and_column() {
    let mut bytes = b"id,text\n1,\"hello".to_vec();
    bytes.push(0xFF);
    bytes.extend_from_slice(b"\"\n");

    let sheet = import_csv_to_worksheet(1, "Data", Cursor::new(bytes), CsvOptions::default())
        .expect("CSV import should fall back to Windows-1252 in Auto mode");
    assert_eq!(
        sheet.value(CellRef::new(0, 1)),
        CellValue::String("helloÿ".to_string())
    );
}

#[test]
fn csv_import_utf8_encoding_reports_invalid_utf8_with_row_and_column() {
    let mut bytes = b"id,text\n1,\"hello".to_vec();
    bytes.push(0xFF);
    bytes.extend_from_slice(b"\"\n");

    let err = import_csv_to_columnar_table(
        Cursor::new(bytes),
        CsvOptions {
            encoding: CsvTextEncoding::Utf8,
            ..CsvOptions::default()
        },
    )
    .unwrap_err();

    match err {
        formula_model::import::CsvImportError::Parse {
            row,
            column,
            reason,
        } => {
            assert_eq!(row, 2);
            assert_eq!(column, 2);
            assert!(reason.contains("UTF-8"));
        }
        other => panic!("expected CsvImportError::Parse, got {other:?}"),
    }
}

#[test]
fn csv_import_windows1252_encoding_decodes_invalid_utf8_bytes() {
    let mut bytes = b"id,text\n1,\"hello".to_vec();
    bytes.push(0xFF);
    bytes.extend_from_slice(b"\"\n");

    let sheet = import_csv_to_worksheet(
        1,
        "Data",
        Cursor::new(bytes),
        CsvOptions {
            encoding: CsvTextEncoding::Windows1252,
            ..CsvOptions::default()
        },
    )
    .expect("CSV import should decode bytes as Windows-1252");

    assert_eq!(
        sheet.value(CellRef::new(0, 1)),
        CellValue::String("helloÿ".to_string())
    );
}

#[test]
fn csv_import_decodes_windows1252_fallback() {
    // "café" with Windows-1252 byte 0xE9 for "é" (invalid UTF-8).
    let bytes = b"id,text\n1,caf\xe9\n".to_vec();
    let sheet = import_csv_to_worksheet(1, "Data", Cursor::new(bytes), CsvOptions::default())
        .expect("CSV import should decode Windows-1252 bytes in Auto mode");
    assert_eq!(
        sheet.value(CellRef::new(0, 1)),
        CellValue::String("café".to_string())
    );
}

#[test]
fn csv_import_decodes_windows1252_euro_symbol_and_parses_currency() {
    // Windows-1252 byte 0x80 is "€". Ensure we decode it correctly (not ISO-8859-1 control chars)
    // and that currency parsing works on the decoded string.
    let bytes = b"amount\n\x8012.34\n".to_vec();
    let sheet = import_csv_to_worksheet(1, "Data", Cursor::new(bytes), CsvOptions::default())
        .expect("CSV import should decode Windows-1252 bytes in Auto mode");
    assert_eq!(sheet.value(CellRef::new(0, 0)), CellValue::Number(12.34));
}

#[test]
fn csv_import_decodes_windows1252_header_names() {
    // Header "café" with Windows-1252 byte 0xE9 for "é" (invalid UTF-8).
    let bytes = b"caf\xe9\n1\n".to_vec();
    let table = import_csv_to_columnar_table(Cursor::new(bytes), CsvOptions::default()).unwrap();
    assert_eq!(table.schema()[0].name, "café");
}

#[test]
fn csv_import_strips_utf8_bom_from_first_header_field() {
    let bytes = b"\xEF\xBB\xBFid,text\n1,hello\n".to_vec();
    let table = import_csv_to_columnar_table(Cursor::new(bytes), CsvOptions::default()).unwrap();
    assert_eq!(table.schema()[0].name, "id");
}

#[test]
fn csv_import_strips_utf8_bom_when_first_header_field_is_quoted() {
    // If the BOM appears before the opening quote, the CSV parser should still treat the field as
    // quoted after the BOM is stripped.
    let bytes = b"\xEF\xBB\xBF\"id\",text\n1,hello\n".to_vec();
    let table = import_csv_to_columnar_table(Cursor::new(bytes), CsvOptions::default()).unwrap();
    assert_eq!(table.schema()[0].name, "id");
}

#[test]
fn csv_import_strips_utf8_bom_from_first_data_field_when_no_header() {
    let bytes = b"\xEF\xBB\xBFhello,world\nfoo,bar\n".to_vec();
    let sheet = import_csv_to_worksheet(
        1,
        "Data",
        Cursor::new(bytes),
        CsvOptions {
            has_header: false,
            ..CsvOptions::default()
        },
    )
    .unwrap();

    assert_eq!(
        sheet.value(CellRef::new(0, 0)),
        CellValue::String("hello".to_string())
    );
    assert_eq!(
        sheet.value(CellRef::new(0, 1)),
        CellValue::String("world".to_string())
    );
}
