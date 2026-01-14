use formula_engine::{parse_formula, LocaleConfig, ParseOptions, SerializeOptions};

fn roundtrip(formula: &str, opts: ParseOptions, ser: SerializeOptions) {
    let ast1 = parse_formula(formula, opts.clone()).unwrap();
    let s = ast1.to_string(ser).unwrap();
    let ast2 = parse_formula(&s, opts).unwrap();
    assert_eq!(ast1, ast2, "formula `{formula}` -> `{s}`");
}

#[test]
fn roundtrip_with_quoted_sheet_and_structured_refs() {
    let opts = ParseOptions::default();
    let ser = SerializeOptions::default();
    roundtrip("='My Sheet'!$A$1+Table1[Column]+[@Column]", opts, ser);
}

#[test]
fn roundtrip_with_nested_structured_refs() {
    let opts = ParseOptions::default();
    let ser = SerializeOptions::default();
    roundtrip("=Table1[[#Headers],[Column]]+1", opts, ser);
}

#[test]
fn roundtrip_with_external_ref_and_array_literal() {
    let opts = ParseOptions::default();
    let ser = SerializeOptions::default();
    roundtrip("=[Book.xlsx]Sheet1!A1+SUM({1,2;3,4})", opts, ser);
}

#[test]
fn roundtrip_with_quoted_external_ref() {
    let opts = ParseOptions::default();
    let ser = SerializeOptions::default();
    roundtrip("='[Book Name.xlsx]Sheet 1'!A1+1", opts, ser);
}

#[test]
fn roundtrip_with_path_qualified_external_ref_in_quoted_sheet_prefix() {
    let opts = ParseOptions::default();
    let ser = SerializeOptions::default();
    roundtrip("='C:\\path\\[Book.xlsx]Sheet1'!A1+1", opts, ser);
}

#[test]
fn roundtrip_with_path_qualified_external_ref_with_brackets_in_path_prefix() {
    let opts = ParseOptions::default();
    let ser = SerializeOptions::default();
    roundtrip("='C:\\[foo]\\[Book.xlsx]Sheet1'!A1+1", opts, ser);
}

#[test]
fn roundtrip_with_unquoted_external_ref_with_spaces_in_workbook_name() {
    let opts = ParseOptions::default();
    let ser = SerializeOptions::default();
    roundtrip("=[Work Book-1.xlsx]Sheet1!A1+1", opts, ser);
}

#[test]
fn roundtrip_with_unquoted_external_ref_with_apostrophe_in_workbook_name() {
    let opts = ParseOptions::default();
    let ser = SerializeOptions::default();
    roundtrip("=[Bob's.xlsx]Sheet1!A1+1", opts, ser);
}

#[test]
fn roundtrip_with_external_workbook_scoped_name_ref() {
    let opts = ParseOptions::default();
    let ser = SerializeOptions::default();

    // Workbook-scoped external defined names are emitted as a single quoted token so they remain
    // parseable (`[Book]Name` is ambiguous with structured references for our lexer/parser).
    roundtrip("='[Book2.xlsb]MyName'", opts.clone(), ser.clone());
    roundtrip("='[AddIn]MyAddinConst'", opts, ser);
}

#[test]
fn roundtrip_preserves_sheet_quoting_for_names_that_cannot_be_unquoted() {
    let opts = ParseOptions::default();
    let ser = SerializeOptions::default();
    // Sheet names that look like tokens (cell refs / booleans) or include non-identifier
    // characters must remain quoted to round-trip.
    roundtrip("='A1'!B2+1", opts.clone(), ser.clone());
    roundtrip("='My-Sheet'!A1+1", opts.clone(), ser.clone());
    roundtrip("='TRUE'!A1+1", opts, ser);
}

#[test]
fn roundtrip_with_xlfn_prefix_in_file_mode() {
    let opts = ParseOptions::default();
    let mut ser = SerializeOptions::default();
    ser.include_xlfn_prefix = true;
    roundtrip("=_xlfn.XLOOKUP(A1,B1,C1)", opts, ser);
}

#[test]
fn roundtrip_de_de_locale() {
    let mut opts = ParseOptions::default();
    opts.locale = LocaleConfig::de_de();
    let mut ser = SerializeOptions::default();
    ser.locale = LocaleConfig::de_de();
    roundtrip("=SUM(1,23;{1\\2;3\\4})", opts, ser);
}

#[test]
fn roundtrip_union_inside_function_arg_is_parenthesized() {
    let opts = ParseOptions::default();
    let ser = SerializeOptions::default();
    // Union operator inside function arguments requires parentheses to avoid being parsed
    // as multiple arguments.
    roundtrip("=SUM((A1,B1)+C1)", opts, ser);
}

#[test]
fn roundtrip_call_expression() {
    let opts = ParseOptions::default();
    let ser = SerializeOptions::default();
    roundtrip("=LAMBDA(x,x+1)(5)", opts, ser);
}

#[test]
fn serializes_sheet_names_that_require_quoting() {
    let opts = ParseOptions::default();
    let ser = SerializeOptions::default();

    // Excel requires quoting sheet names that look like cell references.
    let ast = parse_formula("='A1'!B2", opts.clone()).unwrap();
    assert_eq!(ast.to_string(ser.clone()).unwrap(), "='A1'!B2");

    let ast = parse_formula("='R1C1'!A1", opts.clone()).unwrap();
    assert_eq!(ast.to_string(ser.clone()).unwrap(), "='R1C1'!A1");

    // Reserved boolean keywords must be quoted to avoid parsing as literals.
    let ast = parse_formula("='TRUE'!A1", opts.clone()).unwrap();
    assert_eq!(ast.to_string(ser.clone()).unwrap(), "='TRUE'!A1");

    // Sheet names starting with digits must be quoted to avoid parsing as row references / numbers.
    let ast = parse_formula("='2019'!A1", opts.clone()).unwrap();
    assert_eq!(ast.to_string(ser.clone()).unwrap(), "='2019'!A1");

    // Non-ASCII sheet names must be quoted to remain parseable by the canonical lexer.
    let ast = parse_formula("='Résumé'!A1", opts).unwrap();
    assert_eq!(ast.to_string(ser).unwrap(), "='Résumé'!A1");
}

#[test]
fn serializes_external_workbook_prefixes_without_quotes_when_possible() {
    let opts = ParseOptions::default();
    let ser = SerializeOptions::default();
    let ast = parse_formula("=[Book.xlsx]Sheet1!A1+1", opts).unwrap();
    assert_eq!(ast.to_string(ser).unwrap(), "=[Book.xlsx]Sheet1!A1+1");
}

#[test]
fn serializes_external_workbook_prefixes_with_quoted_sheet_names_when_required() {
    let opts = ParseOptions::default();
    let ser = SerializeOptions::default();

    // Sheet names with spaces must be quoted (even when the workbook is unquoted).
    let ast = parse_formula("=[Book.xlsx]'My Sheet'!A1+1", opts).unwrap();
    assert_eq!(ast.to_string(ser).unwrap(), "='[Book.xlsx]My Sheet'!A1+1");
}

#[test]
fn serializes_degenerate_external_3d_sheet_spans_as_single_sheet() {
    let opts = ParseOptions::default();
    let ser = SerializeOptions::default();
    let ast = parse_formula("=[Book.xlsx]Sheet1:Sheet1!A1", opts).unwrap();
    assert_eq!(ast.to_string(ser).unwrap(), "=[Book.xlsx]Sheet1!A1");
}

#[test]
fn roundtrip_with_external_3d_sheet_span() {
    let opts = ParseOptions::default();
    let ser = SerializeOptions::default();
    roundtrip("=SUM([Book.xlsx]Sheet1:Sheet3!A1)", opts, ser);
}

#[test]
fn roundtrip_with_external_3d_sheet_span_with_quoted_sheet_names() {
    let opts = ParseOptions::default();
    let ser = SerializeOptions::default();
    roundtrip("=SUM([Book.xlsx]'Sheet 1':'Sheet 3'!A1)", opts, ser);
}
