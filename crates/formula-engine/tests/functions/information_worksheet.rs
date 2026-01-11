use formula_engine::eval::parse_a1;
use formula_engine::{Engine, ErrorKind, Value};

use super::harness::{assert_number, TestSheet};

#[test]
fn isblank_distinguishes_blank_cell_from_empty_string() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", Value::Blank);
    sheet.set("A2", "");

    assert_eq!(sheet.eval("=ISBLANK(A1)"), Value::Bool(true));
    assert_eq!(sheet.eval("=ISBLANK(A2)"), Value::Bool(false));
    assert_eq!(sheet.eval("=ISBLANK(\"\")"), Value::Bool(false));
}

#[test]
fn isnumber_istext_islogical_work_on_scalars_and_references() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1.0);
    sheet.set("A2", "x");
    sheet.set("A3", true);

    assert_eq!(sheet.eval("=ISNUMBER(A1)"), Value::Bool(true));
    assert_eq!(sheet.eval("=ISNUMBER(A2)"), Value::Bool(false));

    assert_eq!(sheet.eval("=ISTEXT(A2)"), Value::Bool(true));
    assert_eq!(sheet.eval("=ISTEXT(A1)"), Value::Bool(false));

    assert_eq!(sheet.eval("=ISLOGICAL(A3)"), Value::Bool(true));
    assert_eq!(sheet.eval("=ISLOGICAL(A2)"), Value::Bool(false));
}

#[test]
fn isna_iserr_iserror_distinguish_error_kinds() {
    let mut sheet = TestSheet::new();

    assert_eq!(sheet.eval("=ISNA(#N/A)"), Value::Bool(true));
    assert_eq!(sheet.eval("=ISNA(#DIV/0!)"), Value::Bool(false));

    assert_eq!(sheet.eval("=ISERR(#DIV/0!)"), Value::Bool(true));
    assert_eq!(sheet.eval("=ISERR(#N/A)"), Value::Bool(false));

    assert_eq!(sheet.eval("=ISERROR(#DIV/0!)"), Value::Bool(true));
    assert_eq!(sheet.eval("=ISERROR(#N/A)"), Value::Bool(true));
    assert_eq!(sheet.eval("=ISERROR(1)"), Value::Bool(false));

    // Error-checking functions should not propagate the error value.
    assert_eq!(sheet.eval("=ISERROR(1/0)"), Value::Bool(true));
}

#[test]
fn type_returns_excel_type_codes() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", Value::Blank);

    assert_number(&sheet.eval("=TYPE(A1)"), 1.0);
    assert_number(&sheet.eval("=TYPE(1)"), 1.0);
    assert_number(&sheet.eval("=TYPE(\"x\")"), 2.0);
    assert_number(&sheet.eval("=TYPE(TRUE)"), 4.0);
    assert_number(&sheet.eval("=TYPE(#DIV/0!)"), 16.0);
    assert_number(&sheet.eval("=TYPE({1,2})"), 64.0);
}

#[test]
fn error_type_matches_excel_error_kind_codes() {
    let mut sheet = TestSheet::new();

    assert_number(&sheet.eval("=ERROR.TYPE(#NULL!)"), 1.0);
    assert_number(&sheet.eval("=ERROR.TYPE(#DIV/0!)"), 2.0);
    assert_number(&sheet.eval("=ERROR.TYPE(#VALUE!)"), 3.0);
    assert_number(&sheet.eval("=ERROR.TYPE(#REF!)"), 4.0);
    assert_number(&sheet.eval("=ERROR.TYPE(#NAME?)"), 5.0);
    assert_number(&sheet.eval("=ERROR.TYPE(#NUM!)"), 6.0);
    assert_number(&sheet.eval("=ERROR.TYPE(#N/A)"), 7.0);
    assert_number(&sheet.eval("=ERROR.TYPE(#SPILL!)"), 9.0);
    assert_number(&sheet.eval("=ERROR.TYPE(#CALC!)"), 10.0);

    assert_eq!(sheet.eval("=ERROR.TYPE(1)"), Value::Error(ErrorKind::NA));
}

#[test]
fn isnumber_spills_elementwise_over_array_literals() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "C1", "=ISNUMBER({1;\"x\"})")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "C1").expect("spill range");
    assert_eq!(start, parse_a1("C1").unwrap());
    assert_eq!(end, parse_a1("C2").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Bool(true));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Bool(false));
}

#[test]
fn n_and_t_match_excel_coercions() {
    let mut sheet = TestSheet::new();

    assert_number(&sheet.eval("=N(5)"), 5.0);
    assert_number(&sheet.eval("=N(TRUE)"), 1.0);
    assert_number(&sheet.eval("=N(FALSE)"), 0.0);
    assert_number(&sheet.eval("=N(\"hello\")"), 0.0);
    assert_eq!(sheet.eval("=N(#DIV/0!)"), Value::Error(ErrorKind::Div0));

    assert_eq!(sheet.eval("=T(\"hello\")"), Value::Text("hello".to_string()));
    assert_eq!(sheet.eval("=T(5)"), Value::Text(String::new()));
    assert_eq!(sheet.eval("=T(TRUE)"), Value::Text(String::new()));
    assert_eq!(sheet.eval("=T(#DIV/0!)"), Value::Error(ErrorKind::Div0));
}

