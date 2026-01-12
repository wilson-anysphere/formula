use formula_engine::{ErrorKind, Value};

use super::harness::{assert_number, TestSheet};

#[test]
fn sum_ignores_text_in_ranges_but_coerces_scalar_text() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", Value::Text("5".to_string()));
    sheet.set("A2", 3.0);
    sheet.set("A3", 4.0);

    assert_number(&sheet.eval("=SUM(A1:A3)"), 7.0);
    assert_number(&sheet.eval(r#"=SUM("5", TRUE, 3)"#), 9.0);
}

#[test]
fn average_ignores_text_in_ranges_but_coerces_scalar_text() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", Value::Text("5".to_string()));
    sheet.set("A2", 3.0);
    sheet.set("A3", 5.0);

    // Text in references is ignored by AVERAGE.
    assert_number(&sheet.eval("=AVERAGE(A1:A3)"), 4.0);
    // Scalar text/logicals are coerced by AVERAGE.
    assert_number(&sheet.eval(r#"=AVERAGE("5", TRUE, 3)"#), 3.0);
}

#[test]
fn sum_propagates_errors() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1.0);
    sheet.set("A2", Value::Error(ErrorKind::Div0));
    assert_eq!(sheet.eval("=SUM(A1:A2)"), Value::Error(ErrorKind::Div0));
}

#[test]
fn simd_aggregate_fast_paths_match_scalar_semantics_for_large_arrays() {
    let mut sheet = TestSheet::new();

    // Build a 1x512 array literal with mixed types. Only numeric values should be considered by
    // SUM/AVERAGE/MIN/MAX/COUNT when the argument is an array (Excel reference semantics).
    //
    // Pattern per 4 cells: [number, TRUE, "x", blank]
    let mut elems = Vec::with_capacity(512);
    for i in 1..=512 {
        let entry = match i % 4 {
            0 => i.to_string(),
            1 => "TRUE".to_string(),
            2 => "\"x\"".to_string(),
            _ => String::new(), // blank cell in array literal.
        };
        elems.push(entry);
    }
    let array_literal = format!("{{{}}}", elems.join(","));

    // Numbers are 4, 8, ..., 512 (128 values).
    let expected_sum = 4.0 * (128.0 * 129.0 / 2.0);
    let expected_count = 128.0;
    let expected_avg = expected_sum / expected_count;
    let expected_min = 4.0;
    let expected_max = 512.0;

    assert_number(
        &sheet.eval(&format!("=LET(x,{array_literal},SUM(x))")),
        expected_sum,
    );
    assert_number(
        &sheet.eval(&format!("=LET(x,{array_literal},AVERAGE(x))")),
        expected_avg,
    );
    assert_number(
        &sheet.eval(&format!("=LET(x,{array_literal},MIN(x))")),
        expected_min,
    );
    assert_number(
        &sheet.eval(&format!("=LET(x,{array_literal},MAX(x))")),
        expected_max,
    );
    assert_number(
        &sheet.eval(&format!("=LET(x,{array_literal},COUNT(x))")),
        expected_count,
    );

    // COUNTIF with numeric criteria coerces bool/blank/text exactly like the scalar criteria
    // matcher. Here, TRUE counts as 1 (matches >0) while "x" does not.
    assert_number(
        &sheet.eval(&format!("=LET(x,{array_literal},COUNTIF(x,\">0\"))")),
        256.0,
    );
}

#[test]
fn simd_array_aggregates_propagate_errors() {
    let mut sheet = TestSheet::new();

    let mut elems = Vec::new();
    for i in 1..=64 {
        if i == 40 {
            elems.push("#DIV/0!".to_string());
        } else {
            elems.push("1".to_string());
        }
    }
    let array_literal = format!("{{{}}}", elems.join(","));

    assert_eq!(
        sheet.eval(&format!("=LET(x,{array_literal},SUM(x))")),
        Value::Error(ErrorKind::Div0)
    );
    assert_eq!(
        sheet.eval(&format!("=LET(x,{array_literal},AVERAGE(x))")),
        Value::Error(ErrorKind::Div0)
    );
    assert_eq!(
        sheet.eval(&format!("=LET(x,{array_literal},MIN(x))")),
        Value::Error(ErrorKind::Div0)
    );
    assert_eq!(
        sheet.eval(&format!("=LET(x,{array_literal},MAX(x))")),
        Value::Error(ErrorKind::Div0)
    );
}

#[test]
fn simd_numeric_criteria_aggregates_over_large_arrays() {
    let mut sheet = TestSheet::new();

    // SEQUENCE is not currently bytecode-compiled, ensuring this hits the AST evaluator. These
    // sizes also exceed the SIMD threshold for array aggregates.
    assert_number(&sheet.eval("=COUNTIF(SEQUENCE(256),\">128\")"), 128.0);
    assert_number(
        &sheet.eval("=SUMIF(SEQUENCE(256),\">128\",SEQUENCE(256))"),
        24_640.0,
    );
    assert_number(&sheet.eval("=AVERAGEIF(SEQUENCE(256),\">128\")"), 192.5);
}

#[test]
fn aggregates_reject_lambda_values_inside_arrays() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval("=SUM({LAMBDA(x,x),1})"),
        Value::Error(ErrorKind::Value)
    );
    assert_eq!(
        sheet.eval("=AVERAGE({1,LAMBDA(x,x)})"),
        Value::Error(ErrorKind::Value)
    );
    assert_eq!(
        sheet.eval("=MIN({1,LAMBDA(x,x)})"),
        Value::Error(ErrorKind::Value)
    );
    assert_eq!(
        sheet.eval("=MAX({LAMBDA(x,x),1})"),
        Value::Error(ErrorKind::Value)
    );
}

#[test]
fn average_div0_when_no_numeric_values() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", Value::Text("x".to_string()));
    sheet.set("A2", Value::Blank);
    assert_eq!(sheet.eval("=AVERAGE(A1:A2)"), Value::Error(ErrorKind::Div0));
}

#[test]
fn min_max_ignore_text_in_ranges() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", Value::Text("100".to_string()));
    sheet.set("A2", 3.0);
    sheet.set("A3", 4.0);

    assert_number(&sheet.eval("=MIN(A1:A3)"), 3.0);
    assert_number(&sheet.eval("=MAX(A1:A3)"), 4.0);
    assert_number(&sheet.eval(r#"=MIN("5", TRUE, 3)"#), 1.0);
}

#[test]
fn count_counta_countblank() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1.0);
    sheet.set("A2", Value::Text("x".to_string()));
    sheet.set("A3", true);
    sheet.set("A4", Value::Blank);
    sheet.set("A5", Value::Text("".to_string()));
    sheet.set("A6", Value::Error(ErrorKind::Div0));

    assert_number(&sheet.eval("=COUNT(A1:A6)"), 1.0);
    assert_number(&sheet.eval("=COUNTA(A1:A6)"), 5.0);
    assert_number(&sheet.eval("=COUNTBLANK(A1:A6)"), 2.0);
}

#[test]
fn countif_treats_lambda_cells_like_errors() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval(r#"=COUNTIF({LAMBDA(x,x),1},"<>")"#), 1.0);
    assert_number(&sheet.eval(r##"=COUNTIF({LAMBDA(x,x),1},"#VALUE!")"##), 1.0);
    assert_number(&sheet.eval(r#"=SUMIF({LAMBDA(x,x),1},"<>",{10,20})"#), 20.0);
}

#[test]
fn countif_reference_union_dedupes_overlaps() {
    let mut sheet = TestSheet::new();
    sheet.set("A2", 2.0);
    sheet.set("A3", 3.0);

    // A2 overlaps between the two areas; it should only be counted once.
    assert_number(&sheet.eval(r#"=COUNTIF((A1:A2,A2:A3),">1")"#), 2.0);
}

#[test]
fn countif_reference_union_counts_missing_cells_as_blank() {
    let mut sheet = TestSheet::new();
    sheet.set("A2", Value::Text("".to_string()));

    // Union covers A1:A3, but only A2 is explicitly stored. Missing cells in the union
    // should behave as blanks and should not be double-counted across overlaps.
    assert_number(&sheet.eval(r#"=COUNTIF((A1:A2,A2:A3),"")"#), 3.0);
}

#[test]
fn countif_reference_union_counts_blanks_across_non_overlapping_areas() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1.0);
    sheet.set("C2", Value::Text("".to_string()));

    // (A1:A2,C1:C2) is 4 cells total; A2/C1 are missing (blank) and C2 is an explicit empty
    // string (also treated as blank).
    assert_number(&sheet.eval(r#"=COUNTIF((A1:A2,C1:C2),"")"#), 3.0);
}

#[test]
fn countblank_reference_union_dedupes_overlaps() {
    let mut sheet = TestSheet::new();
    sheet.set("A2", 1.0);

    // Union covers A1:A3 with an overlap at A2.
    assert_number(&sheet.eval("=COUNTBLANK((A1:A2,A2:A3))"), 2.0);
}

#[test]
fn countblank_reference_union_counts_blanks_across_non_overlapping_areas() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1.0);
    sheet.set("C2", 2.0);

    // (A1:A2,C1:C2) is 4 cells total, 2 non-blank => 2 blanks.
    assert_number(&sheet.eval("=COUNTBLANK((A1:A2,C1:C2))"), 2.0);
}

#[test]
fn round_variants() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=ROUND(2.5,0)"), 3.0);
    assert_number(&sheet.eval("=ROUND(-2.5,0)"), -3.0);
    assert_number(&sheet.eval("=ROUND(1234,-2)"), 1200.0);

    assert_number(&sheet.eval("=ROUNDDOWN(1.29,1)"), 1.2);
    assert_number(&sheet.eval("=ROUNDDOWN(-1.29,1)"), -1.2);
    assert_number(&sheet.eval("=ROUNDUP(1.21,1)"), 1.3);
    assert_number(&sheet.eval("=ROUNDUP(-1.21,1)"), -1.3);
}

#[test]
fn trunc_truncates_toward_zero() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=TRUNC(8.9)"), 8.0);
    assert_number(&sheet.eval("=TRUNC(-8.9)"), -8.0);
    assert_number(&sheet.eval("=TRUNC(1.29,1)"), 1.2);
    assert_number(&sheet.eval("=TRUNC(-1.29,1)"), -1.2);
    assert_number(&sheet.eval("=TRUNC(1234.567,-2)"), 1200.0);
    assert_number(&sheet.eval("=TRUNC(-1234.567,-2)"), -1200.0);

    sheet.set_formula("A1", "=TRUNC({1.9;2.1})");
    sheet.recalc();
    assert_number(&sheet.get("A1"), 1.0);
    assert_number(&sheet.get("A2"), 2.0);
}

#[test]
fn int_abs_mod() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=INT(2.9)"), 2.0);
    assert_number(&sheet.eval("=INT(-2.1)"), -3.0);

    assert_number(&sheet.eval("=ABS(-3)"), 3.0);

    assert_number(&sheet.eval("=MOD(5,2)"), 1.0);
    assert_number(&sheet.eval("=MOD(-3,2)"), 1.0);
    assert_number(&sheet.eval("=MOD(3,-2)"), -1.0);
    assert_eq!(sheet.eval("=MOD(5,0)"), Value::Error(ErrorKind::Div0));
}

#[test]
fn sign_returns_expected_signum() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=SIGN(-2)"), -1.0);
    assert_number(&sheet.eval("=SIGN(0)"), 0.0);
    assert_number(&sheet.eval("=SIGN(2)"), 1.0);
}

#[test]
fn sign_accepts_xlfn_prefix() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=_xlfn.SIGN(-2)"), -1.0);
}

#[test]
fn sumproduct_rejects_lambda_values() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval("=SUMPRODUCT(LAMBDA(x,x),1)"),
        Value::Error(ErrorKind::Value)
    );
}
