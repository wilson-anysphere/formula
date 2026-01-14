use formula_engine::{ErrorKind, Value};

use super::harness::{assert_number, TestSheet};

#[test]
fn fixed_width_base_conversions_match_expected_values() {
    let mut sheet = TestSheet::new();

    assert_number(&sheet.eval("=BIN2DEC(\"1010\")"), 10.0);
    assert_number(&sheet.eval("=BIN2DEC(\"1111111111\")"), -1.0);
    assert_number(&sheet.eval("=BIN2DEC(\"1000000000\")"), -512.0);
    assert_number(&sheet.eval("=BIN2DEC(\"0111111111\")"), 511.0);

    assert_eq!(sheet.eval("=DEC2BIN(10)"), Value::Text("1010".to_string()));
    assert_eq!(
        sheet.eval("=DEC2BIN(10,6)"),
        Value::Text("001010".to_string())
    );
    assert_eq!(
        sheet.eval("=DEC2BIN(-1)"),
        Value::Text("1111111111".to_string())
    );
    assert_eq!(
        sheet.eval("=DEC2BIN(-512)"),
        Value::Text("1000000000".to_string())
    );

    assert_eq!(sheet.eval("=OCT2DEC(\"17\")"), Value::Number(15.0));
    assert_eq!(sheet.eval("=OCT2DEC(\"7777777777\")"), Value::Number(-1.0));
    assert_eq!(sheet.eval("=HEX2DEC(\"FF\")"), Value::Number(255.0));
    assert_eq!(sheet.eval("=HEX2DEC(\"FFFFFFFFFF\")"), Value::Number(-1.0));

    assert_eq!(
        sheet.eval("=BIN2HEX(\"1111111111\")"),
        Value::Text("FFFFFFFFFF".to_string())
    );
    assert_eq!(
        sheet.eval("=HEX2BIN(\"FFFFFFFFFF\")"),
        Value::Text("1111111111".to_string())
    );
}

#[test]
fn fixed_width_base_conversions_validate_inputs() {
    let mut sheet = TestSheet::new();

    // Invalid digits.
    assert_eq!(
        sheet.eval("=BIN2DEC(\"102\")"),
        Value::Error(ErrorKind::Num)
    );
    assert_eq!(sheet.eval("=OCT2DEC(\"8\")"), Value::Error(ErrorKind::Num));
    assert_eq!(sheet.eval("=HEX2DEC(\"GG\")"), Value::Error(ErrorKind::Num));

    // Too many digits.
    assert_eq!(
        sheet.eval("=BIN2DEC(\"11111111111\")"),
        Value::Error(ErrorKind::Num)
    );

    // Range overflow.
    assert_eq!(sheet.eval("=DEC2BIN(512)"), Value::Error(ErrorKind::Num));
    assert_eq!(
        sheet.eval("=HEX2BIN(\"800\")"),
        Value::Error(ErrorKind::Num)
    );

    // Places validation.
    assert_eq!(sheet.eval("=DEC2BIN(10,2)"), Value::Error(ErrorKind::Num));
    assert_eq!(sheet.eval("=DEC2BIN(10,11)"), Value::Error(ErrorKind::Num));
}

#[test]
fn base_and_decimal_support_radix_and_min_length() {
    let mut sheet = TestSheet::new();

    assert_eq!(sheet.eval("=BASE(15,16)"), Value::Text("F".to_string()));
    assert_eq!(
        sheet.eval("=BASE(15,16,4)"),
        Value::Text("000F".to_string())
    );
    assert_eq!(sheet.eval("=DECIMAL(\"FF\",16)"), Value::Number(255.0));

    assert_eq!(sheet.eval("=BASE(1,1)"), Value::Error(ErrorKind::Num));
    assert_eq!(
        sheet.eval("=DECIMAL(\"2\",2)"),
        Value::Error(ErrorKind::Num)
    );
}

#[test]
fn bit_functions_match_expected_results_and_reject_invalid_inputs() {
    let mut sheet = TestSheet::new();

    assert_eq!(sheet.eval("=BITAND(5,3)"), Value::Number(1.0));
    assert_eq!(sheet.eval("=BITOR(5,3)"), Value::Number(7.0));
    assert_eq!(sheet.eval("=BITXOR(5,3)"), Value::Number(6.0));
    assert_eq!(sheet.eval("=BITLSHIFT(1,3)"), Value::Number(8.0));
    assert_eq!(sheet.eval("=BITRSHIFT(8,3)"), Value::Number(1.0));

    // Non-integer / negative / out-of-range inputs should fail.
    assert_eq!(sheet.eval("=BITAND(1.5,1)"), Value::Error(ErrorKind::Num));
    assert_eq!(sheet.eval("=BITOR(-1,1)"), Value::Error(ErrorKind::Num));
    assert_eq!(
        sheet.eval(&format!("=BITXOR({},1)", (1u64 << 48))),
        Value::Error(ErrorKind::Num)
    );
    assert_eq!(sheet.eval("=BITLSHIFT(1,48)"), Value::Error(ErrorKind::Num));
}
