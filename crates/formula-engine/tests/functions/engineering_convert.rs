use formula_engine::{ErrorKind, Value};

use super::harness::{assert_number, TestSheet};

#[test]
fn convert_length_meters_and_feet() {
    let mut sheet = TestSheet::new();
    assert_number(
        &sheet.eval(r#"=CONVERT(1,"m","ft")"#),
        3.280_839_895_013_123,
    );
    assert_number(&sheet.eval(r#"=CONVERT(1,"ft","m")"#), 0.3048);
}

#[test]
fn convert_mass_kilograms_and_pounds() {
    let mut sheet = TestSheet::new();
    assert_number(
        &sheet.eval(r#"=CONVERT(1,"kg","lbm")"#),
        2.204_622_621_848_7757,
    );
    assert_number(&sheet.eval(r#"=CONVERT(1,"lbm","kg")"#), 0.453_592_37);
}

#[test]
fn convert_length_inches_and_centimeters() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval(r#"=CONVERT(1,"in","cm")"#), 2.54);
    assert_number(&sheet.eval(r#"=CONVERT(2.54,"cm","in")"#), 1.0);
}

#[test]
fn convert_volume_liters_and_gallons() {
    let mut sheet = TestSheet::new();
    assert_number(
        &sheet.eval(r#"=CONVERT(1,"L","gal")"#),
        0.264_172_052_358_148_4,
    );
}

#[test]
fn convert_temperature_celsius_to_fahrenheit() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval(r#"=CONVERT(0,"C","F")"#), 32.0);
}

#[test]
fn convert_invalid_units_return_na() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval(r#"=CONVERT(1,"no_such_unit","m")"#),
        Value::Error(ErrorKind::NA)
    );
    assert_eq!(
        sheet.eval(r#"=CONVERT(1,"m","kg")"#),
        Value::Error(ErrorKind::NA)
    );
}
