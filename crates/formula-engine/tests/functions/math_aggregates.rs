use chrono::{TimeZone, Utc};

use formula_engine::coercion::ValueLocaleConfig;
use formula_engine::date::ExcelDateSystem;
use formula_engine::functions::math;
use formula_engine::value::NumberLocale;
use formula_engine::{ErrorKind, Value};

#[test]
fn sumif_supports_numeric_criteria() {
    let now_utc = Utc.with_ymd_and_hms(2024, 1, 1, 0, 0, 0).unwrap();
    let criteria_range = vec![1.into(), 2.into(), 3.into(), 4.into()];
    let sum_range = vec![10.into(), 20.into(), 30.into(), 40.into()];

    let criteria = Value::from(">2");
    assert_eq!(
        math::sumif(
            &criteria_range,
            &criteria,
            Some(&sum_range),
            ValueLocaleConfig::en_us(),
            now_utc,
            ExcelDateSystem::EXCEL_1900,
        )
        .unwrap(),
        70.0
    );

    let criteria = Value::Number(2.0);
    assert_eq!(
        math::sumif(
            &criteria_range,
            &criteria,
            Some(&sum_range),
            ValueLocaleConfig::en_us(),
            now_utc,
            ExcelDateSystem::EXCEL_1900,
        )
        .unwrap(),
        20.0
    );
}

#[test]
fn sumif_supports_wildcards_and_blanks() {
    let now_utc = Utc.with_ymd_and_hms(2024, 1, 1, 0, 0, 0).unwrap();
    let criteria_range = vec![
        Value::from("apple"),
        Value::from("banana"),
        Value::from("apricot"),
        Value::Blank,
        Value::from(""),
    ];
    let sum_range = vec![1.into(), 2.into(), 3.into(), 4.into(), 5.into()];

    let criteria = Value::from("ap*");
    assert_eq!(
        math::sumif(
            &criteria_range,
            &criteria,
            Some(&sum_range),
            ValueLocaleConfig::en_us(),
            now_utc,
            ExcelDateSystem::EXCEL_1900,
        )
        .unwrap(),
        4.0
    );

    let criteria = Value::from("");
    assert_eq!(
        math::sumif(
            &criteria_range,
            &criteria,
            Some(&sum_range),
            ValueLocaleConfig::en_us(),
            now_utc,
            ExcelDateSystem::EXCEL_1900,
        )
        .unwrap(),
        9.0
    );
}

#[test]
fn sumifs_requires_all_criteria_to_match() {
    let now_utc = Utc.with_ymd_and_hms(2024, 1, 1, 0, 0, 0).unwrap();
    let sum_range = vec![10.into(), 20.into(), 30.into(), 40.into()];
    let range1 = vec![
        Value::from("A"),
        Value::from("A"),
        Value::from("B"),
        Value::from("B"),
    ];
    let range2 = vec![1.into(), 2.into(), 3.into(), 4.into()];

    let crit1 = Value::from("A");
    let crit2 = Value::from(">1");
    let criteria_pairs = [(&range1[..], &crit1), (&range2[..], &crit2)];
    assert_eq!(
        math::sumifs(
            &sum_range,
            &criteria_pairs,
            ValueLocaleConfig::en_us(),
            now_utc,
            ExcelDateSystem::EXCEL_1900,
        )
        .unwrap(),
        20.0
    );
}

#[test]
fn sumifs_length_mismatch_is_value_error() {
    let now_utc = Utc.with_ymd_and_hms(2024, 1, 1, 0, 0, 0).unwrap();
    let sum_range = vec![1.into(), 2.into()];
    let range = vec![1.into()];
    let crit = Value::from("1");
    let criteria_pairs = [(&range[..], &crit)];
    assert_eq!(
        math::sumifs(
            &sum_range,
            &criteria_pairs,
            ValueLocaleConfig::en_us(),
            now_utc,
            ExcelDateSystem::EXCEL_1900,
        )
        .unwrap_err(),
        ErrorKind::Value
    );
}

#[test]
fn sumproduct_multiplies_and_sums() {
    let a = vec![1.into(), 2.into(), 3.into()];
    let b = vec![4.into(), 5.into(), 6.into()];
    assert_eq!(
        math::sumproduct(&[&a, &b], NumberLocale::en_us()).unwrap(),
        32.0
    );
}

#[test]
fn sumproduct_propagates_errors() {
    let a = vec![1.into(), Value::Error(ErrorKind::Div0)];
    let b = vec![2.into(), 3.into()];
    assert_eq!(
        math::sumproduct(&[&a, &b], NumberLocale::en_us()).unwrap_err(),
        ErrorKind::Div0
    );
}

#[test]
fn sumproduct_treats_non_numeric_text_as_zero() {
    let a = vec![Value::from("x"), 2.into()];
    let b = vec![3.into(), 4.into()];
    // "x" coerces to 0, so SUMPRODUCT is 0*3 + 2*4.
    assert_eq!(
        math::sumproduct(&[&a, &b], NumberLocale::en_us()).unwrap(),
        8.0
    );
}

#[test]
fn sumproduct_broadcasts_scalar_to_other_array_len() {
    let scalar = vec![2.into()];
    let values = vec![1.into(), 2.into(), 3.into()];

    // 2 is broadcast to [2,2,2].
    assert_eq!(
        math::sumproduct(&[&scalar, &values], NumberLocale::en_us()).unwrap(),
        12.0
    );
    // Broadcast should work regardless of whether the scalar is the first or second array.
    assert_eq!(
        math::sumproduct(&[&values, &scalar], NumberLocale::en_us()).unwrap(),
        12.0
    );
}

#[test]
fn subtotal_implements_common_function_nums() {
    let values = vec![1.into(), 2.into(), 3.into(), Value::from("x"), Value::Blank];
    assert_eq!(math::subtotal(9, &values).unwrap(), 6.0);
    assert_eq!(math::subtotal(1, &values).unwrap(), 2.0);
    assert_eq!(math::subtotal(2, &values).unwrap(), 3.0);
    assert_eq!(math::subtotal(3, &values).unwrap(), 4.0);
}

#[test]
fn aggregate_can_ignore_errors() {
    let values = vec![1.into(), Value::Error(ErrorKind::Div0), 2.into()];
    assert_eq!(math::aggregate(9, 2, &values).unwrap(), 3.0);
    assert_eq!(math::aggregate(9, 4, &values).unwrap_err(), ErrorKind::Div0);
}
