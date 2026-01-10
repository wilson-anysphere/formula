use formula_engine::date::{serial_to_ymd, ymd_to_serial, ExcelDate, ExcelDateSystem};

#[test]
fn excel_1900_date_system_emulates_lotus_bug() {
    let system = ExcelDateSystem::EXCEL_1900;

    assert_eq!(
        ymd_to_serial(ExcelDate::new(1900, 1, 1), system).unwrap(),
        1
    );
    assert_eq!(
        ymd_to_serial(ExcelDate::new(1900, 2, 28), system).unwrap(),
        59
    );
    assert_eq!(
        ymd_to_serial(ExcelDate::new(1900, 2, 29), system).unwrap(),
        60
    );
    assert_eq!(
        ymd_to_serial(ExcelDate::new(1900, 3, 1), system).unwrap(),
        61
    );

    assert_eq!(
        serial_to_ymd(59, system).unwrap(),
        ExcelDate::new(1900, 2, 28)
    );
    assert_eq!(
        serial_to_ymd(60, system).unwrap(),
        ExcelDate::new(1900, 2, 29)
    );
    assert_eq!(
        serial_to_ymd(61, system).unwrap(),
        ExcelDate::new(1900, 3, 1)
    );
}

#[test]
fn excel_1904_date_system_has_different_epoch() {
    let system = ExcelDateSystem::Excel1904;
    assert_eq!(
        ymd_to_serial(ExcelDate::new(1904, 1, 1), system).unwrap(),
        0
    );
    assert_eq!(
        serial_to_ymd(0, system).unwrap(),
        ExcelDate::new(1904, 1, 1)
    );
}

#[test]
fn lotus_bug_can_be_disabled() {
    let system = ExcelDateSystem::Excel1900 {
        lotus_compat: false,
    };

    // Without the Lotus bug, 1900-03-01 is serial 60 (no fictitious Feb 29).
    assert_eq!(
        ymd_to_serial(ExcelDate::new(1900, 3, 1), system).unwrap(),
        60
    );
    assert_eq!(
        serial_to_ymd(60, system).unwrap(),
        ExcelDate::new(1900, 3, 1)
    );

    // 1900-02-29 is not a real Gregorian date and should be rejected.
    assert!(ymd_to_serial(ExcelDate::new(1900, 2, 29), system).is_err());
}
