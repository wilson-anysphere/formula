use formula_model::import::{import_csv_to_worksheet, CsvOptions};
use formula_model::{CellRef, CellValue};
use std::io::Cursor;

#[test]
fn csv_import_streams_into_columnar_backed_worksheet() {
    let csv = concat!(
        "id,amount,ratio,flag,ts,category\n",
        "1,$12.34,50%,true,1970-01-02,A\n",
        "2,$0.01,12.5%,false,1970-01-03,B\n",
    );

    let sheet = import_csv_to_worksheet(1, "Data", Cursor::new(csv.as_bytes()), CsvOptions::default())
        .unwrap();

    assert_eq!(sheet.value(CellRef::new(0, 0)), CellValue::Number(1.0));
    assert_eq!(sheet.value(CellRef::new(0, 1)), CellValue::Number(12.34));
    assert_eq!(sheet.value(CellRef::new(0, 2)), CellValue::Number(0.5));
    assert_eq!(sheet.value(CellRef::new(0, 3)), CellValue::Boolean(true));
    assert_eq!(sheet.value(CellRef::new(0, 4)), CellValue::Number(86_400_000.0));
    assert_eq!(
        sheet.value(CellRef::new(0, 5)),
        CellValue::String("A".to_string())
    );

    assert_eq!(sheet.value(CellRef::new(1, 0)), CellValue::Number(2.0));
    assert_eq!(sheet.value(CellRef::new(1, 1)), CellValue::Number(0.01));
    assert_eq!(sheet.value(CellRef::new(1, 2)), CellValue::Number(0.125));
    assert_eq!(sheet.value(CellRef::new(1, 3)), CellValue::Boolean(false));
    assert_eq!(sheet.value(CellRef::new(1, 4)), CellValue::Number(172_800_000.0));
    assert_eq!(
        sheet.value(CellRef::new(1, 5)),
        CellValue::String("B".to_string())
    );
}

