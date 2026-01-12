use formula_io::{open_workbook, save_workbook, Workbook};
use formula_model::CellValue;

#[test]
fn opens_csv_and_saves_as_xlsx() {
    let dir = tempfile::tempdir().expect("temp dir");
    let csv_path = dir.path().join("data.csv");
    std::fs::write(&csv_path, "col1,col2\n1,hello\n2,world\n").expect("write csv");

    let wb = open_workbook(&csv_path).expect("open csv workbook");
    match wb {
        Workbook::Model(_) => {}
        other => panic!("expected Workbook::Model, got {other:?}"),
    }

    let out_path = dir.path().join("out.xlsx");
    save_workbook(&wb, &out_path).expect("save xlsx");

    let file = std::fs::File::open(&out_path).expect("open output xlsx");
    let model = formula_io::xlsx::read_workbook_from_reader(file).expect("read output xlsx");
    let sheet = model.sheets.first().expect("sheet");

    assert_eq!(sheet.value_a1("A1").unwrap(), CellValue::Number(1.0));
    assert_eq!(
        sheet.value_a1("B1").unwrap(),
        CellValue::String("hello".to_string())
    );
    assert_eq!(sheet.value_a1("A2").unwrap(), CellValue::Number(2.0));
    assert_eq!(
        sheet.value_a1("B2").unwrap(),
        CellValue::String("world".to_string())
    );
}

