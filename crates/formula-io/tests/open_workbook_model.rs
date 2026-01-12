use std::io::Write as _;
use std::path::PathBuf;

use formula_model::{
    sanitize_sheet_name, CellRef, CellValue, DateSystem, SheetVisibility, EXCEL_MAX_SHEET_NAME_LEN,
};

fn fixture_path(rel: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../fixtures").join(rel)
}

fn xlsb_fixture_path(rel: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../formula-xlsb/tests/fixtures")
        .join(rel)
}

fn xls_fixture_path(rel: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../formula-xls/tests/fixtures")
        .join(rel)
}

fn xlsx_test_fixture_path(rel: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../formula-xlsx/tests/fixtures")
        .join(rel)
}

#[cfg(feature = "parquet")]
fn parquet_fixture_path(rel: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../packages/data-io/test/fixtures")
        .join(rel)
}

#[test]
fn open_workbook_model_xlsx() {
    let path = fixture_path("xlsx/basic/basic.xlsx");
    let workbook = formula_io::open_workbook_model(&path).expect("open workbook model");

    assert_eq!(workbook.sheets.len(), 1);
    assert_eq!(workbook.sheets[0].name, "Sheet1");

    let sheet = workbook.sheet_by_name("Sheet1").expect("Sheet1 missing");
    assert_eq!(
        sheet.value(CellRef::from_a1("A1").unwrap()),
        CellValue::Number(1.0)
    );
    assert_eq!(
        sheet.value(CellRef::from_a1("B1").unwrap()),
        CellValue::String("Hello".to_string())
    );
}

#[test]
fn open_workbook_model_xlsx_multi_sheet() {
    let path = fixture_path("xlsx/basic/multi-sheet.xlsx");
    let workbook = formula_io::open_workbook_model(&path).expect("open workbook model");

    assert_eq!(workbook.sheets.len(), 2);
    assert_eq!(workbook.sheets[0].name, "Sheet1");
    assert_eq!(workbook.sheets[1].name, "Sheet2");

    let sheet1 = workbook.sheet_by_name("Sheet1").expect("Sheet1 missing");
    assert_eq!(sheet1.value_a1("A1").unwrap(), CellValue::Number(1.0));
    assert_eq!(
        sheet1.value_a1("B1").unwrap(),
        CellValue::String("Hello".to_string())
    );

    let sheet2 = workbook.sheet_by_name("Sheet2").expect("Sheet2 missing");
    assert_eq!(sheet2.value_a1("A1").unwrap(), CellValue::Number(2.0));
}

#[test]
fn open_workbook_model_xlsx_shared_strings() {
    let path = fixture_path("xlsx/basic/shared-strings.xlsx");
    let workbook = formula_io::open_workbook_model(&path).expect("open workbook model");

    assert_eq!(workbook.sheets.len(), 1);
    assert_eq!(workbook.sheets[0].name, "Sheet1");

    let sheet = workbook.sheet_by_name("Sheet1").expect("Sheet1 missing");
    assert_eq!(
        sheet.value(CellRef::from_a1("A1").unwrap()),
        CellValue::String("Hello".to_string())
    );
    assert_eq!(
        sheet.value(CellRef::from_a1("B1").unwrap()),
        CellValue::String("World".to_string())
    );
}

#[test]
fn open_workbook_model_xlsx_date_system_1904() {
    let path = fixture_path("xlsx/metadata/date-system-1904.xlsx");
    let workbook = formula_io::open_workbook_model(&path).expect("open workbook model");
    assert_eq!(workbook.date_system, DateSystem::Excel1904);
}

#[test]
fn open_workbook_model_xlsx_sheet_visibility() {
    let path = xlsx_test_fixture_path("sheet-metadata.xlsx");
    let workbook = formula_io::open_workbook_model(&path).expect("open workbook model");

    assert_eq!(workbook.sheets.len(), 3);
    assert_eq!(workbook.sheets[0].name, "Visible");
    assert_eq!(workbook.sheets[0].visibility, SheetVisibility::Visible);

    assert_eq!(workbook.sheets[1].name, "Hidden");
    assert_eq!(workbook.sheets[1].visibility, SheetVisibility::Hidden);

    assert_eq!(workbook.sheets[2].name, "VeryHidden");
    assert_eq!(workbook.sheets[2].visibility, SheetVisibility::VeryHidden);
}

#[test]
fn open_workbook_model_xlsx_reads_formulas() {
    let path = fixture_path("xlsx/formulas/formulas.xlsx");
    let workbook = formula_io::open_workbook_model(&path).expect("open workbook model");

    assert_eq!(workbook.sheets.len(), 1);
    assert_eq!(workbook.sheets[0].name, "Sheet1");

    let sheet = workbook.sheet_by_name("Sheet1").expect("Sheet1 missing");
    assert_eq!(sheet.formula(CellRef::from_a1("C1").unwrap()), Some("A1+B1"));
}

#[test]
fn open_workbook_model_xlsm() {
    let path = fixture_path("xlsx/macros/basic.xlsm");
    let workbook = formula_io::open_workbook_model(&path).expect("open workbook model");

    assert_eq!(workbook.sheets.len(), 1);
    assert_eq!(workbook.sheets[0].name, "Sheet1");
}

#[test]
fn open_workbook_model_xltx() {
    let src = fixture_path("xlsx/basic/basic.xlsx");
    let tmp = tempfile::tempdir().expect("temp dir");
    let dst = tmp.path().join("basic.xltx");
    std::fs::copy(&src, &dst).expect("copy xlsx fixture to .xltx");

    let workbook = formula_io::open_workbook_model(&dst).expect("open workbook model");
    assert_eq!(workbook.sheets.len(), 1);
    assert_eq!(workbook.sheets[0].name, "Sheet1");

    let sheet = workbook.sheet_by_name("Sheet1").expect("Sheet1 missing");
    assert_eq!(sheet.value_a1("A1").unwrap(), CellValue::Number(1.0));
}

#[test]
fn open_workbook_model_xltm_and_xlam() {
    let src = fixture_path("xlsx/macros/basic.xlsm");
    let tmp = tempfile::tempdir().expect("temp dir");

    for ext in ["xltm", "xlam"] {
        let dst = tmp.path().join(format!("basic.{ext}"));
        std::fs::copy(&src, &dst).expect("copy xlsm fixture to template/add-in extension");

        let workbook = formula_io::open_workbook_model(&dst).expect("open workbook model");
        assert_eq!(workbook.sheets.len(), 1);
        assert_eq!(workbook.sheets[0].name, "Sheet1");
    }
}

#[test]
fn open_workbook_model_xlsx_ignores_chart_parts() {
    let path = fixture_path("charts/xlsx/basic-chart.xlsx");
    let workbook = formula_io::open_workbook_model(&path).expect("open workbook model");

    assert_eq!(workbook.sheets.len(), 1);
    assert_eq!(workbook.sheets[0].name, "Sheet1");
}

#[test]
fn open_workbook_model_sniffs_extensionless_xlsx() {
    let path = fixture_path("xlsx/basic/basic.xlsx");
    let bytes = std::fs::read(&path).expect("read fixture");

    let mut tmp = tempfile::Builder::new()
        .prefix("basic_xlsx_")
        .tempfile()
        .expect("tempfile");
    tmp.write_all(&bytes).expect("write tempfile");

    let workbook = formula_io::open_workbook_model(tmp.path()).expect("open workbook model");
    assert_eq!(workbook.sheets.len(), 1);
    assert_eq!(workbook.sheets[0].name, "Sheet1");
}

#[test]
fn open_workbook_model_sniffs_xlsx_with_wrong_extension() {
    let path = fixture_path("xlsx/basic/basic.xlsx");
    let bytes = std::fs::read(&path).expect("read fixture");

    let mut tmp = tempfile::Builder::new()
        .prefix("basic_xlsx_wrong_ext_")
        .suffix(".xls")
        .tempfile()
        .expect("tempfile");
    tmp.write_all(&bytes).expect("write tempfile");

    let workbook = formula_io::open_workbook_model(tmp.path()).expect("open workbook model");
    assert_eq!(workbook.sheets.len(), 1);
    assert_eq!(workbook.sheets[0].name, "Sheet1");
}

#[test]
fn open_workbook_model_xlsb() {
    let path = xlsb_fixture_path("simple.xlsb");
    let workbook = formula_io::open_workbook_model(&path).expect("open workbook model");

    assert_eq!(workbook.sheets.len(), 1);
    assert_eq!(workbook.sheets[0].name, "Sheet1");

    let sheet = workbook.sheet_by_name("Sheet1").expect("Sheet1 missing");
    assert_eq!(
        sheet.value(CellRef::from_a1("A1").unwrap()),
        CellValue::String("Hello".to_string())
    );
    assert_eq!(
        sheet.value(CellRef::from_a1("B1").unwrap()),
        CellValue::Number(42.5)
    );
    assert_eq!(sheet.formula(CellRef::from_a1("C1").unwrap()), Some("B1*2"));
}

#[test]
fn open_workbook_model_sniffs_extensionless_xlsb() {
    let path = xlsb_fixture_path("simple.xlsb");
    let bytes = std::fs::read(&path).expect("read fixture");

    let mut tmp = tempfile::Builder::new()
        .prefix("simple_xlsb_")
        .tempfile()
        .expect("tempfile");
    tmp.write_all(&bytes).expect("write tempfile");

    let workbook = formula_io::open_workbook_model(tmp.path()).expect("open workbook model");
    assert_eq!(workbook.sheets.len(), 1);
    assert_eq!(workbook.sheets[0].name, "Sheet1");
}

#[test]
fn open_workbook_model_sniffs_xlsb_with_wrong_extension() {
    let path = xlsb_fixture_path("simple.xlsb");
    let bytes = std::fs::read(&path).expect("read fixture");

    let mut tmp = tempfile::Builder::new()
        .prefix("simple_xlsb_wrong_ext_")
        .suffix(".xlsx")
        .tempfile()
        .expect("tempfile");
    tmp.write_all(&bytes).expect("write tempfile");

    let workbook = formula_io::open_workbook_model(tmp.path()).expect("open workbook model");
    assert_eq!(workbook.sheets.len(), 1);
    assert_eq!(workbook.sheets[0].name, "Sheet1");
}

#[test]
fn open_workbook_model_xls() {
    let path = xls_fixture_path("basic.xls");
    let workbook = formula_io::open_workbook_model(&path).expect("open workbook model");

    assert_eq!(workbook.sheets.len(), 2);
    assert_eq!(workbook.sheets[0].name, "Sheet1");
    assert_eq!(workbook.sheets[1].name, "Second");

    let sheet1 = workbook.sheet_by_name("Sheet1").expect("Sheet1 missing");
    assert_eq!(
        sheet1.value(CellRef::from_a1("A1").unwrap()),
        CellValue::String("Hello".to_string())
    );
    assert_eq!(
        sheet1.value(CellRef::from_a1("B2").unwrap()),
        CellValue::Number(123.0)
    );
    assert_eq!(
        sheet1.formula(CellRef::from_a1("C3").unwrap()),
        Some("B2*2")
    );

    let sheet2 = workbook.sheet_by_name("Second").expect("Second missing");
    assert_eq!(
        sheet2.value(CellRef::from_a1("A1").unwrap()),
        CellValue::String("Second sheet".to_string())
    );
}

#[test]
fn open_workbook_model_xlt_and_xla() {
    let src = xls_fixture_path("basic.xls");
    let tmp = tempfile::tempdir().expect("temp dir");

    for ext in ["xlt", "xla"] {
        let dst = tmp.path().join(format!("basic.{ext}"));
        std::fs::copy(&src, &dst).expect("copy xls fixture to legacy template/add-in extension");

        let workbook = formula_io::open_workbook_model(&dst).expect("open workbook model");
        assert_eq!(workbook.sheets.len(), 2);
        assert_eq!(workbook.sheets[0].name, "Sheet1");
        assert_eq!(workbook.sheets[1].name, "Second");
    }
}

#[test]
fn open_workbook_model_sniffs_extensionless_xls() {
    let path = xls_fixture_path("basic.xls");
    let bytes = std::fs::read(&path).expect("read fixture");

    let mut tmp = tempfile::Builder::new()
        .prefix("basic_xls_")
        .tempfile()
        .expect("tempfile");
    tmp.write_all(&bytes).expect("write tempfile");

    let workbook = formula_io::open_workbook_model(tmp.path()).expect("open workbook model");
    assert_eq!(workbook.sheets.len(), 2);
    assert_eq!(workbook.sheets[0].name, "Sheet1");
    assert_eq!(workbook.sheets[1].name, "Second");
}

#[test]
fn open_workbook_model_sniffs_xls_with_wrong_extension() {
    let path = xls_fixture_path("basic.xls");
    let bytes = std::fs::read(&path).expect("read fixture");

    let mut tmp = tempfile::Builder::new()
        .prefix("basic_xls_wrong_ext_")
        .suffix(".xlsx")
        .tempfile()
        .expect("tempfile");
    tmp.write_all(&bytes).expect("write tempfile");

    let workbook = formula_io::open_workbook_model(tmp.path()).expect("open workbook model");
    assert_eq!(workbook.sheets.len(), 2);
    assert_eq!(workbook.sheets[0].name, "Sheet1");
    assert_eq!(workbook.sheets[1].name, "Second");
}

#[test]
fn open_workbook_model_csv() {
    let dir = tempfile::tempdir().expect("temp dir");
    let csv_path = dir.path().join("data.csv");
    std::fs::write(&csv_path, "col1,col2\n1,hello\n2,world\n").expect("write csv");

    let workbook = formula_io::open_workbook_model(&csv_path).expect("open workbook model");
    assert_eq!(workbook.sheets.len(), 1);
    assert_eq!(workbook.sheets[0].name, "data");

    let sheet = workbook.sheet_by_name("data").expect("data sheet missing");
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

#[test]
fn open_workbook_model_sniffs_csv_with_wrong_extension() {
    let csv_bytes = b"col1,col2\n1,hello\n2,world\n";

    let mut tmp = tempfile::Builder::new()
        .prefix("data_wrong_ext_")
        .suffix(".xlsx")
        .tempfile()
        .expect("tempfile");
    tmp.write_all(csv_bytes).expect("write tempfile");

    let workbook = formula_io::open_workbook_model(tmp.path()).expect("open workbook model");
    assert_eq!(workbook.sheets.len(), 1);

    let sheet_name = workbook.sheets[0].name.clone();
    let sheet = workbook
        .sheet_by_name(&sheet_name)
        .expect("sheet missing");

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

#[test]
fn open_workbook_model_sniffs_csv_with_wrong_extension_and_sanitizes_sheet_name() {
    let dir = tempfile::tempdir().expect("temp dir");
    // Note: the extension is intentionally wrong; content sniffing should still treat it as CSV.
    let path = dir.path().join("bad[name].xlsx");

    std::fs::write(&path, "col1,col2\n1,hello\n2,world\n").expect("write csv");

    let workbook = formula_io::open_workbook_model(&path).expect("open workbook model");
    assert_eq!(workbook.sheets.len(), 1);
    assert_eq!(workbook.sheets[0].name, "badname");

    let sheet = workbook.sheet_by_name("badname").expect("sheet missing");
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

#[test]
fn open_workbook_model_sniffs_single_line_csv_with_wrong_extension() {
    // Single-row exports (or temp files) may omit a trailing newline. We should still classify the
    // content as CSV via sniffing even when the extension is wrong.
    let csv_bytes = b"a,b";

    let mut tmp = tempfile::Builder::new()
        .prefix("single_line_wrong_ext_")
        .suffix(".xlsx")
        .tempfile()
        .expect("tempfile");
    tmp.write_all(csv_bytes).expect("write tempfile");

    let workbook = formula_io::open_workbook_model(tmp.path()).expect("open workbook model");
    assert_eq!(workbook.sheets.len(), 1);

    let sheet_name = workbook.sheets[0].name.clone();
    let sheet = workbook
        .sheet_by_name(&sheet_name)
        .expect("sheet missing");
    let table = sheet.columnar_table().expect("expected columnar table");

    assert_eq!(table.column_count(), 2);
    assert_eq!(table.row_count(), 0);
    assert_eq!(table.schema()[0].name, "a");
    assert_eq!(table.schema()[1].name, "b");
}

#[test]
fn open_workbook_model_sniffs_extensionless_csv() {
    let csv_bytes = b"col1,col2\n1,hello\n2,world\n";

    let mut tmp = tempfile::Builder::new()
        .prefix("data_no_ext_")
        .tempfile()
        .expect("tempfile");
    tmp.write_all(csv_bytes).expect("write tempfile");

    let workbook = formula_io::open_workbook_model(tmp.path()).expect("open workbook model");
    assert_eq!(workbook.sheets.len(), 1);

    let sheet_name = workbook.sheets[0].name.clone();
    let sheet = workbook
        .sheet_by_name(&sheet_name)
        .expect("sheet missing");

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

#[test]
fn open_workbook_model_sniffs_csv_with_xls_extension() {
    let csv_bytes = b"col1,col2\n1,hello\n2,world\n";

    let mut tmp = tempfile::Builder::new()
        .prefix("data_wrong_ext_xls_")
        .suffix(".xls")
        .tempfile()
        .expect("tempfile");
    tmp.write_all(csv_bytes).expect("write tempfile");

    let workbook = formula_io::open_workbook_model(tmp.path()).expect("open workbook model");
    assert_eq!(workbook.sheets.len(), 1);

    let sheet_name = workbook.sheets[0].name.clone();
    let sheet = workbook
        .sheet_by_name(&sheet_name)
        .expect("sheet missing");

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

#[test]
fn open_workbook_model_sniffs_csv_with_xlsb_extension() {
    let csv_bytes = b"col1,col2\n1,hello\n2,world\n";

    let mut tmp = tempfile::Builder::new()
        .prefix("data_wrong_ext_xlsb_")
        .suffix(".xlsb")
        .tempfile()
        .expect("tempfile");
    tmp.write_all(csv_bytes).expect("write tempfile");

    let workbook = formula_io::open_workbook_model(tmp.path()).expect("open workbook model");
    assert_eq!(workbook.sheets.len(), 1);

    let sheet_name = workbook.sheets[0].name.clone();
    let sheet = workbook
        .sheet_by_name(&sheet_name)
        .expect("sheet missing");

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

#[test]
fn open_workbook_model_csv_decodes_windows1252() {
    let tmp = tempfile::tempdir().expect("temp dir");
    let path = tmp.path().join("data.csv");

    // "café" with Windows-1252 byte 0xE9 for "é" (invalid UTF-8).
    std::fs::write(&path, b"id,text\n1,caf\xe9\n").expect("write csv");

    let workbook = formula_io::open_workbook_model(&path).expect("open csv workbook model");
    assert_eq!(workbook.sheets.len(), 1);
    assert_eq!(workbook.sheets[0].name, "data");

    let sheet = workbook.sheet_by_name("data").expect("sheet missing");
    assert_eq!(
        sheet.value(CellRef::from_a1("B1").unwrap()),
        CellValue::String("café".to_string())
    );
}

#[test]
fn open_workbook_model_csv_strips_utf8_bom() {
    let tmp = tempfile::tempdir().expect("temp dir");
    let path = tmp.path().join("bom.csv");

    std::fs::write(&path, b"\xEF\xBB\xBFid,text\n1,hello\n").expect("write csv");

    let workbook = formula_io::open_workbook_model(&path).expect("open csv workbook model");
    let sheet = workbook.sheet_by_name("bom").expect("sheet missing");
    let table = sheet.columnar_table().expect("expected columnar table");
    assert_eq!(table.schema()[0].name, "id");
}

#[test]
fn open_workbook_model_sniffs_utf16le_tab_delimited_text() {
    let tmp = tempfile::tempdir().expect("temp dir");
    let path = tmp.path().join("data.txt");

    // Excel's "Unicode Text" export is UTF-16LE with a BOM and (typically) tab-delimited.
    let tsv = "col1\tcol2\r\n1\thello\r\n2\tworld\r\n";
    let mut bytes = vec![0xFF, 0xFE];
    for unit in tsv.encode_utf16() {
        bytes.extend_from_slice(&unit.to_le_bytes());
    }
    std::fs::write(&path, &bytes).expect("write utf16 tsv");

    let workbook = formula_io::open_workbook_model(&path).expect("open workbook model");
    assert_eq!(workbook.sheets.len(), 1);
    assert_eq!(workbook.sheets[0].name, "data");

    let sheet = workbook.sheet_by_name("data").expect("data sheet missing");
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

#[test]
fn open_workbook_model_sniffs_utf16be_tab_delimited_text() {
    let tmp = tempfile::tempdir().expect("temp dir");
    let path = tmp.path().join("data.txt");

    let tsv = "col1\tcol2\r\n1\thello\r\n2\tworld\r\n";
    let mut bytes = vec![0xFE, 0xFF];
    for unit in tsv.encode_utf16() {
        bytes.extend_from_slice(&unit.to_be_bytes());
    }
    std::fs::write(&path, &bytes).expect("write utf16be tsv");

    let workbook = formula_io::open_workbook_model(&path).expect("open workbook model");
    assert_eq!(workbook.sheets.len(), 1);
    assert_eq!(workbook.sheets[0].name, "data");

    let sheet = workbook.sheet_by_name("data").expect("data sheet missing");
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

#[test]
fn open_workbook_model_sniffs_utf16le_tab_delimited_text_without_bom() {
    let tmp = tempfile::tempdir().expect("temp dir");
    let path = tmp.path().join("data_no_bom.txt");

    let tsv = "col1\tcol2\r\n1\thello\r\n2\tworld\r\n";
    let mut bytes = Vec::new();
    for unit in tsv.encode_utf16() {
        bytes.extend_from_slice(&unit.to_le_bytes());
    }
    std::fs::write(&path, &bytes).expect("write utf16le tsv");

    let workbook = formula_io::open_workbook_model(&path).expect("open workbook model");
    assert_eq!(workbook.sheets.len(), 1);
    assert_eq!(workbook.sheets[0].name, "data_no_bom");

    let sheet = workbook
        .sheet_by_name("data_no_bom")
        .expect("data_no_bom sheet missing");
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

#[test]
fn open_workbook_model_sniffs_utf16be_tab_delimited_text_without_bom() {
    let tmp = tempfile::tempdir().expect("temp dir");
    let path = tmp.path().join("data_no_bom.txt");

    let tsv = "col1\tcol2\r\n1\thello\r\n2\tworld\r\n";
    let mut bytes = Vec::new();
    for unit in tsv.encode_utf16() {
        bytes.extend_from_slice(&unit.to_be_bytes());
    }
    std::fs::write(&path, &bytes).expect("write utf16be tsv");

    let workbook = formula_io::open_workbook_model(&path).expect("open workbook model");
    assert_eq!(workbook.sheets.len(), 1);
    assert_eq!(workbook.sheets[0].name, "data_no_bom");

    let sheet = workbook
        .sheet_by_name("data_no_bom")
        .expect("data_no_bom sheet missing");
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

#[test]
fn open_workbook_model_sniffs_utf16le_tab_delimited_text_without_bom_non_ascii() {
    let tmp = tempfile::tempdir().expect("temp dir");
    let path = tmp.path().join("jp_no_bom.txt");

    let header_left = "あ".repeat(200);
    let header_right = "い".repeat(200);
    let row_left = "う".repeat(200);
    let row_right = "え".repeat(200);
    let tsv = format!("{header_left}\t{header_right}\r\n{row_left}\t{row_right}\r\n");
    let mut bytes = Vec::new();
    for unit in tsv.encode_utf16() {
        bytes.extend_from_slice(&unit.to_le_bytes());
    }
    std::fs::write(&path, &bytes).expect("write utf16le tsv");

    let workbook = formula_io::open_workbook_model(&path).expect("open workbook model");
    assert_eq!(workbook.sheets.len(), 1);
    assert_eq!(workbook.sheets[0].name, "jp_no_bom");

    let sheet = workbook
        .sheet_by_name("jp_no_bom")
        .expect("jp_no_bom sheet missing");
    assert_eq!(sheet.value_a1("A1").unwrap(), CellValue::String(row_left));
    assert_eq!(sheet.value_a1("B1").unwrap(), CellValue::String(row_right));
}

#[test]
fn open_workbook_model_sniffs_utf16be_tab_delimited_text_without_bom_non_ascii() {
    let tmp = tempfile::tempdir().expect("temp dir");
    let path = tmp.path().join("jp_no_bom_be.txt");

    let header_left = "あ".repeat(200);
    let header_right = "い".repeat(200);
    let row_left = "う".repeat(200);
    let row_right = "え".repeat(200);
    let tsv = format!("{header_left}\t{header_right}\r\n{row_left}\t{row_right}\r\n");
    let mut bytes = Vec::new();
    for unit in tsv.encode_utf16() {
        bytes.extend_from_slice(&unit.to_be_bytes());
    }
    std::fs::write(&path, &bytes).expect("write utf16be tsv");

    let workbook = formula_io::open_workbook_model(&path).expect("open workbook model");
    assert_eq!(workbook.sheets.len(), 1);
    assert_eq!(workbook.sheets[0].name, "jp_no_bom_be");

    let sheet = workbook
        .sheet_by_name("jp_no_bom_be")
        .expect("jp_no_bom_be sheet missing");
    assert_eq!(sheet.value_a1("A1").unwrap(), CellValue::String(row_left));
    assert_eq!(sheet.value_a1("B1").unwrap(), CellValue::String(row_right));
}

#[test]
fn open_workbook_model_csv_honors_excel_sep_directive() {
    let tmp = tempfile::tempdir().expect("temp dir");
    let path = tmp.path().join("sep.csv");

    std::fs::write(&path, "sep=;\na;b\n1;2\n").expect("write csv");

    let workbook = formula_io::open_workbook_model(&path).expect("open csv workbook model");
    let sheet = workbook.sheet_by_name("sep").expect("sheet missing");

    assert_eq!(sheet.value_a1("A1").unwrap(), CellValue::Number(1.0));
    assert_eq!(sheet.value_a1("B1").unwrap(), CellValue::Number(2.0));
}

#[test]
fn open_workbook_model_sniffs_utf16le_csv_honors_excel_sep_directive() {
    let tmp = tempfile::tempdir().expect("temp dir");
    let path = tmp.path().join("sep_utf16.txt");

    let csv = "sep=;\r\na;b\r\n1,hello;world\r\n2,foo;bar\r\n";
    let mut bytes = vec![0xFF, 0xFE];
    for unit in csv.encode_utf16() {
        bytes.extend_from_slice(&unit.to_le_bytes());
    }
    std::fs::write(&path, &bytes).expect("write utf16 csv");

    let workbook = formula_io::open_workbook_model(&path).expect("open workbook model");
    assert_eq!(workbook.sheets.len(), 1);
    assert_eq!(workbook.sheets[0].name, "sep_utf16");

    let sheet = workbook
        .sheet_by_name("sep_utf16")
        .expect("sep_utf16 sheet missing");
    assert_eq!(
        sheet.value_a1("A1").unwrap(),
        CellValue::String("1,hello".to_string())
    );
    assert_eq!(
        sheet.value_a1("B1").unwrap(),
        CellValue::String("world".to_string())
    );
    assert_eq!(
        sheet.value_a1("A2").unwrap(),
        CellValue::String("2,foo".to_string())
    );
    assert_eq!(
        sheet.value_a1("B2").unwrap(),
        CellValue::String("bar".to_string())
    );
}

#[test]
fn open_workbook_model_rejects_unknown_binary() {
    let mut tmp = tempfile::Builder::new()
        .prefix("binary_")
        .tempfile()
        .expect("tempfile");

    tmp.write_all(b"\x00\x01\x02\x03not csv").expect("write tempfile");

    let err = formula_io::open_workbook_model(tmp.path()).expect_err("expected error");
    match err {
        formula_io::Error::UnsupportedExtension { .. } => {}
        other => panic!("expected UnsupportedExtension, got {other:?}"),
    }
}

#[cfg(not(feature = "parquet"))]
#[test]
fn open_workbook_model_parquet_requires_feature() {
    let parquet_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../packages/data-io/test/fixtures")
        .join("simple.parquet");
    let err = formula_io::open_workbook_model(&parquet_path).expect_err("expected error");
    match err {
        formula_io::Error::ParquetSupportNotEnabled { path } => {
            assert_eq!(path, parquet_path);
        }
        other => panic!("expected ParquetSupportNotEnabled, got {other:?}"),
    }
}

#[cfg(not(feature = "parquet"))]
#[test]
fn open_workbook_parquet_requires_feature() {
    let parquet_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../packages/data-io/test/fixtures")
        .join("simple.parquet");
    let err = formula_io::open_workbook(&parquet_path).expect_err("expected error");
    match err {
        formula_io::Error::ParquetSupportNotEnabled { path } => {
            assert_eq!(path, parquet_path);
        }
        other => panic!("expected ParquetSupportNotEnabled, got {other:?}"),
    }
}

#[cfg(feature = "parquet")]
#[test]
fn open_workbook_model_parquet() {
    let parquet_path = parquet_fixture_path("simple.parquet");
    let workbook = formula_io::open_workbook_model(&parquet_path).expect("open parquet workbook");
    assert_eq!(workbook.sheets.len(), 1);
    assert_eq!(workbook.sheets[0].name, "simple");

    let sheet = workbook.sheet_by_name("simple").expect("sheet missing");
    assert_eq!(sheet.value_a1("A1").unwrap(), CellValue::Number(1.0));
    assert_eq!(
        sheet.value_a1("B1").unwrap(),
        CellValue::String("Alice".to_string())
    );
    assert_eq!(sheet.value_a1("C2").unwrap(), CellValue::Boolean(false));
    assert_eq!(sheet.value_a1("D3").unwrap(), CellValue::Number(3.75));
}

#[cfg(feature = "parquet")]
#[test]
fn open_workbook_model_sniffs_parquet_with_wrong_extension_and_sanitizes_sheet_name() {
    let parquet_path = parquet_fixture_path("simple.parquet");

    let dir = tempfile::tempdir().expect("temp dir");
    // Note: extension intentionally wrong; content sniffing should still treat it as Parquet.
    let path = dir.path().join("bad[name].xlsx");
    std::fs::copy(&parquet_path, &path).expect("copy parquet fixture");

    let workbook = formula_io::open_workbook_model(&path).expect("open workbook model");
    assert_eq!(workbook.sheets.len(), 1);

    let expected = sanitize_sheet_name("bad[name]");
    assert_eq!(workbook.sheets[0].name, expected);

    let sheet = workbook.sheet_by_name(&expected).expect("sheet missing");
    assert_eq!(sheet.value_a1("A1").unwrap(), CellValue::Number(1.0));
    assert_eq!(
        sheet.value_a1("B1").unwrap(),
        CellValue::String("Alice".to_string())
    );
    assert_eq!(sheet.value_a1("C2").unwrap(), CellValue::Boolean(false));
    assert_eq!(sheet.value_a1("D3").unwrap(), CellValue::Number(3.75));
}

#[cfg(feature = "parquet")]
#[test]
fn open_workbook_model_parquet_invalid_sheet_name_falls_back_to_sheet1() {
    let parquet_path = parquet_fixture_path("simple.parquet");

    let dir = tempfile::tempdir().expect("temp dir");
    // Use a filename stem that becomes empty after Excel sheet-name sanitization.
    // `[` and `]` are invalid in sheet names but valid on common filesystems.
    let path = dir.path().join("[].parquet");
    std::fs::copy(&parquet_path, &path).expect("copy parquet fixture");

    let workbook = formula_io::open_workbook_model(&path).expect("open workbook model");
    assert_eq!(workbook.sheets.len(), 1);
    assert_eq!(workbook.sheets[0].name, "Sheet1");
}

#[test]
fn open_workbook_model_csv_invalid_sheet_name_falls_back_to_sheet1() {
    let tmp = tempfile::tempdir().expect("temp dir");
    // Use a filename stem that becomes empty after Excel sheet-name sanitization.
    // `[` and `]` are invalid in sheet names but valid on common filesystems.
    let path = tmp.path().join("[].csv");
    std::fs::write(&path, "col1\n1\n").expect("write csv");

    let workbook = formula_io::open_workbook_model(&path).expect("open workbook model");
    assert_eq!(workbook.sheets.len(), 1);
    assert_eq!(workbook.sheets[0].name, "Sheet1");
}

#[test]
fn open_workbook_model_csv_sanitizes_sheet_name() {
    let dir = tempfile::tempdir().expect("temp dir");
    let csv_path = dir.path().join("bad[name]test.csv");
    std::fs::write(&csv_path, "col1\n1\n2\n").expect("write csv");

    let workbook = formula_io::open_workbook_model(&csv_path).expect("open workbook model");
    assert_eq!(workbook.sheets.len(), 1);
    assert_eq!(workbook.sheets[0].name, "badnametest");

    // Regression check: writing to XLSX should succeed (sheet name must be Excel-valid).
    let mut out = std::io::Cursor::new(Vec::<u8>::new());
    formula_io::xlsx::write_workbook_to_writer(&workbook, &mut out).expect("write xlsx");
}

#[test]
fn open_workbook_model_csv_truncates_sheet_name_to_excel_max_len_in_utf16_units() {
    let dir = tempfile::tempdir().expect("temp dir");
    let prefix = "a".repeat(EXCEL_MAX_SHEET_NAME_LEN - 2);
    // 🙂 is a non-BMP character, so it counts as 2 UTF-16 code units in Excel.
    let long_stem = format!("{prefix}🙂{}", "b".repeat(10));
    let csv_path = dir.path().join(format!("{long_stem}.csv"));
    std::fs::write(&csv_path, "col1\n1\n").expect("write csv");

    let workbook = formula_io::open_workbook_model(&csv_path).expect("open workbook model");
    assert_eq!(workbook.sheets.len(), 1);

    let expected = sanitize_sheet_name(&long_stem);
    assert_eq!(expected.encode_utf16().count(), EXCEL_MAX_SHEET_NAME_LEN);
    assert_eq!(workbook.sheets[0].name, expected);

    // Regression check: writing to XLSX should succeed (sheet name must be Excel-valid).
    let mut out = std::io::Cursor::new(Vec::<u8>::new());
    formula_io::xlsx::write_workbook_to_writer(&workbook, &mut out).expect("write xlsx");
}

#[cfg(feature = "parquet")]
#[test]
fn open_workbook_model_parquet_truncates_sheet_name_to_excel_max_len_in_utf16_units() {
    let parquet_path = parquet_fixture_path("simple.parquet");

    let dir = tempfile::tempdir().expect("temp dir");
    let prefix = "a".repeat(EXCEL_MAX_SHEET_NAME_LEN - 2);
    // 🙂 is a non-BMP character, so it counts as 2 UTF-16 code units in Excel.
    let long_stem = format!("{prefix}🙂{}", "b".repeat(10));
    let path = dir.path().join(format!("{long_stem}.parquet"));
    std::fs::copy(&parquet_path, &path).expect("copy parquet fixture");

    let workbook = formula_io::open_workbook_model(&path).expect("open workbook model");
    assert_eq!(workbook.sheets.len(), 1);

    let expected = sanitize_sheet_name(&long_stem);
    assert_eq!(expected.encode_utf16().count(), EXCEL_MAX_SHEET_NAME_LEN);
    assert_eq!(workbook.sheets[0].name, expected);

    // Regression check: writing to XLSX should succeed (sheet name must be Excel-valid).
    let mut out = std::io::Cursor::new(Vec::<u8>::new());
    formula_io::xlsx::write_workbook_to_writer(&workbook, &mut out).expect("write xlsx");
}
