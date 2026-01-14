use std::collections::HashMap;
use std::path::{Path, PathBuf};

use formula_biff::encode_rgce;
use formula_xlsb::rgce::{encode_rgce_with_context, CellCoord};
use formula_xlsb::{CellEdit, CellValue, XlsbWorkbook};

mod fixture_builder;
use fixture_builder::XlsbFixtureBuilder;

fn fixture_styles_date_path() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures_styles/date.xlsb")
}

fn assert_cell(
    wb: &XlsbWorkbook,
    row: u32,
    col: u32,
    expected_style: u32,
    expected_value: CellValue,
    expected_formula_text: &str,
) {
    let sheet = wb.read_sheet(0).expect("read sheet");
    let cell = sheet
        .cells
        .iter()
        .find(|c| c.row == row && c.col == col)
        .unwrap_or_else(|| panic!("missing cell at ({row}, {col})"));

    assert_eq!(cell.style, expected_style);
    assert_eq!(cell.value, expected_value);

    let formula = cell.formula.as_ref().expect("expected formula");
    assert_eq!(formula.text.as_deref(), Some(expected_formula_text));
}

#[test]
fn can_convert_existing_value_cell_to_formula_preserving_style_in_both_patchers() {
    let fixture_path = fixture_styles_date_path();
    let wb = XlsbWorkbook::open(&fixture_path).expect("open xlsb fixture");

    // In fixtures_styles/date.xlsb, A1 is a date-serialized number with a non-zero style.
    let original_sheet = wb.read_sheet(0).expect("read sheet");
    let original = original_sheet
        .cells
        .iter()
        .find(|c| c.row == 0 && c.col == 0)
        .expect("A1 missing");
    let original_style = original.style;
    assert_ne!(original_style, 0, "fixture should use a non-zero style");

    let encoded = encode_rgce_with_context("=1+1", wb.workbook_context(), CellCoord::new(0, 0))
        .expect("encode rgce");
    assert!(encoded.rgcb.is_empty(), "expected non-array formula");

    let edits = [CellEdit {
        row: 0,
        col: 0,
        new_value: CellValue::Number(2.0),
        new_style: None,
        clear_formula: false,
        new_formula: Some(encoded.rgce.clone()),
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
    }];

    let tmpdir = tempfile::tempdir().expect("create tempdir");
    let in_memory_path = tmpdir.path().join("patched_in_memory.xlsb");
    let streaming_path = tmpdir.path().join("patched_streaming.xlsb");

    wb.save_with_cell_edits(&in_memory_path, 0, &edits)
        .expect("save_with_cell_edits");
    wb.save_with_cell_edits_streaming(&streaming_path, 0, &edits)
        .expect("save_with_cell_edits_streaming");

    let wb_in_memory = XlsbWorkbook::open(&in_memory_path).expect("open patched workbook");
    assert_cell(
        &wb_in_memory,
        0,
        0,
        original_style,
        CellValue::Number(2.0),
        "1+1",
    );

    let wb_streaming = XlsbWorkbook::open(&streaming_path).expect("open patched workbook");
    assert_cell(
        &wb_streaming,
        0,
        0,
        original_style,
        CellValue::Number(2.0),
        "1+1",
    );
}

#[test]
fn can_convert_various_value_record_types_to_formula_records() {
    // Build a worksheet that includes all supported non-formula value record types:
    // - FLOAT, NUM (RK), CELL_ST (inline string), STRING (shared string), BOOL, BOOLERR, BLANK.
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_sheet_name("Convert");

    builder.set_cell_number(0, 0, 1.0); // BrtCellReal (FLOAT)
    builder.set_cell_number_rk(0, 1, 2.0); // BrtCellRk (NUM)
    builder.set_cell_inline_string(0, 2, "Hello"); // BrtCellSt (CELL_ST)

    let isst = builder.add_shared_string("Shared");
    builder.set_cell_sst(0, 3, isst); // BrtCellIsst (STRING)

    builder.set_cell_bool(0, 4, true); // BrtCellBool
    builder.set_cell_error(0, 5, 0x07); // BrtCellBoolErr (#DIV/0!)
    builder.set_cell_blank(0, 6); // BrtBlank

    let tmpdir = tempfile::tempdir().expect("create temp dir");
    let input_path = tmpdir.path().join("input.xlsb");
    std::fs::write(&input_path, builder.build_bytes()).expect("write xlsb fixture");

    let wb = XlsbWorkbook::open(&input_path).expect("open xlsb fixture");
    let original_sheet = wb.read_sheet(0).expect("read sheet");
    let original_styles: HashMap<(u32, u32), u32> = original_sheet
        .cells
        .iter()
        .map(|c| ((c.row, c.col), c.style))
        .collect();

    let ctx = wb.workbook_context();
    // Use `formula_biff` for string literals: the minimal encoder used by
    // `formula_xlsb::rgce::encode_rgce_with_context` intentionally supports only a small grammar.
    let _ = ctx;
    let num = encode_rgce("=1+1").expect("encode formula");
    let text = encode_rgce("=\"World\"").expect("encode formula");
    let boolf = encode_rgce("=FALSE").expect("encode formula");
    let err = encode_rgce("=1/0").expect("encode formula");

    let edits = vec![
        // FLOAT -> BrtFmlaNum
        CellEdit {
            row: 0,
            col: 0,
            new_value: CellValue::Number(2.0),
            new_style: None,
            clear_formula: false,
            new_formula: Some(num.clone()),
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
        },
        // NUM (RK) -> BrtFmlaNum
        CellEdit {
            row: 0,
            col: 1,
            new_value: CellValue::Number(4.0),
            new_style: None,
            clear_formula: false,
            new_formula: Some(num.clone()),
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
        },
        // CELL_ST -> BrtFmlaString
        CellEdit {
            row: 0,
            col: 2,
            new_value: CellValue::Text("World".to_string()),
            new_style: None,
            clear_formula: false,
            new_formula: Some(text.clone()),
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
        },
        // STRING (SST) -> BrtFmlaNum (cross-type conversion)
        CellEdit {
            row: 0,
            col: 3,
            new_value: CellValue::Number(5.0),
            new_style: None,
            clear_formula: false,
            new_formula: Some(num.clone()),
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
        },
        // BOOL -> BrtFmlaBool
        CellEdit {
            row: 0,
            col: 4,
            new_value: CellValue::Bool(false),
            new_style: None,
            clear_formula: false,
            new_formula: Some(boolf.clone()),
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
        },
        // BOOLERR -> BrtFmlaError
        CellEdit {
            row: 0,
            col: 5,
            new_value: CellValue::Error(0x07),
            new_style: None,
            clear_formula: false,
            new_formula: Some(err.clone()),
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
        },
        // BLANK -> BrtFmlaNum
        CellEdit {
            row: 0,
            col: 6,
            new_value: CellValue::Number(6.0),
            new_style: None,
            clear_formula: false,
            new_formula: Some(num.clone()),
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
        },
    ];

    let in_memory_path = tmpdir.path().join("patched_in_memory.xlsb");
    let streaming_path = tmpdir.path().join("patched_streaming.xlsb");

    wb.save_with_cell_edits(&in_memory_path, 0, &edits)
        .expect("save_with_cell_edits");
    wb.save_with_cell_edits_streaming(&streaming_path, 0, &edits)
        .expect("save_with_cell_edits_streaming");

    for out_path in [&in_memory_path, &streaming_path] {
        let patched = XlsbWorkbook::open(out_path).expect("open patched workbook");
        let sheet = patched.read_sheet(0).expect("read patched sheet");
        for cell in &sheet.cells {
            let expected_style = *original_styles
                .get(&(cell.row, cell.col))
                .expect("style exists");
            assert_eq!(
                cell.style, expected_style,
                "style should be preserved for ({}, {})",
                cell.row, cell.col
            );
        }

        assert_cell(
            &patched,
            0,
            0,
            original_styles[&(0, 0)],
            CellValue::Number(2.0),
            "1+1",
        );
        assert_cell(
            &patched,
            0,
            1,
            original_styles[&(0, 1)],
            CellValue::Number(4.0),
            "1+1",
        );
        assert_cell(
            &patched,
            0,
            2,
            original_styles[&(0, 2)],
            CellValue::Text("World".to_string()),
            "\"World\"",
        );
        assert_cell(
            &patched,
            0,
            3,
            original_styles[&(0, 3)],
            CellValue::Number(5.0),
            "1+1",
        );
        assert_cell(
            &patched,
            0,
            4,
            original_styles[&(0, 4)],
            CellValue::Bool(false),
            "FALSE",
        );
        assert_cell(
            &patched,
            0,
            5,
            original_styles[&(0, 5)],
            CellValue::Error(0x07),
            "1/0",
        );
        assert_cell(
            &patched,
            0,
            6,
            original_styles[&(0, 6)],
            CellValue::Number(6.0),
            "1+1",
        );
    }
}

#[test]
fn cannot_convert_to_formula_with_blank_cached_value() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_number(0, 0, 1.0);

    let tmpdir = tempfile::tempdir().expect("create temp dir");
    let input_path = tmpdir.path().join("input.xlsb");
    std::fs::write(&input_path, builder.build_bytes()).expect("write xlsb fixture");

    let wb = XlsbWorkbook::open(&input_path).expect("open xlsb fixture");
    let ctx = wb.workbook_context();
    let encoded = encode_rgce_with_context("=1+1", ctx, CellCoord::new(0, 0)).expect("encode rgce");

    let edits = [CellEdit {
        row: 0,
        col: 0,
        new_value: CellValue::Blank,
        new_style: None,
        clear_formula: false,
        new_formula: Some(encoded.rgce),
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
    }];

    let out_path = tmpdir.path().join("out.xlsb");
    let err = wb
        .save_with_cell_edits(&out_path, 0, &edits)
        .expect_err("expected error");
    let msg = err.to_string();
    assert!(
        msg.contains("blank cached value"),
        "unexpected error message: {msg}"
    );
}
