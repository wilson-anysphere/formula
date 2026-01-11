use std::path::Path;

use formula_model::{CellRef, CellValue, Workbook};

fn fixture_path(rel: &str) -> std::path::PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("../../").join(rel)
}

fn assert_cell_matches(full: &Workbook, fast: &Workbook, sheet_idx: usize, a1: &str) {
    let cell_ref = CellRef::from_a1(a1).expect("valid cell ref");
    let full_sheet = &full.sheets[sheet_idx];
    let fast_sheet = &fast.sheets[sheet_idx];

    let full_cell = full_sheet.cell(cell_ref);
    let fast_cell = fast_sheet.cell(cell_ref);

    assert_eq!(
        full_cell.map(|c| &c.value),
        fast_cell.map(|c| &c.value),
        "cell value mismatch for {}!{}",
        full_sheet.name,
        a1
    );
    assert_eq!(
        full_cell.and_then(|c| c.formula.as_deref()),
        fast_cell.and_then(|c| c.formula.as_deref()),
        "cell formula mismatch for {}!{}",
        full_sheet.name,
        a1
    );
    assert_eq!(
        full_cell.map(|c| c.style_id),
        fast_cell.map(|c| c.style_id),
        "cell style_id mismatch for {}!{}",
        full_sheet.name,
        a1
    );
}

#[test]
fn fast_reader_matches_full_reader_for_values_and_formulas() {
    struct Case<'a> {
        fixture: &'a str,
        sheet_idx: usize,
        cells: &'a [&'a str],
    }

    let cases = [
        Case {
            fixture: "fixtures/xlsx/basic/basic.xlsx",
            sheet_idx: 0,
            cells: &["A1", "B1"],
        },
        Case {
            fixture: "fixtures/xlsx/formulas/formulas.xlsx",
            sheet_idx: 0,
            cells: &["A1", "B1", "C1"],
        },
        Case {
            fixture: "fixtures/xlsx/styles/rich-text-shared-strings.xlsx",
            sheet_idx: 0,
            cells: &["A1", "A2"],
        },
        // Explicitly exercises `styles.xml` + `c/@s` -> `style_id` mapping.
        Case {
            fixture: "fixtures/xlsx/styles/styles.xlsx",
            sheet_idx: 0,
            cells: &["A1"],
        },
    ];

    for case in cases {
        let bytes = std::fs::read(fixture_path(case.fixture)).expect("read fixture");
        let full = formula_xlsx::load_from_bytes(&bytes)
            .expect("load_from_bytes")
            .workbook;
        let fast = formula_xlsx::read_workbook_model_from_bytes(&bytes)
            .expect("fast reader should succeed");

        let full_names: Vec<_> = full.sheets.iter().map(|s| s.name.as_str()).collect();
        let fast_names: Vec<_> = fast.sheets.iter().map(|s| s.name.as_str()).collect();
        assert_eq!(
            full_names, fast_names,
            "sheet list mismatch for fixture {}",
            case.fixture
        );

        assert_eq!(
            full.styles.len(),
            fast.styles.len(),
            "style table size mismatch for fixture {}",
            case.fixture
        );

        for cell in case.cells {
            assert_cell_matches(&full, &fast, case.sheet_idx, cell);
        }

        // Spot-check expected semantics for the key fixtures so this test catches
        // regressions even if both readers drift together.
        if case.fixture.ends_with("basic.xlsx") {
            assert_eq!(
                full.sheets[0]
                    .cell(CellRef::from_a1("A1").unwrap())
                    .unwrap()
                    .value,
                CellValue::Number(1.0)
            );
            assert_eq!(
                full.sheets[0]
                    .cell(CellRef::from_a1("B1").unwrap())
                    .unwrap()
                    .value,
                CellValue::String("Hello".to_string())
            );
        }

        if case.fixture.ends_with("formulas.xlsx") {
            assert_eq!(
                fast.sheets[0]
                    .cell(CellRef::from_a1("C1").unwrap())
                    .unwrap()
                    .formula
                    .as_deref(),
                Some("A1+B1")
            );
        }

        if case.fixture.ends_with("rich-text-shared-strings.xlsx") {
            assert_eq!(
                fast.sheets[0]
                    .cell(CellRef::from_a1("A1").unwrap())
                    .unwrap()
                    .value,
                CellValue::String("Hello Bold Italic".to_string())
            );
        }
    }
}

#[test]
fn fast_reader_does_not_require_unmodeled_parts() {
    let bytes =
        std::fs::read(fixture_path("fixtures/xlsx/charts/basic-chart.xlsx")).expect("read fixture");
    let workbook = formula_xlsx::read_workbook_model_from_bytes(&bytes)
        .expect("fast reader should ignore chart/drawing parts");
    assert_eq!(workbook.sheets.len(), 1);
    assert_eq!(workbook.sheets[0].name, "Sheet1");
}

