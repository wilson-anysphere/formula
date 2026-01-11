use formula_model::{
    ArrayValue, CalcSettings, CalculationMode, Cell, CellRef, CellValue, DateSystem, ErrorValue,
    Font, RichText, SheetVisibility, SpillValue, Style, TabColor,
};
use formula_storage::{ImportModelWorkbookOptions, Storage};

use std::collections::BTreeMap;

fn cells_as_map(sheet: &formula_model::Worksheet) -> BTreeMap<(u32, u32), Cell> {
    sheet
        .iter_cells()
        .map(|(cell_ref, cell)| ((cell_ref.row, cell_ref.col), cell.clone()))
        .collect()
}

#[test]
fn model_workbook_import_export_round_trips() {
    let mut workbook = formula_model::Workbook::new();
    workbook.id = 99;
    workbook.schema_version = formula_model::SCHEMA_VERSION;
    workbook.date_system = DateSystem::Excel1904;
    workbook.calc_settings = CalcSettings {
        calculation_mode: CalculationMode::Manual,
        calculate_before_save: false,
        iterative: formula_model::IterativeCalculationSettings {
            enabled: true,
            max_iterations: 7,
            max_change: 0.1,
        },
        full_precision: false,
    };

    let sheet_a = workbook.add_sheet("Äbc").expect("add sheet A");
    let sheet_b = workbook.add_sheet("Data").expect("add sheet B");

    // Add some styles to the workbook style table.
    let style_bold = Style {
        font: Some(Font {
            bold: true,
            ..Default::default()
        }),
        number_format: Some("0.00".to_string()),
        ..Default::default()
    };
    let bold_id = workbook.intern_style(style_bold.clone());

    let style_plain = Style {
        number_format: Some("@".to_string()),
        ..Default::default()
    };
    let plain_id = workbook.intern_style(style_plain.clone());

    {
        let sheet = workbook.sheet_mut(sheet_a).unwrap();
        sheet.visibility = SheetVisibility::VeryHidden;
        sheet.xlsx_sheet_id = Some(42);
        sheet.xlsx_rel_id = Some("rId7".to_string());
        sheet.frozen_rows = 2;
        sheet.frozen_cols = 1;
        sheet.zoom = 1.25;
        sheet.view.pane.frozen_rows = sheet.frozen_rows;
        sheet.view.pane.frozen_cols = sheet.frozen_cols;
        sheet.view.zoom = sheet.zoom;
        // Tab color using theme + tint to ensure we preserve the full struct, not just rgb.
        sheet.tab_color = Some(TabColor {
            theme: Some(3),
            tint: Some(0.5),
            ..Default::default()
        });

        sheet.set_cell(
            CellRef::new(0, 0),
            Cell {
                value: CellValue::Number(1.0),
                formula: None,
                style_id: bold_id,
            },
        );
        sheet.set_cell(
            CellRef::new(0, 1),
            Cell {
                value: CellValue::Empty,
                formula: Some("SUM(A1)".to_string()),
                style_id: 0,
            },
        );
        // Very sparse cell.
        sheet.set_cell(
            CellRef::new(1_000_000, 10),
            Cell {
                value: CellValue::String("far".to_string()),
                formula: None,
                style_id: 0,
            },
        );
    }

    let rich = RichText::from_segments(vec![
        ("Hello ".to_string(), Default::default()),
        (
            "World".to_string(),
            formula_model::rich_text::RichTextRunStyle {
                bold: Some(true),
                ..Default::default()
            },
        ),
    ]);
    let array = ArrayValue {
        data: vec![
            vec![CellValue::Number(1.0), CellValue::String("x".to_string())],
            vec![CellValue::Boolean(true), CellValue::Error(ErrorValue::NA)],
        ],
    };

    {
        let sheet = workbook.sheet_mut(sheet_b).unwrap();
        sheet.visibility = SheetVisibility::Hidden;
        sheet.tab_color = Some(TabColor::rgb("FF00FF00"));

        sheet.set_cell(
            CellRef::new(0, 0),
            Cell {
                value: CellValue::RichText(rich.clone()),
                formula: None,
                style_id: plain_id,
            },
        );
        sheet.set_cell(
            CellRef::new(1, 0),
            Cell {
                value: CellValue::Array(array.clone()),
                formula: None,
                style_id: 0,
            },
        );
        sheet.set_cell(
            CellRef::new(2, 0),
            Cell {
                value: CellValue::Spill(SpillValue {
                    origin: CellRef::new(1, 0),
                }),
                formula: None,
                style_id: 0,
            },
        );
        sheet.set_cell(
            CellRef::new(3, 0),
            Cell {
                value: CellValue::Error(ErrorValue::Div0),
                formula: None,
                style_id: 0,
            },
        );
    }

    let storage = Storage::open_in_memory().expect("open storage");
    let meta = storage
        .import_model_workbook(&workbook, ImportModelWorkbookOptions::new("ModelBook"))
        .expect("import workbook");

    // Sparse invariant: only stored (non-truly-empty) cells should be persisted.
    let total_model_cells: u64 = workbook
        .sheets
        .iter()
        .map(|s| s.iter_cells().count() as u64)
        .sum();
    let total_db_cells: u64 = storage
        .list_sheets(meta.id)
        .expect("list sheets")
        .iter()
        .map(|s| storage.cell_count(s.id).expect("cell count"))
        .sum();
    assert_eq!(total_db_cells, total_model_cells);

    let exported = storage
        .export_model_workbook(meta.id)
        .expect("export workbook");

    assert_eq!(exported.id, workbook.id);
    assert_eq!(exported.date_system, workbook.date_system);
    assert_eq!(exported.calc_settings, workbook.calc_settings);
    assert_eq!(exported.styles.styles, workbook.styles.styles);

    assert_eq!(exported.sheets.len(), workbook.sheets.len());
    for (expected, actual) in workbook.sheets.iter().zip(exported.sheets.iter()) {
        assert_eq!(actual.id, expected.id);
        assert_eq!(actual.name, expected.name);
        assert_eq!(actual.visibility, expected.visibility);
        assert_eq!(actual.tab_color, expected.tab_color);
        assert_eq!(actual.xlsx_sheet_id, expected.xlsx_sheet_id);
        assert_eq!(actual.xlsx_rel_id, expected.xlsx_rel_id);
        assert_eq!(actual.frozen_rows, expected.frozen_rows);
        assert_eq!(actual.frozen_cols, expected.frozen_cols);
        assert!((actual.zoom - expected.zoom).abs() < f32::EPSILON);

        assert_eq!(cells_as_map(actual), cells_as_map(expected));
    }
}

