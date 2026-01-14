use formula_model::{
    ArrayValue, CalcSettings, CalculationMode, Cell, CellRef, CellValue, DateSystem, ErrorValue,
    DefinedNameScope, Font, Range, RichText, SheetVisibility, SpillValue, Style, TabColor,
    ThemePalette, WorkbookProtection, WorkbookView, WorkbookWindow, WorkbookWindowState,
};
use formula_model::{Comment, CommentAuthor, CommentKind, SheetSelection};
use formula_model::drawings::{
    Anchor, AnchorPoint, CellOffset, DrawingObject, DrawingObjectId, DrawingObjectKind, EmuSize,
    ImageData, ImageId,
};
use formula_storage::{ImportModelWorkbookOptions, Storage};

use std::collections::BTreeMap;

fn cells_as_map(sheet: &formula_model::Worksheet) -> BTreeMap<(u32, u32), Cell> {
    sheet
        .iter_cells()
        .map(|(cell_ref, cell)| ((cell_ref.row, cell_ref.col), cell.clone()))
        .collect()
}

fn images_as_map(
    workbook: &formula_model::Workbook,
) -> BTreeMap<String, (Vec<u8>, Option<String>)> {
    workbook
        .images
        .iter()
        .map(|(id, data)| {
            (
                id.as_str().to_string(),
                (data.bytes.clone(), data.content_type.clone()),
            )
        })
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
        full_calc_on_load: true,
    };
    workbook.theme = ThemePalette::office_2007();
    workbook.workbook_protection = WorkbookProtection {
        lock_structure: true,
        lock_windows: true,
        password_hash: Some(123),
    };
    let image_id = ImageId::new("image1.png");
    workbook.images.insert(
        image_id.clone(),
        ImageData {
            bytes: vec![0, 1, 2, 3],
            content_type: Some("image/png".to_string()),
        },
    );

    let sheet_a = workbook.add_sheet("Äbc").expect("add sheet A");
    let sheet_b = workbook.add_sheet("Data").expect("add sheet B");
    workbook
        .create_defined_name(
            DefinedNameScope::Workbook,
            "MyGlobalName",
            "=Data!$A$1",
            Some("comment".to_string()),
            true,
            None,
        )
        .expect("create workbook defined name");
    workbook
        .create_defined_name(
            DefinedNameScope::Sheet(sheet_a),
            "MyLocalName",
            "Data!$A$2",
            None,
            false,
            Some(7),
        )
        .expect("create sheet defined name");
    assert!(
        workbook.set_sheet_print_area(
            sheet_b,
            Some(vec![Range::new(CellRef::new(0, 0), CellRef::new(10, 3))]),
        ),
        "set print area"
    );
    workbook.view = WorkbookView {
        active_sheet_id: Some(sheet_b),
        window: Some(WorkbookWindow {
            x: Some(10),
            y: Some(20),
            width: Some(800),
            height: Some(600),
            state: Some(WorkbookWindowState::Maximized),
        }),
    };

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
        sheet.drawings.push(DrawingObject {
            id: DrawingObjectId(1),
            kind: DrawingObjectKind::Image {
                image_id: image_id.clone(),
            },
            anchor: Anchor::OneCell {
                from: AnchorPoint::new(CellRef::new(0, 0), CellOffset::new(0, 0)),
                ext: EmuSize::new(100, 200),
            },
            z_order: 0,
            size: Some(EmuSize::new(100, 200)),
            preserved: Default::default(),
        });
        sheet.view.show_grid_lines = false;
        sheet.view.selection = Some(SheetSelection::new(
            CellRef::new(5, 5),
            vec![Range::new(CellRef::new(5, 5), CellRef::new(6, 6))],
        ));
        sheet.set_row_height(10, Some(20.0));
        sheet.set_row_hidden(10, true);
        sheet.set_col_width(3, Some(12.0));
        sheet.merged_regions
            .add(Range::new(CellRef::new(10, 10), CellRef::new(11, 11)))
            .expect("add merge");
        sheet.outline.pr.summary_below = false;
        sheet.outline.group_rows(1, 2);
        sheet.add_comment(
            CellRef::new(5, 5),
            Comment {
                id: String::new(),
                cell_ref: CellRef::new(0, 0),
                author: CommentAuthor {
                    id: "u1".to_string(),
                    name: "Alice".to_string(),
                },
                created_at: 1,
                updated_at: 2,
                resolved: false,
                kind: CommentKind::Threaded,
                content: "Hello".to_string(),
                mentions: Vec::new(),
                replies: Vec::new(),
            },
        )
        .expect("add comment");

        sheet.set_cell(
            CellRef::new(0, 0),
            Cell {
                value: CellValue::Number(1.0),
                phonetic: None,
                formula: None,
                phonetic: None,
                style_id: bold_id,
            },
        );
        sheet.set_cell(
            CellRef::new(0, 1),
            Cell {
                value: CellValue::Empty,
                phonetic: None,
                formula: Some("SUM(A1)".to_string()),
                phonetic: None,
                style_id: 0,
            },
        );
        // Very sparse cell.
        sheet.set_cell(
            CellRef::new(1_000_000, 10),
            Cell {
                value: CellValue::String("far".to_string()),
                phonetic: None,
                formula: None,
                phonetic: None,
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
                phonetic: None,
                formula: None,
                phonetic: None,
                style_id: plain_id,
            },
        );
        sheet.set_cell(
            CellRef::new(1, 0),
            Cell {
                value: CellValue::Array(array.clone()),
                phonetic: None,
                formula: None,
                phonetic: None,
                style_id: 0,
            },
        );
        sheet.set_cell(
            CellRef::new(2, 0),
            Cell {
                value: CellValue::Spill(SpillValue {
                    origin: CellRef::new(1, 0),
                }),
                phonetic: None,
                formula: None,
                phonetic: None,
                style_id: 0,
            },
        );
        sheet.set_cell(
            CellRef::new(3, 0),
            Cell {
                value: CellValue::Error(ErrorValue::Div0),
                phonetic: None,
                formula: None,
                phonetic: None,
                style_id: 0,
            },
        );
    }

    let storage = Storage::open_in_memory().expect("open storage");
    let meta = storage
        .import_model_workbook(&workbook, ImportModelWorkbookOptions::new("ModelBook"))
        .expect("import workbook");

    // Imported defined names should also be visible through the legacy named-ranges API.
    let global = storage
        .get_named_range(meta.id, "MyGlobalName", "workbook")
        .expect("get named range")
        .expect("global name exists");
    assert_eq!(global.reference, "Data!$A$1");
    let local = storage
        .get_named_range(meta.id, "MyLocalName", "Äbc")
        .expect("get named range")
        .expect("local name exists");
    assert_eq!(local.reference, "Data!$A$2");

    // Ensure legacy tab-color API overrides the richer tab_color_json persisted during import.
    let sheet_a_storage_id = storage
        .list_sheets(meta.id)
        .expect("list sheets")
        .into_iter()
        .find(|s| s.name == "Äbc")
        .expect("sheet a")
        .id;
    let tab_color = TabColor::rgb("FF112233");
    storage
        .set_sheet_tab_color(sheet_a_storage_id, Some(&tab_color))
        .expect("set tab color");

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
    assert_eq!(exported.theme, workbook.theme);
    assert_eq!(exported.workbook_protection, workbook.workbook_protection);
    assert_eq!(exported.defined_names, workbook.defined_names);
    assert_eq!(exported.print_settings, workbook.print_settings);
    assert_eq!(exported.view, workbook.view);
    assert_eq!(exported.styles.styles, workbook.styles.styles);
    assert_eq!(images_as_map(&exported), images_as_map(&workbook));

    assert_eq!(exported.sheets.len(), workbook.sheets.len());
    for (expected, actual) in workbook.sheets.iter().zip(exported.sheets.iter()) {
        assert_eq!(actual.id, expected.id);
        assert_eq!(actual.name, expected.name);
        assert_eq!(actual.visibility, expected.visibility);
        if actual.name == "Äbc" {
            assert_eq!(actual.tab_color, Some(TabColor::rgb("FF112233")));
        } else {
            assert_eq!(actual.tab_color, expected.tab_color);
        }
        assert_eq!(actual.xlsx_sheet_id, expected.xlsx_sheet_id);
        assert_eq!(actual.xlsx_rel_id, expected.xlsx_rel_id);
        assert_eq!(actual.frozen_rows, expected.frozen_rows);
        assert_eq!(actual.frozen_cols, expected.frozen_cols);
        assert!((actual.zoom - expected.zoom).abs() < f32::EPSILON);
        assert_eq!(actual.drawings, expected.drawings);
        assert_eq!(actual.view, expected.view);
        assert_eq!(actual.row_properties, expected.row_properties);
        assert_eq!(actual.col_properties, expected.col_properties);
        assert_eq!(actual.outline, expected.outline);
        assert_eq!(actual.merged_regions.regions, expected.merged_regions.regions);

        let expected_comments: Vec<_> = expected
            .iter_comments()
            .map(|(cell, comment)| (cell, comment.clone()))
            .collect();
        let actual_comments: Vec<_> = actual
            .iter_comments()
            .map(|(cell, comment)| (cell, comment.clone()))
            .collect();
        assert_eq!(actual_comments, expected_comments);

        assert_eq!(cells_as_map(actual), cells_as_map(expected));
    }
}
