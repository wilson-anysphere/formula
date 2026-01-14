use std::cell::Cell;

use formula_model::{
    parse_range_a1, validate_sheet_name, CellRef, CellValue, CfRule, CfRuleKind, CfRuleSchema,
    CfStyleOverride, Color, Comment, CommentKind, CommentPatch, DuplicateSheetError,
    FormulaEvaluator, Range, SheetNameError, Table, TableColumn, Workbook,
};

#[test]
fn duplicate_sheet_rewrites_explicit_self_references() {
    let mut wb = Workbook::new();
    let sheet1 = wb.add_sheet("Sheet1").unwrap();

    wb.sheet_mut(sheet1)
        .unwrap()
        .set_formula_a1("B2", Some("=Sheet1!A1".to_string()))
        .unwrap();

    let copied = wb.duplicate_sheet(sheet1, None).unwrap();

    assert_eq!(wb.sheets.len(), 2);
    assert_eq!(wb.sheets[0].id, sheet1);
    assert_eq!(wb.sheets[1].id, copied);

    let copied_sheet = wb.sheet(copied).unwrap();
    assert_eq!(copied_sheet.name, "Sheet1 (2)");
    assert_eq!(
        copied_sheet.formula(CellRef::from_a1("B2").unwrap()),
        Some("'Sheet1 (2)'!A1")
    );

    // The source sheet is unchanged.
    let source_sheet = wb.sheet(sheet1).unwrap();
    assert_eq!(
        source_sheet.formula(CellRef::from_a1("B2").unwrap()),
        Some("Sheet1!A1")
    );
}

#[test]
fn duplicate_sheet_does_not_rewrite_external_workbook_references() {
    let mut wb = Workbook::new();
    let sheet1 = wb.add_sheet("Sheet1").unwrap();

    wb.sheet_mut(sheet1)
        .unwrap()
        .set_formula_a1("B2", Some("='[Book1.xlsx]Sheet1'!A1".to_string()))
        .unwrap();

    let copied = wb.duplicate_sheet(sheet1, None).unwrap();
    let copied_sheet = wb.sheet(copied).unwrap();

    assert_eq!(
        copied_sheet.formula(CellRef::from_a1("B2").unwrap()),
        Some("'[Book1.xlsx]Sheet1'!A1")
    );
}

#[test]
fn duplicate_sheet_renames_tables_and_updates_structured_refs() {
    let mut wb = Workbook::new();
    let sheet1 = wb.add_sheet("Sheet1").unwrap();

    let table = Table {
        id: 1,
        name: "Table1".to_string(),
        display_name: "Table1".to_string(),
        range: Range::from_a1("A1:B3").unwrap(),
        header_row_count: 1,
        totals_row_count: 0,
        columns: vec![
            TableColumn {
                id: 1,
                name: "Col1".to_string(),
                formula: None,
                totals_formula: None,
            },
            TableColumn {
                id: 2,
                name: "Col2".to_string(),
                formula: None,
                totals_formula: None,
            },
        ],
        style: None,
        auto_filter: None,
        relationship_id: None,
        part_path: None,
    };

    {
        let sheet = wb.sheet_mut(sheet1).unwrap();
        sheet.tables.push(table);
        sheet
            .set_formula_a1("C1", Some("=SUM(Table1[Col1])".to_string()))
            .unwrap();
        sheet
            .set_formula_a1("C2", Some("=SUM(Table1)".to_string()))
            .unwrap();
    }

    let copied = wb.duplicate_sheet(sheet1, None).unwrap();
    let copied_sheet = wb.sheet(copied).unwrap();

    assert_eq!(copied_sheet.tables.len(), 1);
    assert_eq!(copied_sheet.tables[0].name, "Table1_1");
    assert_ne!(copied_sheet.tables[0].id, 1);

    assert_eq!(
        copied_sheet.formula(CellRef::from_a1("C1").unwrap()),
        Some("SUM(Table1_1[Col1])")
    );
    assert_eq!(
        copied_sheet.formula(CellRef::from_a1("C2").unwrap()),
        Some("SUM(Table1_1)")
    );

    // The source sheet's table name and formula should be unchanged.
    let source_sheet = wb.sheet(sheet1).unwrap();
    assert_eq!(source_sheet.tables[0].name, "Table1");
    assert_eq!(
        source_sheet.formula(CellRef::from_a1("C1").unwrap()),
        Some("SUM(Table1[Col1])")
    );
    assert_eq!(
        source_sheet.formula(CellRef::from_a1("C2").unwrap()),
        Some("SUM(Table1)")
    );
}

#[test]
fn duplicate_sheet_name_collisions_match_excel_style_suffixes() {
    let mut wb = Workbook::new();
    let sheet1 = wb.add_sheet("Sheet1").unwrap();

    let copy2 = wb.duplicate_sheet(sheet1, None).unwrap();
    assert_eq!(wb.sheet(copy2).unwrap().name, "Sheet1 (2)");

    let copy3 = wb.duplicate_sheet(sheet1, None).unwrap();
    assert_eq!(wb.sheet(copy3).unwrap().name, "Sheet1 (3)");
}

#[test]
fn duplicate_sheet_name_generation_respects_utf16_length_limit() {
    let mut wb = Workbook::new();
    let base = format!("{}A", "😀".repeat(15)); // 15 emoji (30 UTF-16) + 'A' = 31
    let sheet = wb.add_sheet(base).unwrap();

    let copied = wb.duplicate_sheet(sheet, None).unwrap();
    let copy_name = wb.sheet(copied).unwrap().name.clone();

    // Should not exceed Excel's 31 UTF-16 code unit limit.
    validate_sheet_name(&copy_name).unwrap();
    assert_eq!(copy_name.encode_utf16().count(), 30); // 13 emoji (26) + " (2)" (4)
    assert_eq!(copy_name, format!("{} (2)", "😀".repeat(13)));
}

#[test]
fn duplicate_sheet_clears_conditional_formatting_cache_after_rewrites() {
    struct CountingEvaluator {
        calls: Cell<usize>,
    }

    impl CountingEvaluator {
        fn new() -> Self {
            Self {
                calls: Cell::new(0),
            }
        }

        fn calls(&self) -> usize {
            self.calls.get()
        }
    }

    impl FormulaEvaluator for CountingEvaluator {
        fn eval(&self, formula: &str, _ctx: CellRef) -> Option<CellValue> {
            self.calls.set(self.calls.get() + 1);
            Some(CellValue::Boolean(formula.contains("Sheet1!A1")))
        }
    }

    let mut wb = Workbook::new();
    let sheet1 = wb.add_sheet("Sheet1").unwrap();

    let visible = parse_range_a1("A1").unwrap();
    let rules = vec![CfRule {
        schema: CfRuleSchema::Office2007,
        id: Some("1".to_string()),
        priority: 1,
        applies_to: vec![visible],
        dxf_id: Some(0),
        stop_if_true: false,
        kind: CfRuleKind::Expression {
            formula: "Sheet1!A1".to_string(),
        },
        // Depend on the target cell itself so edits would invalidate the cache.
        dependencies: vec![visible],
    }];
    let dxfs = vec![CfStyleOverride {
        fill: Some(Color::new_argb(0xFFFF0000)),
        font_color: None,
        bold: None,
        italic: None,
    }];

    wb.sheet_mut(sheet1)
        .unwrap()
        .set_conditional_formatting(rules, dxfs);

    let evaluator = CountingEvaluator::new();

    // Evaluate conditional formatting on the source sheet to populate its cache.
    {
        let sheet = wb.sheet(sheet1).unwrap();
        let eval = sheet.evaluate_conditional_formatting(visible, sheet, Some(&evaluator));
        assert_eq!(
            eval.get(CellRef::from_a1("A1").unwrap())
                .unwrap()
                .style
                .fill,
            Some(Color::new_argb(0xFFFF0000))
        );
    }
    assert_eq!(evaluator.calls(), 1);

    let copied = wb.duplicate_sheet(sheet1, None).unwrap();

    // After duplication, the CF rule formula is rewritten to the new sheet name. The duplicated
    // sheet must not re-use the source sheet's cached evaluation result.
    let copied_sheet = wb.sheet(copied).unwrap();
    let eval =
        copied_sheet.evaluate_conditional_formatting(visible, copied_sheet, Some(&evaluator));
    assert_eq!(
        eval.get(CellRef::from_a1("A1").unwrap())
            .unwrap()
            .style
            .fill,
        None
    );
    assert_eq!(evaluator.calls(), 2);
}

#[test]
fn duplicate_sheet_copies_comments_and_outline() {
    let mut wb = Workbook::new();
    let sheet1 = wb.add_sheet("Sheet1").unwrap();

    {
        let sheet = wb.sheet_mut(sheet1).unwrap();
        sheet.group_rows(1, 3);
        sheet
            .add_comment(
                CellRef::from_a1("A1").unwrap(),
                Comment {
                    kind: CommentKind::Note,
                    content: "hello".to_string(),
                    ..Comment::default()
                },
            )
            .unwrap();
    };

    let source_outline = wb.sheet(sheet1).unwrap().outline.clone();

    let copied = wb.duplicate_sheet(sheet1, None).unwrap();

    // Comments are copied.
    let copied_comment_id = {
        let copied_sheet = wb.sheet(copied).unwrap();
        let copied_comments = copied_sheet.comments_for_cell(CellRef::from_a1("A1").unwrap());
        assert_eq!(copied_comments.len(), 1);
        assert_eq!(copied_comments[0].content, "hello");

        // Outline state is copied.
        assert_eq!(copied_sheet.outline, source_outline);

        copied_comments[0].id.clone()
    };

    // Mutating the copied sheet should not affect the source sheet.
    wb.sheet_mut(copied)
        .unwrap()
        .update_comment(
            &copied_comment_id,
            CommentPatch {
                content: Some("copied".to_string()),
                ..Default::default()
            },
        )
        .unwrap();
    wb.sheet_mut(copied).unwrap().ungroup_rows(1, 3);

    let source_sheet = wb.sheet(sheet1).unwrap();
    assert_eq!(
        source_sheet.comments_for_cell(CellRef::from_a1("A1").unwrap())[0].content,
        "hello"
    );
    assert_eq!(source_sheet.outline, source_outline);
}

#[test]
fn duplicate_sheet_rejects_duplicate_target_name() {
    let mut wb = Workbook::new();
    let sheet1 = wb.add_sheet("Sheet1").unwrap();
    let _ = wb.add_sheet("Other").unwrap();

    let err = wb.duplicate_sheet(sheet1, Some("Other")).unwrap_err();
    assert_eq!(
        err,
        DuplicateSheetError::InvalidName(SheetNameError::DuplicateName)
    );
}
