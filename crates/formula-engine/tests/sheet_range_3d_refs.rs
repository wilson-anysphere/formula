use formula_engine::{parse_formula, Engine, ErrorKind, ParseOptions, SerializeOptions, Value};
use pretty_assertions::assert_eq;

#[test]
fn parse_and_roundtrip_sheet_range_ref() {
    let ast = parse_formula("=SUM(Sheet1:Sheet3!A1)", ParseOptions::default()).unwrap();
    let roundtrip = ast.to_string(SerializeOptions::default()).unwrap();
    assert_eq!(roundtrip, "=SUM(Sheet1:Sheet3!A1)");
}

#[test]
fn parses_quoted_sheet_range_prefix() {
    let ast = parse_formula("=SUM('Sheet1:Sheet3'!A1)", ParseOptions::default()).unwrap();
    let roundtrip = ast.to_string(SerializeOptions::default()).unwrap();
    // Canonical serialization does not preserve the single-quoted span; it emits the equivalent
    // unquoted form when possible.
    assert_eq!(roundtrip, "=SUM(Sheet1:Sheet3!A1)");
}

#[test]
fn roundtrip_preserves_single_quoted_sheet_range_when_required() {
    let ast = parse_formula("=SUM('Sheet 1:Sheet 3'!A1)", ParseOptions::default()).unwrap();
    let roundtrip = ast.to_string(SerializeOptions::default()).unwrap();
    assert_eq!(roundtrip, "=SUM('Sheet 1:Sheet 3'!A1)");
}

#[test]
fn parses_double_quoted_sheet_range_prefix() {
    let ast = parse_formula("=SUM('Sheet 1':'Sheet 3'!A1)", ParseOptions::default()).unwrap();
    let roundtrip = ast.to_string(SerializeOptions::default()).unwrap();
    assert_eq!(roundtrip, "=SUM('Sheet 1:Sheet 3'!A1)");
}

#[test]
fn collapses_degenerate_sheet_range_refs_to_single_sheet() {
    let ast = parse_formula("=SUM(Sheet1:Sheet1!A1)", ParseOptions::default()).unwrap();
    let roundtrip = ast.to_string(SerializeOptions::default()).unwrap();
    assert_eq!(roundtrip, "=SUM(Sheet1!A1)");

    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 4.0).unwrap();
    engine
        .set_cell_formula("Summary", "A1", "=SUM(Sheet1:Sheet1!A1)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Summary", "A1"), Value::Number(4.0));
}

#[test]
fn evaluates_sum_over_sheet_range_cell_ref() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "A1", 2.0).unwrap();
    engine.set_cell_value("Sheet3", "A1", 3.0).unwrap();

    engine
        .set_cell_formula("Summary", "A1", "=SUM(Sheet1:Sheet3!A1)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Summary", "A1"), Value::Number(6.0));
}

#[test]
fn reorder_sheet_updates_sheet_range_expansion() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "A1", 2.0).unwrap();
    engine.set_cell_value("Sheet3", "A1", 3.0).unwrap();
    engine.set_cell_value("Sheet4", "A1", 10.0).unwrap();

    engine
        .set_cell_formula("Summary", "A1", "=SUM(Sheet1:Sheet3!A1)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Summary", "A1"), Value::Number(6.0));

    // Insert Sheet4 between Sheet1 and Sheet3 so the sheet span expansion changes.
    assert!(engine.reorder_sheet("Sheet4", 1));
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Summary", "A1"), Value::Number(16.0));
}

#[test]
fn evaluates_sum_over_quoted_sheet_range_with_spaces() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet 1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet 2", "A1", 2.0).unwrap();
    engine.set_cell_value("Sheet 3", "A1", 3.0).unwrap();

    engine
        .set_cell_formula("Summary", "A1", "=SUM('Sheet 1:Sheet 3'!A1)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Summary", "A1"), Value::Number(6.0));
}

#[test]
fn evaluates_sum_over_sheet_range_area_ref() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 10.0).unwrap();
    engine.set_cell_value("Sheet2", "A1", 2.0).unwrap();
    engine.set_cell_value("Sheet2", "A2", 20.0).unwrap();
    engine.set_cell_value("Sheet3", "A1", 3.0).unwrap();
    engine.set_cell_value("Sheet3", "A2", 30.0).unwrap();

    engine
        .set_cell_formula("Summary", "A1", "=SUM(Sheet1:Sheet3!A1:A2)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Summary", "A1"), Value::Number(66.0));
}

#[test]
fn index_over_sheet_range_uses_tab_order_after_reorder() {
    // This validates that evaluation-time ordering of multi-area 3D references (and therefore
    // INDEX(..., area_num)) follows workbook tab order, not stable numeric sheet id.
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "A1", 2.0).unwrap();
    engine.set_cell_value("Sheet3", "A1", 3.0).unwrap();

    engine
        .set_cell_formula("Summary", "A1", "=SUM(INDEX(Sheet1:Sheet3!A1,1,1,1))")
        .unwrap();

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Summary", "A1"), Value::Number(1.0));

    // Reverse the sheet tab order: Sheet3, Sheet2, Sheet1, Summary.
    assert!(engine.reorder_sheet("Sheet3", 0));
    assert!(engine.reorder_sheet("Sheet2", 1));

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Summary", "A1"), Value::Number(3.0));
}

#[test]
fn index_over_reversed_sheet_range_uses_tab_order() {
    // Reversed 3D spans (e.g. `Sheet3:Sheet1`) should refer to the same set of sheets as the
    // forward span, ordered by workbook tab order.
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "A1", 2.0).unwrap();
    engine.set_cell_value("Sheet3", "A1", 3.0).unwrap();

    engine
        .set_cell_formula("Summary", "A1", "=SUM(INDEX(Sheet3:Sheet1!A1,1,1,1))")
        .unwrap();

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Summary", "A1"), Value::Number(1.0));

    // Reverse the sheet tab order: Sheet3, Sheet2, Sheet1, Summary.
    assert!(engine.reorder_sheet("Sheet3", 0));
    assert!(engine.reorder_sheet("Sheet2", 1));

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Summary", "A1"), Value::Number(3.0));
}

#[test]
fn sheet_range_membership_updates_after_reorder() {
    // Reordering sheets can change which intermediate sheets are included in a 3D span. Ensure
    // both the AST evaluator and the bytecode backend reflect the updated membership after
    // `reorder_sheet` triggers a dependency-graph rebuild.
    fn setup(engine: &mut Engine) {
        engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
        engine.set_cell_value("Sheet2", "A1", 2.0).unwrap();
        engine.set_cell_value("Sheet3", "A1", 3.0).unwrap();
        engine.set_cell_value("Sheet4", "A1", 4.0).unwrap();

        // Span boundaries are Sheet1 and Sheet3; Sheet2 is initially between them.
        engine
            .set_cell_formula("Summary", "A1", "=SUM(Sheet1:Sheet3!A1)")
            .unwrap();
    }

    let mut engine = Engine::new();
    setup(&mut engine);
    assert_eq!(
        engine.bytecode_compile_report(10),
        Vec::new(),
        "expected SUM over 3D span to compile to bytecode"
    );
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Summary", "A1"), Value::Number(6.0));

    // Move Sheet2 outside of the Sheet1..Sheet3 span.
    assert!(engine.reorder_sheet("Sheet2", 3));
    assert_eq!(
        engine.bytecode_compile_report(10),
        Vec::new(),
        "expected formula to remain bytecode-compiled after reorder"
    );
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Summary", "A1"), Value::Number(4.0));

    // Repeat the same check on the AST evaluator.
    let mut ast_engine = Engine::new();
    ast_engine.set_bytecode_enabled(false);
    setup(&mut ast_engine);
    ast_engine.recalculate_single_threaded();
    assert_eq!(
        ast_engine.get_cell_value("Summary", "A1"),
        Value::Number(6.0)
    );

    assert!(ast_engine.reorder_sheet("Sheet2", 3));
    ast_engine.recalculate_single_threaded();
    assert_eq!(
        ast_engine.get_cell_value("Summary", "A1"),
        Value::Number(4.0)
    );
}

#[test]
fn bytecode_sum_over_sheet_range_uses_tab_order_after_reorder_for_error_precedence() {
    // The bytecode backend expands 3D sheet spans at compile time. Ensure that expansion follows
    // workbook tab order (not the textual order of the boundary sheets).
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine
        .set_cell_value("Sheet2", "A1", Value::Error(ErrorKind::Ref))
        .unwrap();
    engine
        .set_cell_value("Sheet3", "A1", Value::Error(ErrorKind::Div0))
        .unwrap();

    engine
        .set_cell_formula("Summary", "A1", "=SUM(Sheet1:Sheet3!A1)")
        .unwrap();

    // Ensure the formula is bytecode-compiled so this test covers the bytecode span expander.
    assert_eq!(
        engine.bytecode_compile_report(10),
        Vec::new(),
        "expected formula to compile to bytecode"
    );

    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Summary", "A1"),
        Value::Error(ErrorKind::Ref)
    );

    // Reverse the sheet tab order: Sheet3, Sheet2, Sheet1, Summary.
    assert!(engine.reorder_sheet("Sheet3", 0));
    assert!(engine.reorder_sheet("Sheet2", 1));

    // Reordering triggers a dependency-graph rebuild + bytecode recompilation; ensure we are still
    // on the bytecode backend.
    assert_eq!(
        engine.bytecode_compile_report(10),
        Vec::new(),
        "expected formula to remain bytecode-compiled after reorder"
    );

    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Summary", "A1"),
        Value::Error(ErrorKind::Div0)
    );
}

#[test]
fn bytecode_concat_over_sheet_range_uses_tab_order_after_reorder() {
    // CONCAT is order-dependent, so it should concatenate referenced cells in workbook tab order
    // for 3D sheet spans.
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", "1").unwrap();
    engine.set_cell_value("Sheet2", "A1", "2").unwrap();
    engine.set_cell_value("Sheet3", "A1", "3").unwrap();

    engine
        .set_cell_formula("Summary", "A1", "=CONCAT(Sheet1:Sheet3!A1)")
        .unwrap();

    // Ensure the formula is bytecode-compiled so this test covers the bytecode 3D span expander.
    assert_eq!(
        engine.bytecode_compile_report(10),
        Vec::new(),
        "expected CONCAT over 3D span to compile to bytecode"
    );

    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Summary", "A1"),
        Value::Text("123".into())
    );

    // Reverse the sheet tab order: Sheet3, Sheet2, Sheet1, Summary.
    assert!(engine.reorder_sheet("Sheet3", 0));
    assert!(engine.reorder_sheet("Sheet2", 1));

    // Reordering triggers a dependency-graph rebuild + bytecode recompilation; ensure we are still
    // on the bytecode backend.
    assert_eq!(
        engine.bytecode_compile_report(10),
        Vec::new(),
        "expected formula to remain bytecode-compiled after reorder"
    );

    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Summary", "A1"),
        Value::Text("321".into())
    );
}

#[test]
fn evaluates_sum_over_sheet_range_column_range() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet3", "A3", 3.0).unwrap();
    engine.set_cell_value("Sheet2", "B1", 100.0).unwrap();

    engine
        .set_cell_formula("Summary", "A1", "=SUM(Sheet1:Sheet3!A:A)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Summary", "A1"), Value::Number(6.0));
}

#[test]
fn evaluates_sum_over_sheet_range_row_range() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "B1", 2.0).unwrap();
    engine.set_cell_value("Sheet3", "C1", 3.0).unwrap();
    engine.set_cell_value("Sheet3", "A2", 100.0).unwrap();

    engine
        .set_cell_formula("Summary", "A1", "=SUM(Sheet1:Sheet3!1:1)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Summary", "A1"), Value::Number(6.0));
}

#[test]
fn sum_full_sheet_range_over_sheet_range_is_sparse_and_marks_dirty() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "B2", 2.0).unwrap();
    engine.set_cell_value("Sheet3", "XFD1048576", 3.0).unwrap();

    engine
        // Place the formula on a different sheet so we don't create a circular reference: `A:XFD`
        // covers the entire sheet.
        .set_cell_formula("Summary", "C1", "=SUM(Sheet1:Sheet3!A:XFD)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Summary", "C1"), Value::Number(6.0));

    // Updating a cell that was previously blank should still dirty the formula cell even though
    // the audit graph does not expand full-sheet ranges.
    engine.set_cell_value("Sheet2", "C3", 10.0).unwrap();
    assert!(engine.is_dirty("Summary", "C1"));
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Summary", "C1"), Value::Number(16.0));
}

#[test]
fn union_over_sheet_range_refs_is_ref_error() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "A1", 2.0).unwrap();
    engine.set_cell_value("Sheet3", "A1", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 10.0).unwrap();

    // Union inside a function argument must be parenthesized to avoid being parsed as multiple
    // arguments. Excel's union/intersection reference algebra does not allow combining references
    // that resolve to different sheets, so 3D spans cannot participate (they expand to multiple
    // per-sheet areas).
    engine
        .set_cell_formula("Summary", "A1", "=SUM((Sheet1:Sheet3!A1,Sheet1!A2))")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Summary", "A1"),
        Value::Error(formula_engine::ErrorKind::Ref)
    );
}

#[test]
fn bytecode_union_operator_uses_current_sheet_for_unqualified_refs() {
    // Regression test: the bytecode VM must treat unqualified reference unions like `(A1,B1)` as
    // referring to the *current sheet* (not always Sheet1 / sheet_id=0).
    //
    // This relies on the VM passing the current sheet id through to the reference-union operator
    // so it can tag the resulting `MultiRange` with the correct sheet.
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", 2.0).unwrap();
    engine.set_cell_value("Sheet2", "A1", 10.0).unwrap();
    engine.set_cell_value("Sheet2", "B1", 20.0).unwrap();

    // Place the formula on Sheet2 so A1/B1 should resolve to the Sheet2 values.
    engine
        .set_cell_formula("Sheet2", "C1", "=SUM((A1,B1))")
        .unwrap();

    // Ensure the bytecode backend accepted the formula.
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    let bytecode_value = engine.get_cell_value("Sheet2", "C1");

    engine.set_bytecode_enabled(false);
    engine.recalculate_single_threaded();
    let ast_value = engine.get_cell_value("Sheet2", "C1");

    assert_eq!(bytecode_value, ast_value);
    assert_eq!(bytecode_value, Value::Number(30.0));
}

#[test]
fn evaluates_sum_over_sheet_range_ref_and_additional_argument() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "A1", 2.0).unwrap();
    engine.set_cell_value("Sheet3", "A1", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 10.0).unwrap();

    // Use separate function arguments rather than the reference union operator.
    engine
        .set_cell_formula("Summary", "A1", "=SUM(Sheet1:Sheet3!A1,Sheet1!A2)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Summary", "A1"), Value::Number(16.0));
}

#[test]
fn recalculates_when_intermediate_sheet_changes() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "A1", 2.0).unwrap();
    engine.set_cell_value("Sheet3", "A1", 3.0).unwrap();

    engine
        .set_cell_formula("Summary", "A1", "=SUM(Sheet1:Sheet3!A1)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Summary", "A1"), Value::Number(6.0));

    engine.set_cell_value("Sheet2", "A1", 5.0).unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Summary", "A1"), Value::Number(9.0));
}

#[test]
fn evaluates_sum_over_reversed_sheet_range_ref() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "A1", 2.0).unwrap();
    engine.set_cell_value("Sheet3", "A1", 3.0).unwrap();

    // Excel resolves 3D spans by workbook sheet order regardless of whether the
    // user writes them forward or reversed.
    engine
        .set_cell_formula("Summary", "A1", "=SUM(Sheet3:Sheet1!A1)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Summary", "A1"), Value::Number(6.0));
}

#[test]
fn bytecode_compiles_sum_over_sheet_range_cell_ref_and_matches_ast() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "A1", 2.0).unwrap();
    engine.set_cell_value("Sheet3", "A1", 3.0).unwrap();

    // Place the formula on Sheet1 so the 3D span includes the current sheet.
    engine
        .set_cell_formula("Sheet1", "B1", "=SUM(Sheet1:Sheet3!A1)")
        .unwrap();

    // Ensure the bytecode backend accepted the 3D reference.
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    let bytecode_value = engine.get_cell_value("Sheet1", "B1");

    engine.set_bytecode_enabled(false);
    engine.recalculate_single_threaded();
    let ast_value = engine.get_cell_value("Sheet1", "B1");

    assert_eq!(bytecode_value, ast_value);
    assert_eq!(bytecode_value, Value::Number(6.0));
}

#[test]
fn bytecode_compiles_sum_over_sheet_range_when_formula_on_other_sheet_and_matches_ast() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "A1", 2.0).unwrap();
    engine.set_cell_value("Sheet3", "A1", 3.0).unwrap();

    // Place the formula on a sheet outside the span so bytecode evaluation must read values from
    // other sheets explicitly.
    engine
        .set_cell_formula("Summary", "A1", "=SUM(Sheet1:Sheet3!A1)")
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    let bytecode_value = engine.get_cell_value("Summary", "A1");

    engine.set_bytecode_enabled(false);
    engine.recalculate_single_threaded();
    let ast_value = engine.get_cell_value("Summary", "A1");

    assert_eq!(bytecode_value, ast_value);
    assert_eq!(bytecode_value, Value::Number(6.0));
}

#[test]
fn bytecode_compiles_count_over_sheet_range_area_ref_and_matches_ast() {
    let mut engine = Engine::new();
    for (sheet, values) in [
        ("Sheet1", [1.0, 2.0, 3.0]),
        ("Sheet2", [4.0, 5.0, 6.0]),
        ("Sheet3", [7.0, 8.0, 9.0]),
    ] {
        for (idx, v) in values.into_iter().enumerate() {
            engine
                .set_cell_value(sheet, &format!("A{}", idx + 1), v)
                .unwrap();
        }
    }

    engine
        .set_cell_formula("Sheet1", "B1", "=COUNT(Sheet1:Sheet3!A1:A3)")
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    let bytecode_value = engine.get_cell_value("Sheet1", "B1");

    engine.set_bytecode_enabled(false);
    engine.recalculate_single_threaded();
    let ast_value = engine.get_cell_value("Sheet1", "B1");

    assert_eq!(bytecode_value, ast_value);
    assert_eq!(bytecode_value, Value::Number(9.0));
}

#[test]
fn bytecode_dynamic_deref_sheet_range_ref_matches_ast() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "A1", 2.0).unwrap();
    engine.set_cell_value("Sheet3", "A1", 3.0).unwrap();

    // A 3D reference used as a formula result produces a multi-area reference union, which cannot
    // be spilled as a single rectangular array. The engine surfaces this as #VALUE!.
    engine
        .set_cell_formula("Summary", "A1", "=Sheet1:Sheet3!A1")
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    let bytecode_value = engine.get_cell_value("Summary", "A1");

    engine.set_bytecode_enabled(false);
    engine.recalculate_single_threaded();
    let ast_value = engine.get_cell_value("Summary", "A1");

    assert_eq!(bytecode_value, ast_value);
    assert_eq!(
        bytecode_value,
        Value::Error(formula_engine::ErrorKind::Value)
    );
}

#[test]
fn bytecode_compiles_counta_and_countblank_over_sheet_range_area_ref_and_matches_ast() {
    let mut engine = Engine::new();

    // 3 sheets x 3 rows = 9 cells total.
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    // Sheet1!A2 left blank
    engine.set_cell_value("Sheet1", "A3", "").unwrap(); // empty string

    engine.set_cell_value("Sheet2", "A1", "hello").unwrap();
    engine.set_cell_value("Sheet2", "A2", Value::Blank).unwrap();
    engine
        .set_cell_value("Sheet2", "A3", Value::Error(ErrorKind::Div0))
        .unwrap();

    engine.set_cell_value("Sheet3", "A1", true).unwrap();
    engine.set_cell_value("Sheet3", "A2", 2.0).unwrap();
    // Sheet3!A3 left blank

    engine
        .set_cell_formula("Sheet1", "B1", "=COUNTA(Sheet1:Sheet3!A1:A3)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B2", "=COUNTBLANK(Sheet1:Sheet3!A1:A3)")
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 2);

    engine.recalculate_single_threaded();
    let bytecode_counta = engine.get_cell_value("Sheet1", "B1");
    let bytecode_countblank = engine.get_cell_value("Sheet1", "B2");

    engine.set_bytecode_enabled(false);
    engine.recalculate_single_threaded();
    let ast_counta = engine.get_cell_value("Sheet1", "B1");
    let ast_countblank = engine.get_cell_value("Sheet1", "B2");

    assert_eq!(bytecode_counta, ast_counta);
    assert_eq!(bytecode_countblank, ast_countblank);

    assert_eq!(bytecode_counta, Value::Number(6.0));
    assert_eq!(bytecode_countblank, Value::Number(4.0));
}

#[test]
fn bytecode_compiles_and_or_over_sheet_range_cell_ref_and_matches_ast() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", "hello").unwrap();
    engine.set_cell_value("Sheet2", "A1", 0.0).unwrap();
    engine.set_cell_value("Sheet3", "A1", 1.0).unwrap();

    // Place the formula on Sheet1 so the 3D span includes the current sheet.
    engine
        .set_cell_formula("Sheet1", "B1", "=AND(Sheet1:Sheet3!A1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B2", "=OR(Sheet1:Sheet3!A1)")
        .unwrap();

    // Ensure the bytecode backend accepted the 3D reference.
    assert_eq!(engine.bytecode_program_count(), 2);

    engine.recalculate_single_threaded();
    let bytecode_and = engine.get_cell_value("Sheet1", "B1");
    let bytecode_or = engine.get_cell_value("Sheet1", "B2");

    engine.set_bytecode_enabled(false);
    engine.recalculate_single_threaded();
    let ast_and = engine.get_cell_value("Sheet1", "B1");
    let ast_or = engine.get_cell_value("Sheet1", "B2");

    assert_eq!(bytecode_and, ast_and);
    assert_eq!(bytecode_or, ast_or);

    // `Sheet1!A1` is text and should be ignored for the 3D reference union; `0` forces AND false,
    // and `1` forces OR true.
    assert_eq!(bytecode_and, Value::Bool(false));
    assert_eq!(bytecode_or, Value::Bool(true));
}

#[test]
fn bytecode_compiles_counta_and_countblank_over_large_sheet_range_span() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "B2", "hello").unwrap();
    engine.set_cell_value("Sheet3", "F262145", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "C3", "").unwrap(); // empty string

    engine
        .set_cell_formula("Summary", "A1", "=COUNTA(Sheet1:Sheet3!A1:F262145)")
        .unwrap();
    engine
        .set_cell_formula("Summary", "A2", "=COUNTBLANK(Sheet1:Sheet3!A1:F262145)")
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 2);

    engine.recalculate_single_threaded();

    // Non-empty cells include empty strings for COUNTA.
    assert_eq!(engine.get_cell_value("Summary", "A1"), Value::Number(4.0));
    // 3 sheets x 262,145 rows x 6 cols = 4,718,610 cells total.
    // Only 3 of them are non-blank (empty strings count as blank for COUNTBLANK).
    assert_eq!(
        engine.get_cell_value("Summary", "A2"),
        Value::Number(4_718_607.0)
    );
}

#[test]
fn bytecode_compiles_countif_over_sheet_range_area_ref_and_matches_ast() {
    let mut engine = Engine::new();

    for sheet in ["Sheet1", "Sheet2", "Sheet3"] {
        engine.set_cell_value(sheet, "A1", "x").unwrap();
        engine.set_cell_value(sheet, "A2", 0.0).unwrap();
        // sheet!A3 left blank.
    }

    engine
        .set_cell_formula("Summary", "A1", "=COUNTIF(Sheet1:Sheet3!A1:A3,0)")
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    let bytecode_value = engine.get_cell_value("Summary", "A1");

    engine.set_bytecode_enabled(false);
    engine.recalculate_single_threaded();
    let ast_value = engine.get_cell_value("Summary", "A1");

    assert_eq!(bytecode_value, ast_value);
    // `A1` on each sheet is text and should not be coerced to 0 for numeric COUNTIF criteria.
    // Blanks match `0`, so the total is: 3 sheets x (0 + blank) = 6.
    assert_eq!(bytecode_value, Value::Number(6.0));
}
