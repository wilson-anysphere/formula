#![cfg(not(target_arch = "wasm32"))]

use formula_engine::Engine;

#[test]
fn bytecode_compile_report_orders_by_tab_order_after_sheet_reorder() {
    let mut engine = Engine::new();

    // Force AST fallback on multiple sheets. Unresolved UDF calls are treated as non-thread-safe,
    // preventing bytecode compilation.
    for sheet in ["Sheet1", "Sheet2", "Sheet3"] {
        engine
            .set_cell_formula(sheet, "A1", "=BYTECODE_COMPILE_REPORT_FALLBACK()")
            .unwrap();
    }

    let report = engine.bytecode_compile_report(10);
    assert_eq!(
        report.iter().map(|e| e.sheet.as_str()).collect::<Vec<_>>(),
        vec!["Sheet1", "Sheet2", "Sheet3"],
        "expected report entries to follow default tab order"
    );

    // Move Sheet3 to the front of the tab order. Sheet ids are stable, so sorting by numeric id
    // would keep Sheet1 first; sorting by tab index should reflect the reorder.
    assert!(engine.reorder_sheet("Sheet3", 0));

    let report = engine.bytecode_compile_report(10);
    assert_eq!(
        report.iter().map(|e| e.sheet.as_str()).collect::<Vec<_>>(),
        vec!["Sheet3", "Sheet1", "Sheet2"],
        "expected report entries to follow the updated tab order after reorder"
    );
}
