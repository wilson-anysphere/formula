use formula_engine::{eval::parse_a1, BytecodeCompileReason, Engine, Value};

#[test]
fn bytecode_large_sparse_range() {
    let mut engine = Engine::new();
    engine
        .set_sheet_dimensions("Sheet1", 6_000_000, 10)
        .expect("set sheet dimensions");

    engine.set_cell_value("Sheet1", "A1", 1.0).expect("set A1");
    engine
        .set_cell_value("Sheet1", "A6000000", 2.0)
        .expect("set A6000000");

    // Use a huge-but-sparse range. The bytecode backend should be able to compile this even
    // though the resolved range spans more than the historical `BYTECODE_MAX_RANGE_CELLS` limit.
    engine
        .set_cell_formula("Sheet1", "B1", "=SUM(A:A)")
        .expect("set formula");

    let report = engine.bytecode_compile_report(usize::MAX);
    let b1 = parse_a1("B1").expect("parse B1");

    assert!(
        !report.iter().any(|e| {
            e.sheet == "Sheet1"
                && e.addr == b1
                && matches!(e.reason, BytecodeCompileReason::ExceedsRangeCellLimit)
        }),
        "unexpected ExceedsRangeCellLimit for Sheet1!B1: {report:?}"
    );
    assert!(
        report
            .iter()
            .find(|e| e.sheet == "Sheet1" && e.addr == b1)
            .is_none(),
        "expected Sheet1!B1 to compile to bytecode; report: {report:?}"
    );
    assert!(
        engine.bytecode_program_count() > 0,
        "expected at least one bytecode program to be compiled"
    );

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(3.0));
}
