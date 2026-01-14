use crate::functions::harness::{assert_number, TestSheet};
use formula_engine::value::{ErrorKind, Value};

fn seed_database(sheet: &mut TestSheet) {
    // Database (A1:D5)
    sheet.set("A1", "Name");
    sheet.set("B1", "Dept");
    sheet.set("C1", "Age");
    sheet.set("D1", "Salary");

    sheet.set("A2", "Alice");
    sheet.set("B2", "Sales");
    sheet.set("C2", 30);
    sheet.set("D2", 1000);

    sheet.set("A3", "Bob");
    sheet.set("B3", "Sales");
    sheet.set("C3", 35);
    sheet.set("D3", 1500);

    sheet.set("A4", "Carol");
    sheet.set("B4", "HR");
    sheet.set("C4", 28);
    sheet.set("D4", "n/a"); // non-numeric salary to probe DCOUNT vs DCOUNTA vs DSUM

    sheet.set("A5", "Dan");
    sheet.set("B5", "HR");
    sheet.set("C5", 40);
    sheet.set("D5", 2000);
}

#[test]
fn database_functions_or_of_and_criteria() {
    let mut sheet = TestSheet::new();
    seed_database(&mut sheet);

    // Criteria (F1:G3):
    // (Dept="Sales" AND Age>30) OR (Dept="HR" AND Age<30)
    sheet.set("F1", "Dept");
    sheet.set("G1", "Age");
    sheet.set("F2", "Sales");
    sheet.set("G2", ">30");
    sheet.set("F3", "HR");
    sheet.set("G3", "<30");

    // Matches Bob (1500) and Carol ("n/a")
    assert_number(&sheet.eval("=DSUM(A1:D5,\"Salary\",F1:G3)"), 1500.0);
    assert_number(&sheet.eval("=DAVERAGE(A1:D5,\"Salary\",F1:G3)"), 1500.0);
    assert_number(&sheet.eval("=DMAX(A1:D5,\"Salary\",F1:G3)"), 1500.0);
    assert_number(&sheet.eval("=DMIN(A1:D5,\"Salary\",F1:G3)"), 1500.0);
    assert_number(&sheet.eval("=DPRODUCT(A1:D5,\"Salary\",F1:G3)"), 1500.0);
    assert_number(&sheet.eval("=DCOUNT(A1:D5,\"Salary\",F1:G3)"), 1.0);
    assert_number(&sheet.eval("=DCOUNTA(A1:D5,\"Salary\",F1:G3)"), 2.0);

    // Sample variance/stdev should error for a single numeric value.
    assert_eq!(
        sheet.eval("=DVAR(A1:D5,\"Salary\",F1:G3)"),
        Value::Error(ErrorKind::Div0)
    );
    assert_eq!(
        sheet.eval("=DSTDEV(A1:D5,\"Salary\",F1:G3)"),
        Value::Error(ErrorKind::Div0)
    );

    // Population variance/stdev for a single numeric value is 0.
    assert_number(&sheet.eval("=DVARP(A1:D5,\"Salary\",F1:G3)"), 0.0);
    assert_number(&sheet.eval("=DSTDEVP(A1:D5,\"Salary\",F1:G3)"), 0.0);
}

#[test]
fn dget_errors_and_success() {
    let mut sheet = TestSheet::new();
    seed_database(&mut sheet);

    // DGET with a single match.
    sheet.set("F1", "Name");
    sheet.set("F2", "Alice");
    assert_number(&sheet.eval("=DGET(A1:D5,\"Salary\",F1:F2)"), 1000.0);

    // Zero matches -> #VALUE!
    sheet.set("F2", "Nope");
    assert_eq!(
        sheet.eval("=DGET(A1:D5,\"Salary\",F1:F2)"),
        Value::Error(ErrorKind::Value)
    );

    // Multiple matches -> #NUM!
    sheet.set("F1", "Dept");
    sheet.set("F2", "Sales");
    assert_eq!(
        sheet.eval("=DGET(A1:D5,\"Salary\",F1:F2)"),
        Value::Error(ErrorKind::Num)
    );
}

#[test]
fn database_functions_computed_criteria_basic() {
    let mut sheet = TestSheet::new();
    seed_database(&mut sheet);

    // Criteria (F1:F2):
    // Blank header + a formula means "computed criteria".
    //
    // The formula is written against the first database record row (row 2) and evaluated as if it
    // were filled down over the database.
    sheet.set_formula("F2", "=C2>30");

    // Matches Bob (35) and Dan (40) => Salary sum = 1500 + 2000.
    assert_number(&sheet.eval("=DSUM(A1:D5,\"Salary\",F1:F2)"), 3500.0);
    assert_number(&sheet.eval("=DAVERAGE(A1:D5,\"Salary\",F1:F2)"), 1750.0);
    assert_number(&sheet.eval("=DCOUNT(A1:D5,\"Salary\",F1:F2)"), 2.0);
    assert_number(&sheet.eval("=DCOUNTA(A1:D5,\"Salary\",F1:F2)"), 2.0);
    assert_number(&sheet.eval("=DMAX(A1:D5,\"Salary\",F1:F2)"), 2000.0);
    assert_number(&sheet.eval("=DMIN(A1:D5,\"Salary\",F1:F2)"), 1500.0);
    assert_number(
        &sheet.eval("=DPRODUCT(A1:D5,\"Salary\",F1:F2)"),
        3_000_000.0,
    );

    // Variance/stdev over two numeric values.
    assert_number(&sheet.eval("=DVAR(A1:D5,\"Salary\",F1:F2)"), 125_000.0);
    assert_number(
        &sheet.eval("=DSTDEV(A1:D5,\"Salary\",F1:F2)"),
        (125_000.0_f64).sqrt(),
    );
    assert_number(&sheet.eval("=DVARP(A1:D5,\"Salary\",F1:F2)"), 62_500.0);
    assert_number(
        &sheet.eval("=DSTDEVP(A1:D5,\"Salary\",F1:F2)"),
        (62_500.0_f64).sqrt(),
    );

    // Single-match computed criteria.
    sheet.set_formula("F2", "=C2>35");
    assert_number(&sheet.eval("=DGET(A1:D5,\"Salary\",F1:F2)"), 2000.0);
}

#[test]
fn database_functions_computed_criteria_with_nonmatching_header() {
    let mut sheet = TestSheet::new();
    seed_database(&mut sheet);

    // Excel also allows computed criteria when the header is any label that does not match a
    // database field name.
    sheet.set("F1", "Criteria");
    sheet.set_formula("F2", "=C2>30");

    // Matches Bob (35) and Dan (40) => Salary sum = 1500 + 2000.
    assert_number(&sheet.eval("=DSUM(A1:D5,\"Salary\",F1:F2)"), 3500.0);
}

#[test]
fn database_functions_computed_criteria_or_with_standard_criteria() {
    let mut sheet = TestSheet::new();
    seed_database(&mut sheet);

    // Criteria (F1:H3):
    // (Dept="Sales" AND computed Age>32) OR (Dept="HR" AND Age<30)
    sheet.set("F1", "Dept");
    // G1 is intentionally blank to enable computed criteria.
    sheet.set("H1", "Age");

    sheet.set("F2", "Sales");
    sheet.set_formula("G2", "=C2>32");

    sheet.set("F3", "HR");
    sheet.set("H3", "<30");

    // Matches Bob (1500) and Carol ("n/a")
    assert_number(&sheet.eval("=DSUM(A1:D5,\"Salary\",F1:H3)"), 1500.0);
    assert_number(&sheet.eval("=DCOUNT(A1:D5,\"Salary\",F1:H3)"), 1.0);
    assert_number(&sheet.eval("=DCOUNTA(A1:D5,\"Salary\",F1:H3)"), 2.0);
}

#[test]
fn database_functions_computed_criteria_error_propagation() {
    let mut sheet = TestSheet::new();
    seed_database(&mut sheet);

    // Criteria (F1:F2): formula errors on the "Bob" record (Age==35).
    sheet.set_formula("F2", "=1/(C2-35)>0");

    assert_eq!(
        sheet.eval("=DSUM(A1:D5,\"Salary\",F1:F2)"),
        Value::Error(ErrorKind::Div0)
    );
    assert_eq!(
        sheet.eval("=DGET(A1:D5,\"Salary\",F1:F2)"),
        Value::Error(ErrorKind::Div0)
    );
}

#[test]
fn database_functions_computed_criteria_requires_formula() {
    let mut sheet = TestSheet::new();
    seed_database(&mut sheet);

    // Criteria header blank + non-blank cell that is not a formula => #VALUE!
    sheet.set("F2", ">30");
    assert_eq!(
        sheet.eval("=DSUM(A1:D5,\"Salary\",F1:F2)"),
        Value::Error(ErrorKind::Value)
    );
}

#[test]
fn database_functions_computed_criteria_multiple_computed_columns() {
    let mut sheet = TestSheet::new();
    seed_database(&mut sheet);

    // Criteria (F1:G2): both headers blank -> both columns are computed criteria.
    //
    // Age>30 AND Dept="HR" matches only Dan (salary 2000).
    sheet.set_formula("F2", "=C2>30");
    sheet.set_formula("G2", "=B2=\"HR\"");

    assert_number(&sheet.eval("=DSUM(A1:D5,\"Salary\",F1:G2)"), 2000.0);
    assert_number(&sheet.eval("=DGET(A1:D5,\"Salary\",F1:G2)"), 2000.0);
}

#[test]
fn database_functions_computed_criteria_or_across_computed_rows() {
    let mut sheet = TestSheet::new();
    seed_database(&mut sheet);

    // Criteria (F1:F3): blank header -> computed criteria.
    //
    // Age>35 OR Age<30 matches Dan (2000) and Carol ("n/a").
    sheet.set_formula("F2", "=C2>35");
    sheet.set_formula("F3", "=C2<30");

    assert_number(&sheet.eval("=DSUM(A1:D5,\"Salary\",F1:F3)"), 2000.0);
    assert_number(&sheet.eval("=DCOUNT(A1:D5,\"Salary\",F1:F3)"), 1.0);
    assert_number(&sheet.eval("=DCOUNTA(A1:D5,\"Salary\",F1:F3)"), 2.0);
}

#[test]
fn database_functions_computed_criteria_with_offset_database_range() {
    let mut sheet = TestSheet::new();

    // Database (B3:E7)
    sheet.set("B3", "Name");
    sheet.set("C3", "Dept");
    sheet.set("D3", "Age");
    sheet.set("E3", "Salary");

    sheet.set("B4", "Alice");
    sheet.set("C4", "Sales");
    sheet.set("D4", 30);
    sheet.set("E4", 1000);

    sheet.set("B5", "Bob");
    sheet.set("C5", "Sales");
    sheet.set("D5", 35);
    sheet.set("E5", 1500);

    sheet.set("B6", "Carol");
    sheet.set("C6", "HR");
    sheet.set("D6", 28);
    sheet.set("E6", "n/a");

    sheet.set("B7", "Dan");
    sheet.set("C7", "HR");
    sheet.set("D7", 40);
    sheet.set("E7", 2000);

    // Criteria (H1:H2): computed criteria referencing the first record row of the database (row 4).
    sheet.set_formula("H2", "=D4>30");

    // Matches Bob (35) and Dan (40) => Salary sum = 1500 + 2000.
    assert_number(&sheet.eval("=DSUM(B3:E7,\"Salary\",H1:H2)"), 3500.0);
    assert_number(&sheet.eval("=DSUM(B3:E7,4,H1:H2)"), 3500.0);
}

#[test]
fn database_functions_computed_criteria_respects_absolute_references() {
    let mut sheet = TestSheet::new();
    seed_database(&mut sheet);

    // Absolute row reference: C$2 is pinned to the first record row (Alice, age 30).
    // Therefore the computed criteria is always FALSE (30 > 30 == FALSE) for every record row.
    sheet.set_formula("F2", "=C$2>30");

    assert_number(&sheet.eval("=DSUM(A1:D5,\"Salary\",F1:F2)"), 0.0);
    assert_number(&sheet.eval("=DCOUNT(A1:D5,\"Salary\",F1:F2)"), 0.0);
    assert_eq!(
        sheet.eval("=DGET(A1:D5,\"Salary\",F1:F2)"),
        Value::Error(ErrorKind::Value)
    );
}

#[test]
fn database_functions_computed_criteria_coerces_numeric_results_to_bool() {
    let mut sheet = TestSheet::new();
    seed_database(&mut sheet);

    // Computed-criteria formulas don't have to literally return TRUE/FALSE; Excel boolean coercion
    // treats non-zero numbers as TRUE.
    sheet.set_formula("F2", "=--(C2>30)");

    // Matches Bob (35) and Dan (40) => Salary sum = 1500 + 2000.
    assert_number(&sheet.eval("=DSUM(A1:D5,\"Salary\",F1:F2)"), 3500.0);

    // Single-match computed criteria (numeric TRUE/FALSE).
    sheet.set_formula("F2", "=--(C2>35)");
    assert_number(&sheet.eval("=DGET(A1:D5,\"Salary\",F1:F2)"), 2000.0);
}

#[test]
fn database_functions_computed_criteria_blank_row_matches_all_records() {
    let mut sheet = TestSheet::new();
    seed_database(&mut sheet);

    // Criteria (F1:F2): blank header + blank criteria row should behave like "no criteria".
    //
    // DSUM ignores non-numeric values in the field ("n/a" for Carol).
    assert_number(&sheet.eval("=DSUM(A1:D5,\"Salary\",F1:F2)"), 4500.0);
    assert_number(&sheet.eval("=DAVERAGE(A1:D5,\"Salary\",F1:F2)"), 1500.0);
    assert_number(&sheet.eval("=DCOUNT(A1:D5,\"Salary\",F1:F2)"), 3.0);
    assert_number(&sheet.eval("=DCOUNTA(A1:D5,\"Salary\",F1:F2)"), 4.0);

    // DGET requires exactly one matching record; "match all" is #NUM!.
    assert_eq!(
        sheet.eval("=DGET(A1:D5,\"Salary\",F1:F2)"),
        Value::Error(ErrorKind::Num)
    );
}

#[test]
fn database_functions_computed_criteria_accepts_true_false_text_results() {
    let mut sheet = TestSheet::new();
    seed_database(&mut sheet);

    // Computed criteria formulas can return text. Excel boolean coercion treats "TRUE"/"FALSE"
    // (case-insensitive) as logical values.
    sheet.set_formula("F2", "=\"TRUE\"");

    assert_number(&sheet.eval("=DSUM(A1:D5,\"Salary\",F1:F2)"), 4500.0);
    assert_number(&sheet.eval("=DAVERAGE(A1:D5,\"Salary\",F1:F2)"), 1500.0);
    assert_number(&sheet.eval("=DCOUNT(A1:D5,\"Salary\",F1:F2)"), 3.0);
    assert_number(&sheet.eval("=DCOUNTA(A1:D5,\"Salary\",F1:F2)"), 4.0);
}

#[test]
fn database_functions_computed_criteria_falsey_empty_string_results_in_no_matches() {
    let mut sheet = TestSheet::new();
    seed_database(&mut sheet);

    // Empty string coerces to FALSE.
    sheet.set_formula("F2", "=\"\"");

    assert_number(&sheet.eval("=DSUM(A1:D5,\"Salary\",F1:F2)"), 0.0);
    assert_number(&sheet.eval("=DCOUNT(A1:D5,\"Salary\",F1:F2)"), 0.0);
    assert_number(&sheet.eval("=DCOUNTA(A1:D5,\"Salary\",F1:F2)"), 0.0);
    assert_eq!(
        sheet.eval("=DGET(A1:D5,\"Salary\",F1:F2)"),
        Value::Error(ErrorKind::Value)
    );
}

#[test]
fn database_functions_computed_criteria_can_reference_other_sheets() {
    let mut sheet = TestSheet::new();
    seed_database(&mut sheet);

    // Use a threshold from another sheet.
    sheet.set_on("Sheet2", "A1", 30);
    sheet.set_formula("F2", "=C2>Sheet2!A1");

    // Matches Bob (35) and Dan (40) => Salary sum = 1500 + 2000.
    assert_number(&sheet.eval("=DSUM(A1:D5,\"Salary\",F1:F2)"), 3500.0);
}
