use std::cell::RefCell;
use std::collections::HashMap;

use formula_model::{
    parse_range_a1, CellIsOperator, CellRef, CellValue, CellValueProvider, CfRule, CfRuleKind,
    CfRuleSchema, CfStyleOverride, Worksheet,
};

fn sample_rule_and_dxfs() -> (Vec<CfRule>, Vec<CfStyleOverride>) {
    let dxfs = vec![CfStyleOverride {
        fill: Some(formula_model::Color::new_argb(0xFFFF0000)),
        font_color: None,
        bold: Some(true),
        italic: None,
    }];

    let rule = CfRule {
        schema: CfRuleSchema::Office2007,
        id: Some("1".to_string()),
        priority: 1,
        applies_to: vec![parse_range_a1("A1:A2").unwrap()],
        dxf_id: Some(0),
        stop_if_true: false,
        kind: CfRuleKind::CellIs {
            operator: CellIsOperator::GreaterThan,
            formulas: vec!["5".to_string()],
        },
        // Depend on the target cells themselves so edits invalidate the cache.
        dependencies: vec![parse_range_a1("A1:A2").unwrap()],
    };

    (vec![rule], dxfs)
}

#[test]
fn worksheet_stores_and_evaluates_conditional_formatting() {
    let mut sheet = Worksheet::new(1, "Sheet1");
    sheet.set_value(CellRef::from_a1("A1").unwrap(), CellValue::Number(10.0));
    sheet.set_value(CellRef::from_a1("A2").unwrap(), CellValue::Number(3.0));

    let (rules, dxfs) = sample_rule_and_dxfs();
    sheet.set_conditional_formatting(rules, dxfs);

    let visible = parse_range_a1("A1:A2").unwrap();
    let eval = sheet.evaluate_conditional_formatting(visible, &sheet, None);

    let a1 = eval.get(CellRef::from_a1("A1").unwrap()).unwrap();
    assert_eq!(
        a1.style.fill,
        Some(formula_model::Color::new_argb(0xFFFF0000))
    );
    assert_eq!(a1.style.bold, Some(true));

    let a2 = eval.get(CellRef::from_a1("A2").unwrap()).unwrap();
    assert_eq!(a2.style.fill, None);
    assert_eq!(a2.style.bold, None);
}

#[test]
fn worksheet_serialization_roundtrip_preserves_conditional_formatting() {
    let mut sheet = Worksheet::new(1, "Sheet1");
    let (rules, dxfs) = sample_rule_and_dxfs();
    sheet.set_conditional_formatting(rules.clone(), dxfs.clone());

    let json = serde_json::to_string(&sheet).unwrap();
    let decoded: Worksheet = serde_json::from_str(&json).unwrap();

    assert_eq!(decoded.conditional_formatting_rules, rules);
    assert_eq!(decoded.conditional_formatting_dxfs, dxfs);
}

struct MutableValues {
    values: RefCell<HashMap<CellRef, CellValue>>,
}

impl MutableValues {
    fn new(values: HashMap<CellRef, CellValue>) -> Self {
        Self {
            values: RefCell::new(values),
        }
    }

    fn set(&self, cell: CellRef, value: CellValue) {
        self.values.borrow_mut().insert(cell, value);
    }
}

impl CellValueProvider for MutableValues {
    fn get_value(&self, cell: CellRef) -> Option<CellValue> {
        self.values.borrow().get(&cell).cloned()
    }
}

#[test]
fn invalidation_clears_cached_evaluation_entries() {
    let mut sheet = Worksheet::new(1, "Sheet1");

    let (mut rules, dxfs) = sample_rule_and_dxfs();
    // Only depend on A1 for this test so we can invalidate a single cell.
    rules[0].applies_to = vec![parse_range_a1("A1").unwrap()];
    rules[0].dependencies = vec![parse_range_a1("A1").unwrap()];
    sheet.set_conditional_formatting(rules, dxfs);

    let values = MutableValues::new(HashMap::from([(
        CellRef::from_a1("A1").unwrap(),
        CellValue::Number(10.0),
    )]));

    let visible = parse_range_a1("A1").unwrap();

    let eval1 = sheet.evaluate_conditional_formatting(visible, &values, None);
    assert_eq!(
        eval1
            .get(CellRef::from_a1("A1").unwrap())
            .unwrap()
            .style
            .fill,
        Some(formula_model::Color::new_argb(0xFFFF0000))
    );

    // Mutate the value provider without telling the worksheet; cached result should remain.
    values.set(CellRef::from_a1("A1").unwrap(), CellValue::Number(0.0));
    let eval2 = sheet.evaluate_conditional_formatting(visible, &values, None);
    assert_eq!(
        eval2
            .get(CellRef::from_a1("A1").unwrap())
            .unwrap()
            .style
            .fill,
        Some(formula_model::Color::new_argb(0xFFFF0000))
    );

    // Now invalidate the cache entry and ensure the evaluation reflects the changed value.
    sheet.invalidate_conditional_formatting_cells([CellRef::from_a1("A1").unwrap()]);
    let eval3 = sheet.evaluate_conditional_formatting(visible, &values, None);
    assert_eq!(
        eval3
            .get(CellRef::from_a1("A1").unwrap())
            .unwrap()
            .style
            .fill,
        None
    );
}
