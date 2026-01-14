use formula_model::Worksheet;

#[test]
fn grouping_respects_existing_collapsed_summary_row() {
    let mut sheet = Worksheet::new(1, "Sheet1");

    // Collapse the (future) group first, then apply grouping. The grouping operation should
    // recompute outline-hidden rows so the collapsed summary immediately takes effect.
    sheet.toggle_row_group(5);
    sheet.group_rows(2, 4);

    assert!(!sheet.is_row_hidden_effective(1));
    for row in 2..=4 {
        assert!(sheet.row_outline_entry(row).hidden.outline);
        assert!(sheet.is_row_hidden_effective(row));
    }
    // The summary row itself stays visible (Excel semantics).
    assert!(!sheet.is_row_hidden_effective(5));
}

#[test]
fn user_hidden_and_outline_hidden_both_contribute_to_effective_hidden() {
    let mut sheet = Worksheet::new(1, "Sheet1");

    sheet.group_rows(2, 4);
    sheet.toggle_row_group(5); // collapse

    // Hide row 3 (0-based: 2) explicitly.
    sheet.set_row_hidden(2, true);

    let entry = sheet.row_outline_entry(3);
    assert!(entry.hidden.outline);
    assert!(entry.hidden.user);
    assert!(sheet.is_row_hidden_effective(3));

    // Expanding the group clears outline-hidden, but preserves user-hidden.
    sheet.toggle_row_group(5); // expand
    assert!(!sheet.row_outline_entry(2).hidden.outline);
    assert!(!sheet.is_row_hidden_effective(2));

    assert!(sheet.row_outline_entry(3).hidden.user);
    assert!(sheet.is_row_hidden_effective(3));
}

#[test]
fn filter_hidden_can_be_cleared_without_affecting_user_hidden() {
    let mut sheet = Worksheet::new(1, "Sheet1");

    // Row 3 is user-hidden.
    sheet.set_row_hidden(2, true);
    assert!(sheet.row_outline_entry(3).hidden.user);

    // Then it's also hidden by filter.
    sheet.set_filter_hidden_row(3, true);
    let entry = sheet.row_outline_entry(3);
    assert!(entry.hidden.user);
    assert!(entry.hidden.filter);

    // Clearing filter-hidden should keep user-hidden.
    sheet.clear_filter_hidden_range(1, 10);
    let entry = sheet.row_outline_entry(3);
    assert!(entry.hidden.user);
    assert!(!entry.hidden.filter);
}

#[test]
fn missing_outline_in_json_payload_still_deserializes_and_preserves_user_hidden() {
    let mut sheet = Worksheet::new(1, "Sheet1");
    sheet.set_row_hidden(2, true); // row 3 user-hidden

    let mut json = serde_json::to_value(&sheet).unwrap();
    json.as_object_mut().unwrap().remove("outline");

    let deserialized: Worksheet = serde_json::from_value(json).unwrap();
    assert!(deserialized.is_row_hidden_effective(3));
    assert!(deserialized.row_outline_entry(3).hidden.user);
    assert!(deserialized.outline.rows.entry(3).hidden.user);
}
