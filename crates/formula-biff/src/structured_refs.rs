use std::borrow::Cow;

use formula_model::external_refs::escape_bracketed_identifier_content;

pub const FLAG_ALL: u16 = 0x0001;
pub const FLAG_HEADERS: u16 = 0x0002;
pub const FLAG_DATA: u16 = 0x0004;
pub const FLAG_TOTALS: u16 = 0x0008;
pub const FLAG_THIS_ROW: u16 = 0x0010;
pub const KNOWN_FLAGS_MASK: u16 = FLAG_ALL | FLAG_HEADERS | FLAG_DATA | FLAG_TOTALS | FLAG_THIS_ROW;

/// Structured reference item selector (Excel table "special items").
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum StructuredRefItem {
    All,
    Data,
    Headers,
    Totals,
    ThisRow,
}

/// Structured reference column selector.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum StructuredColumns {
    /// No explicit column selector (all columns).
    All,
    /// Single column selector (e.g. `[Col]`).
    Single(String),
    /// Column range selector (e.g. `[[Col1]:[Col2]]`).
    Range { start: String, end: String },
}

pub fn structured_ref_item_from_flags(flags: u16) -> Option<StructuredRefItem> {
    // Flags are not strictly documented as mutually exclusive. Prefer the same priority order as
    // the XLSB decoder: this-row beats header/totals/all/data.
    if flags & FLAG_THIS_ROW != 0 {
        Some(StructuredRefItem::ThisRow)
    } else if flags & FLAG_HEADERS != 0 {
        Some(StructuredRefItem::Headers)
    } else if flags & FLAG_TOTALS != 0 {
        Some(StructuredRefItem::Totals)
    } else if flags & FLAG_ALL != 0 {
        Some(StructuredRefItem::All)
    } else if flags & FLAG_DATA != 0 {
        Some(StructuredRefItem::Data)
    } else {
        None
    }
}

pub fn structured_columns_placeholder_from_ids(
    col_first: u32,
    col_last: u32,
) -> StructuredColumns {
    if col_first == 0 && col_last == 0 {
        StructuredColumns::All
    } else if col_first == col_last {
        StructuredColumns::Single(format!("Column{col_first}"))
    } else {
        StructuredColumns::Range {
            start: format!("Column{col_first}"),
            end: format!("Column{col_last}"),
        }
    }
}

pub fn structured_ref_is_single_cell(
    item: Option<StructuredRefItem>,
    columns: &StructuredColumns,
) -> bool {
    match (item, columns) {
        (Some(StructuredRefItem::ThisRow), StructuredColumns::Single(_)) => true,
        // `Table1[[#Headers],[Col]]` and `Table1[[#Totals],[Col]]` resolve to a single cell.
        (
            Some(StructuredRefItem::Headers | StructuredRefItem::Totals),
            StructuredColumns::Single(_),
        ) => true,
        _ => false,
    }
}

pub fn structured_ref_item_literal(item: StructuredRefItem) -> &'static str {
    match item {
        StructuredRefItem::All => "#All",
        StructuredRefItem::Data => "#Data",
        StructuredRefItem::Headers => "#Headers",
        StructuredRefItem::Totals => "#Totals",
        StructuredRefItem::ThisRow => "#This Row",
    }
}

fn escape_structured_ref_bracket_content(raw: &str) -> Cow<'_, str> {
    // Excel escapes `]` as `]]` within structured reference bracketed identifiers.
    escape_bracketed_identifier_content(raw)
}

pub fn format_structured_ref(
    table_name: Option<&str>,
    item: Option<StructuredRefItem>,
    columns: &StructuredColumns,
) -> String {
    // This-row shorthand: `[@Col]`, `[@]`, and `[@[Col1]:[Col2]]`.
    if matches!(item, Some(StructuredRefItem::ThisRow)) {
        match columns {
            StructuredColumns::Single(col) => {
                let col = escape_structured_ref_bracket_content(col);
                return format!("[@{col}]");
            }
            StructuredColumns::All => return "[@]".to_string(),
            StructuredColumns::Range { start, end } => {
                let start = escape_structured_ref_bracket_content(start);
                let end = escape_structured_ref_bracket_content(end);
                return format!("[@[{start}]:[{end}]]");
            }
        }
    }

    let table = table_name.unwrap_or("");

    // Item-only selections: `Table1[#All]`, `Table1[#Headers]`, etc.
    if columns == &StructuredColumns::All {
        return match item {
            Some(item) => format!("{table}[{}]", structured_ref_item_literal(item)),
            // Default row selector with no column selection: treat as `[#Data]`.
            None => format!("{table}[#Data]"),
        };
    }

    // Single-column selection with default/data item: `Table1[Col]` or `Table1[[Col1]:[Col2]]`.
    if matches!(item, None | Some(StructuredRefItem::Data)) {
        return match columns {
            StructuredColumns::Single(col) => {
                let col = escape_structured_ref_bracket_content(col);
                format!("{table}[{col}]")
            }
            StructuredColumns::Range { start, end } => {
                let start = escape_structured_ref_bracket_content(start);
                let end = escape_structured_ref_bracket_content(end);
                format!("{table}[[{start}]:[{end}]]")
            }
            StructuredColumns::All => {
                // Covered by the `columns == All` early return above.
                format!("{table}[#Data]")
            }
        };
    }

    // General nested form: `Table1[[#Headers],[Col]]` or `Table1[[#Headers],[Col1]:[Col2]]`.
    let item = item.unwrap_or(StructuredRefItem::Data);
    match columns {
        StructuredColumns::Single(col) => {
            let col = escape_structured_ref_bracket_content(col);
            format!("{table}[[{}],[{col}]]", structured_ref_item_literal(item))
        }
        StructuredColumns::Range { start, end } => {
            let start = escape_structured_ref_bracket_content(start);
            let end = escape_structured_ref_bracket_content(end);
            format!(
                "{table}[[{}],[{start}]:[{end}]]",
                structured_ref_item_literal(item)
            )
        }
        StructuredColumns::All => {
            // Covered by the `columns == All` early return above.
            format!("{table}[{}]", structured_ref_item_literal(item))
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn format_structured_ref_defaults_to_data_item_when_unqualified() {
        assert_eq!(
            format_structured_ref(Some("Table1"), None, &StructuredColumns::All),
            "Table1[#Data]"
        );
    }

    #[test]
    fn format_structured_ref_emits_item_only_form() {
        assert_eq!(
            format_structured_ref(
                Some("Table1"),
                Some(StructuredRefItem::Headers),
                &StructuredColumns::All
            ),
            "Table1[#Headers]"
        );
    }

    #[test]
    fn format_structured_ref_escapes_close_brackets_in_column_names() {
        assert_eq!(
            format_structured_ref(
                Some("Table1"),
                None,
                &StructuredColumns::Single("A]B".to_string())
            ),
            "Table1[A]]B]"
        );
        assert_eq!(
            format_structured_ref(
                None,
                Some(StructuredRefItem::ThisRow),
                &StructuredColumns::Single("A]B".to_string())
            ),
            "[@A]]B]"
        );
    }

    #[test]
    fn format_structured_ref_emits_nested_form_for_items_with_columns() {
        assert_eq!(
            format_structured_ref(
                Some("Table1"),
                Some(StructuredRefItem::Totals),
                &StructuredColumns::Single("Qty".to_string())
            ),
            "Table1[[#Totals],[Qty]]"
        );
    }

    #[test]
    fn structured_ref_is_single_cell_matches_expected_cases() {
        assert!(structured_ref_is_single_cell(
            Some(StructuredRefItem::ThisRow),
            &StructuredColumns::Single("Qty".to_string())
        ));
        assert!(structured_ref_is_single_cell(
            Some(StructuredRefItem::Headers),
            &StructuredColumns::Single("Qty".to_string())
        ));
        assert!(!structured_ref_is_single_cell(
            Some(StructuredRefItem::Headers),
            &StructuredColumns::All
        ));
    }

    #[test]
    fn structured_ref_item_from_flags_prefers_this_row_over_other_items() {
        assert_eq!(
            structured_ref_item_from_flags(FLAG_THIS_ROW | FLAG_HEADERS),
            Some(StructuredRefItem::ThisRow)
        );
    }

    #[test]
    fn structured_columns_placeholder_from_ids_formats_expected_names() {
        assert_eq!(
            structured_columns_placeholder_from_ids(0, 0),
            StructuredColumns::All
        );
        assert_eq!(
            structured_columns_placeholder_from_ids(2, 2),
            StructuredColumns::Single("Column2".to_string())
        );
        assert_eq!(
            structured_columns_placeholder_from_ids(2, 3),
            StructuredColumns::Range {
                start: "Column2".to_string(),
                end: "Column3".to_string()
            }
        );
    }
}

