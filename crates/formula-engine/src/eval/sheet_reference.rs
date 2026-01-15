use super::ast::SheetReference;
use crate::SheetRef;
use formula_model::sheet_name_eq_case_insensitive;

pub(crate) fn lower_sheet_reference(
    workbook: &Option<String>,
    sheet: &Option<SheetRef>,
) -> SheetReference<String> {
    match (workbook.as_ref(), sheet.as_ref()) {
        (Some(book), Some(sheet_ref)) => match sheet_ref {
            SheetRef::Sheet(sheet) => {
                SheetReference::External(crate::external_refs::format_external_key(book, sheet))
            }
            SheetRef::SheetRange { start, end } => {
                if sheet_name_eq_case_insensitive(start, end) {
                    SheetReference::External(crate::external_refs::format_external_key(book, start))
                } else {
                    SheetReference::External(crate::external_refs::format_external_span_key(
                        book, start, end,
                    ))
                }
            }
        },
        (Some(book), None) => {
            SheetReference::External(crate::external_refs::format_external_workbook_key(book))
        }
        (None, Some(sheet_ref)) => match sheet_ref {
            SheetRef::Sheet(sheet) => SheetReference::Sheet(sheet.clone()),
            SheetRef::SheetRange { start, end } if sheet_name_eq_case_insensitive(start, end) => {
                SheetReference::Sheet(start.clone())
            }
            SheetRef::SheetRange { start, end } => SheetReference::SheetRange(start.clone(), end.clone()),
        },
        (None, None) => SheetReference::Current,
    }
}

