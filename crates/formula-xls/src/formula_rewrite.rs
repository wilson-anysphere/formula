use std::collections::HashSet;

use formula_model::{rewrite_sheet_names_in_formula, Workbook};

/// Rewrite all workbook cell formulas to account for worksheet renames during import.
///
/// `sheet_rename_pairs` contains `(old_name, new_name)` pairs, using the sheet names as they
/// appear in the source workbook (calamine metadata and/or BIFF BoundSheet names).
pub(crate) fn rewrite_workbook_formulas_for_sheet_renames(
    workbook: &mut Workbook,
    sheet_rename_pairs: &[(String, String)],
) {
    if sheet_rename_pairs.is_empty() {
        return;
    }

    // Rewriting multiple renames sequentially can cascade when one sheet's *new* name equals
    // another sheet's *old* name (case-insensitive), e.g.:
    //   A -> B
    //   B -> C
    //
    // If we naively apply `A -> B` then `B -> C`, references to A would end up pointing at C.
    //
    // Avoid this by rewriting through temporary unique names first:
    //   old -> tmp
    //   tmp -> new
    let mut reserved: HashSet<String> = HashSet::new();
    reserved.extend(
        sheet_rename_pairs
            .iter()
            .map(|(old, _)| super::normalize_sheet_name_for_match(old)),
    );
    reserved.extend(
        sheet_rename_pairs
            .iter()
            .map(|(_, new)| super::normalize_sheet_name_for_match(new)),
    );
    reserved.extend(
        workbook
            .sheets
            .iter()
            .map(|s| super::normalize_sheet_name_for_match(&s.name)),
    );

    let mut old_to_tmp: Vec<(String, String)> = Vec::with_capacity(sheet_rename_pairs.len());
    let mut tmp_to_new: Vec<(String, String)> = Vec::with_capacity(sheet_rename_pairs.len());

    let mut tmp_index: usize = 0;
    for (old, new) in sheet_rename_pairs {
        let tmp = loop {
            // Keep these short (and Excel-valid) even though they only exist transiently in formula
            // strings. Start with a letter/underscore so they never require quoting.
            let candidate = format!("_xls_import_tmp_{tmp_index}");
            tmp_index = tmp_index.saturating_add(1);
            let candidate_norm = super::normalize_sheet_name_for_match(&candidate);
            if reserved.insert(candidate_norm) {
                break candidate;
            }
        };

        old_to_tmp.push((old.clone(), tmp.clone()));
        tmp_to_new.push((tmp, new.clone()));
    }

    for sheet in &mut workbook.sheets {
        for (_, cell) in sheet.iter_cells_mut() {
            let Some(formula) = cell.formula.as_mut() else {
                continue;
            };

            // Avoid cloning the existing formula string.
            let mut rewritten = std::mem::take(formula);

            for (old, tmp) in &old_to_tmp {
                rewritten = rewrite_sheet_names_in_formula(&rewritten, old, tmp);
            }
            for (tmp, new) in &tmp_to_new {
                rewritten = rewrite_sheet_names_in_formula(&rewritten, tmp, new);
            }

            *formula = rewritten;
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use formula_model::{CellRef, Workbook};

    #[test]
    fn rewrites_workbook_formulas_for_sheet_renames() {
        let mut workbook = Workbook::new();

        // Final (imported) sheet names.
        let _sheet_b = workbook.add_sheet("B").unwrap();
        let sheet_c = workbook.add_sheet("C").unwrap();
        let _sheet_my = workbook.add_sheet("My Sheet").unwrap();

        {
            let sheet = workbook.sheet_mut(sheet_c).unwrap();

            // A->B, B->C: ensure we do not cascade so `A!A1` becomes `B!A1`, not `C!A1`.
            sheet.set_formula(
                CellRef::from_a1("A1").unwrap(),
                Some("=A!A1+B!A1+\"A!A1\"".to_string()),
            );

            // Preserve the canonical leading '=' behaviour (normalize strips only one leading '=').
            sheet.set_formula(CellRef::from_a1("A2").unwrap(), Some("==A!A1".to_string()));

            // Ensure we rewrite sheet refs inside formulas but not inside string literals.
            sheet.set_formula(
                CellRef::from_a1("A3").unwrap(),
                Some("='bad/name'!A1+\"bad/name!A1\"".to_string()),
            );
        }

        let renames = vec![
            ("a".to_string(), "B".to_string()),
            ("b".to_string(), "C".to_string()),
            ("bad/name".to_string(), "My Sheet".to_string()),
        ];

        rewrite_workbook_formulas_for_sheet_renames(&mut workbook, &renames);

        let sheet = workbook.sheet_by_name("C").unwrap();
        assert_eq!(
            sheet.formula_a1("A1").unwrap(),
            Some("B!A1+'C'!A1+\"A!A1\"")
        );
        assert_eq!(sheet.formula_a1("A2").unwrap(), Some("=B!A1"));
        assert_eq!(
            sheet.formula_a1("A3").unwrap(),
            Some("'My Sheet'!A1+\"bad/name!A1\"")
        );
    }
}
