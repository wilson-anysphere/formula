use std::collections::HashSet;

use formula_model::{rewrite_sheet_names_in_formula, Workbook};

#[derive(Debug, Clone)]
struct SheetRenameRewritePlan {
    old_to_tmp: Vec<(String, String)>,
    tmp_to_new: Vec<(String, String)>,
}

fn build_sheet_rename_rewrite_plan<'a>(
    sheet_names: impl IntoIterator<Item = &'a str>,
    sheet_rename_pairs: &[(String, String)],
) -> Option<SheetRenameRewritePlan> {
    if sheet_rename_pairs.is_empty() {
        return None;
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
        sheet_names
            .into_iter()
            .map(super::normalize_sheet_name_for_match),
    );

    let mut old_to_tmp: Vec<(String, String)> = Vec::new();
    let mut tmp_to_new: Vec<(String, String)> = Vec::new();

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

    Some(SheetRenameRewritePlan {
        old_to_tmp,
        tmp_to_new,
    })
}

fn rewrite_formula_with_plan(formula: &str, plan: &SheetRenameRewritePlan) -> String {
    let mut rewritten = formula.to_string();
    for (old, tmp) in &plan.old_to_tmp {
        rewritten = rewrite_sheet_names_in_formula(&rewritten, old, tmp);
    }
    for (tmp, new) in &plan.tmp_to_new {
        rewritten = rewrite_sheet_names_in_formula(&rewritten, tmp, new);
    }
    rewritten
}

/// Rewrite all workbook cell formulas to account for worksheet renames during import.
///
/// `sheet_rename_pairs` contains `(old_name, new_name)` pairs, using the sheet names as they
/// appear in the source workbook (calamine metadata and/or BIFF BoundSheet names).
pub(crate) fn rewrite_workbook_formulas_for_sheet_renames(
    workbook: &mut Workbook,
    sheet_rename_pairs: &[(String, String)],
) {
    let Some(plan) = build_sheet_rename_rewrite_plan(
        workbook.sheets.iter().map(|s| s.name.as_str()),
        sheet_rename_pairs,
    ) else {
        return;
    };

    for sheet in &mut workbook.sheets {
        for (_, cell) in sheet.iter_cells_mut() {
            let Some(formula) = cell.formula.as_mut() else {
                continue;
            };

            // Avoid cloning the existing formula string.
            let original = std::mem::take(formula);
            *formula = rewrite_formula_with_plan(&original, &plan);
        }
    }
}

#[cfg(test)]
pub(crate) fn rewrite_defined_name_formulas_for_sheet_renames(
    workbook: &mut Workbook,
    sheet_rename_pairs: &[(String, String)],
) {
    let Some(plan) = build_sheet_rename_rewrite_plan(
        workbook.sheets.iter().map(|s| s.name.as_str()),
        sheet_rename_pairs,
    ) else {
        return;
    };

    for name in &mut workbook.defined_names {
        let rewritten = rewrite_formula_with_plan(&name.refers_to, &plan);
        if rewritten != name.refers_to {
            name.refers_to = rewritten;
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use formula_engine::{parse_formula, ParseOptions};
    use formula_model::{CellRef, DefinedNameScope, Workbook};

    fn assert_parseable(expr: &str) {
        let expr = expr.trim();
        assert!(!expr.is_empty(), "expected formula text to be non-empty");
        // The formula parser accepts formulas both with and without a leading '='.
        parse_formula(expr, ParseOptions::default()).unwrap_or_else(|err| {
            panic!("expected formula to be parseable, expr={expr:?}, err={err:?}")
        });
    }

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
        workbook
            .create_defined_name(
                DefinedNameScope::Workbook,
                "TestName",
                "A!A1+B!A1+\"A!A1\"",
                None,
                false,
                None,
            )
            .unwrap();
        rewrite_defined_name_formulas_for_sheet_renames(&mut workbook, &renames);

        let sheet = workbook.sheet_by_name("C").unwrap();
        assert_eq!(
            sheet.formula_a1("A1").unwrap(),
            Some("B!A1+'C'!A1+\"A!A1\"")
        );
        assert_parseable(sheet.formula_a1("A1").unwrap().unwrap());
        assert_eq!(sheet.formula_a1("A2").unwrap(), Some("=B!A1"));
        assert_parseable(sheet.formula_a1("A2").unwrap().unwrap());
        assert_eq!(
            sheet.formula_a1("A3").unwrap(),
            Some("'My Sheet'!A1+\"bad/name!A1\"")
        );
        assert_parseable(sheet.formula_a1("A3").unwrap().unwrap());

        let dn = workbook
            .defined_names
            .iter()
            .find(|n| n.name == "TestName")
            .expect("defined name missing");
        assert_eq!(dn.refers_to, "B!A1+'C'!A1+\"A!A1\"");
        assert_parseable(&dn.refers_to);
    }
}
