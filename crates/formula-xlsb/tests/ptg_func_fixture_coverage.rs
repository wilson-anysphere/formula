use std::fs;
use std::path::{Path, PathBuf};

use formula_xlsb::format::format_a1;
use formula_xlsb::rgce::{decode_rgce_with_context_and_rgcb_and_base, CellCoord};
use formula_xlsb::{OpenOptions, XlsbWorkbook};

fn collect_xlsb_fixtures(dir: &Path) -> Vec<PathBuf> {
    let mut out = Vec::new();
    let entries = match fs::read_dir(dir) {
        Ok(entries) => entries,
        Err(_) => return out,
    };
    for entry in entries {
        let Ok(entry) = entry else { continue };
        let path = entry.path();
        if path.extension().and_then(|s| s.to_str()) == Some("xlsb") {
            out.push(path);
        }
    }
    out.sort();
    out
}

#[test]
fn xlsb_fixtures_do_not_contain_unsupported_ptgfunc_iftab_ids() {
    // `PtgFunc` encodes fixed-arity built-in calls without an explicit argc. If `formula-biff`
    // lacks fixed-arity metadata for a function, *entire* formula text decoding fails.
    let manifest = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    let mut fixtures = Vec::new();
    fixtures.extend(collect_xlsb_fixtures(&manifest.join("tests/fixtures")));
    fixtures.extend(collect_xlsb_fixtures(&manifest.join("tests/fixtures_metadata")));
    fixtures.extend(collect_xlsb_fixtures(&manifest.join("tests/fixtures_styles")));

    assert!(!fixtures.is_empty(), "expected at least one XLSB fixture");

    for path in fixtures {
        let wb = XlsbWorkbook::open_with_options(
            &path,
            OpenOptions {
                // Avoid relying on formula decoding during fixture load; we re-run the decoder
                // below so failures can be attributed to a specific ptg + iftab.
                decode_formulas: false,
                ..Default::default()
            },
        )
        .unwrap_or_else(|e| panic!("open XLSB fixture {}: {e}", path.display()));
        let ctx = wb.workbook_context();

        for sheet_index in 0..wb.sheet_metas().len() {
            let sheet_name = wb.sheet_metas()[sheet_index].name.clone();
            wb.for_each_cell(sheet_index, |cell| {
                let Some(formula) = cell.formula else {
                    return;
                };

                let base = CellCoord::new(cell.row, cell.col);
                if let Err(err) =
                    decode_rgce_with_context_and_rgcb_and_base(&formula.rgce, &formula.extra, ctx, base)
                {
                    // Only fail on `PtgFunc` issues; other ptg coverage is tracked separately.
                    if matches!(err.ptg(), Some(0x21 | 0x41 | 0x61)) {
                        let offset = err.offset();
                        assert!(
                            formula.rgce.len() >= offset + 3,
                            "fixture {} {sheet_name}!{} has truncated PtgFunc payload at offset {offset}",
                            path.display(),
                            format_a1(cell.row, cell.col),
                        );
                        let iftab = u16::from_le_bytes([formula.rgce[offset + 1], formula.rgce[offset + 2]]);
                        let spec = formula_biff::function_spec_from_id(iftab).unwrap_or_else(|| {
                            panic!(
                                "fixture {} {sheet_name}!{} uses PtgFunc(iftab={iftab}) but formula-biff has no FunctionSpec",
                                path.display(),
                                format_a1(cell.row, cell.col)
                            )
                        });
                        assert_eq!(
                            spec.min_args, spec.max_args,
                            "fixture {} {sheet_name}!{} uses PtgFunc for non-fixed function {} (arg range {}..={})",
                            path.display(),
                            format_a1(cell.row, cell.col),
                            spec.name,
                            spec.min_args,
                            spec.max_args
                        );
                    }
                }
            })
            .unwrap_or_else(|e| panic!("scan fixture {} sheet {sheet_index}: {e}", path.display()));
        }

        // Also scan defined names (which may contain XLM helper functions like GET.CELL).
        for name in wb.defined_names() {
            let Some(formula) = name.formula.as_ref() else {
                continue;
            };
            let base = CellCoord::new(0, 0);
            if let Err(err) =
                decode_rgce_with_context_and_rgcb_and_base(&formula.rgce, &formula.extra, ctx, base)
            {
                if matches!(err.ptg(), Some(0x21 | 0x41 | 0x61)) {
                    let offset = err.offset();
                    assert!(
                        formula.rgce.len() >= offset + 3,
                        "fixture {} defined name {} has truncated PtgFunc payload at offset {offset}",
                        path.display(),
                        name.name,
                    );
                    let iftab = u16::from_le_bytes([formula.rgce[offset + 1], formula.rgce[offset + 2]]);
                    let spec = formula_biff::function_spec_from_id(iftab).unwrap_or_else(|| {
                        panic!(
                            "fixture {} defined name {} uses PtgFunc(iftab={iftab}) but formula-biff has no FunctionSpec",
                            path.display(),
                            name.name
                        )
                    });
                    assert_eq!(
                        spec.min_args, spec.max_args,
                        "fixture {} defined name {} uses PtgFunc for non-fixed function {} (arg range {}..={})",
                        path.display(),
                        name.name,
                        spec.name,
                        spec.min_args,
                        spec.max_args
                    );
                }
            }
        }
    }
}

