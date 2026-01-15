#![cfg(not(target_arch = "wasm32"))]

use formula_engine::locale;
use formula_model::external_refs::format_external_key;
use proptest::prelude::*;
use proptest::test_runner::{Config, RngAlgorithm, TestRng, TestRunner};
use std::path::Path;
use std::sync::OnceLock;

const CASES: u32 = 64;
const LOCALE_ROUNDTRIP_SEED: [u8; 32] = [0x23; 32];

fn discover_translated_locale_ids() -> Vec<String> {
    let data_dir = Path::new(env!("CARGO_MANIFEST_DIR")).join("src/locale/data");
    let entries = std::fs::read_dir(&data_dir)
        .unwrap_or_else(|err| panic!("failed to read locale data dir {data_dir:?}: {err}"));

    let mut ids = Vec::new();
    for entry in entries {
        let entry = entry.unwrap_or_else(|err| {
            panic!("failed to read directory entry under locale data dir {data_dir:?}: {err}")
        });
        if !entry
            .file_type()
            .unwrap_or_else(|err| {
                panic!("failed to read file type for {entry:?} under {data_dir:?}: {err}")
            })
            .is_file()
        {
            continue;
        }

        let file_name = entry.file_name();
        let file_name = file_name.to_string_lossy();
        if file_name == "README.md" {
            continue;
        }
        if !file_name.ends_with(".tsv") {
            continue;
        }
        if file_name.ends_with(".errors.tsv") {
            continue;
        }

        let id = file_name.strip_suffix(".tsv").expect("checked .tsv suffix");
        ids.push(id.to_string());
    }

    // Keep stable ordering so failures are deterministic.
    ids.sort();
    ids
}

fn translated_locales() -> &'static [&'static locale::FormulaLocale] {
    static LOCALES: OnceLock<Box<[&'static locale::FormulaLocale]>> = OnceLock::new();
    LOCALES.get_or_init(|| {
        let ids = discover_translated_locale_ids();
        assert!(
            ids.iter().any(|id| id == "de-DE")
                && ids.iter().any(|id| id == "fr-FR")
                && ids.iter().any(|id| id == "es-ES"),
            "expected translated locale TSVs to include de-DE, fr-FR, and es-ES; found {ids:?}"
        );

        ids.iter()
            .map(|id| {
                locale::get_locale(id).unwrap_or_else(|| {
                    panic!(
                        "found locale TSV {id:?} but formula_engine::locale::get_locale returned None (is it registered?)"
                    )
                })
            })
            .collect::<Vec<_>>()
            .into_boxed_slice()
    })
}

fn assert_locale_roundtrip(canonical: &str) -> Result<(), TestCaseError> {
    for &loc in translated_locales() {
        let localized = locale::localize_formula(canonical, loc).map_err(|e| {
            TestCaseError::fail(format!(
                "localize_formula failed: locale={} canonical={canonical:?} err={e:?}",
                loc.id
            ))
        })?;

        let roundtrip = locale::canonicalize_formula(&localized, loc).map_err(|e| {
            TestCaseError::fail(format!(
                "canonicalize_formula failed: locale={} canonical={canonical:?} localized={localized:?} err={e:?}",
                loc.id
            ))
        })?;

        prop_assert_eq!(
            roundtrip,
            canonical,
            "canonicalize(localize(canonical)) mismatch for locale={} localized={:?}",
            loc.id,
            localized
        );
    }
    Ok(())
}

fn arb_decimal_number() -> impl Strategy<Value = String> {
    // Keep numeric literals small and printable. We include:
    // - leading decimal (e.g. `.5`)
    // - fixed-point decimals
    // - scientific notation
    prop_oneof![
        Just(".5".to_string()),
        (0u32..=1000, 0u32..=999).prop_map(|(int, frac)| format!("{int}.{frac:03}")),
        (0u32..=1000, 0u32..=99).prop_map(|(int, frac)| format!("{int}.{frac:02}")),
        (1u32..=9, 0u32..=99, 1u32..=6)
            .prop_map(|(int, frac, exp)| format!("{int}.{frac:02}E{exp}")),
        (1u32..=9, 0u32..=99, 1u32..=6)
            .prop_map(|(int, frac, exp)| format!("{int}.{frac:02}E-{exp}")),
    ]
}

fn arb_bool_literal() -> impl Strategy<Value = &'static str> {
    prop_oneof![Just("TRUE"), Just("FALSE")]
}

fn arb_error_literal() -> impl Strategy<Value = &'static str> {
    // Mix of errors with and without locale translations.
    prop_oneof![
        Just("#VALUE!"),
        Just("#NAME?"),
        Just("#REF!"),
        Just("#DIV/0!"),
        Just("#GETTING_DATA"),
        Just("#SPILL!"),
    ]
}

fn arb_cell_ref() -> impl Strategy<Value = &'static str> {
    prop_oneof![
        Just("A1"),
        Just("B2"),
        Just("$C$3"),
        Just("D$4"),
        Just("$E5"),
    ]
}

fn arb_ref_operand() -> impl Strategy<Value = &'static str> {
    prop_oneof![
        Just("A1"),
        Just("B1"),
        Just("C1"),
        Just("A1:B2"),
        Just("B2:C3"),
    ]
}

fn arb_workbook_name() -> impl Strategy<Value = String> {
    // Keep workbook names ASCII; include punctuation that *must not* be translated inside `[...]`.
    prop_oneof![
        Just("Book.xlsx".to_string()),
        Just("Book,1.xlsx".to_string()),
        Just("Work Book-1.xlsx".to_string()),
        Just("Bob's.xlsx".to_string()),
    ]
}

fn arb_external_cell_ref() -> impl Strategy<Value = String> {
    (arb_workbook_name(), arb_cell_ref()).prop_map(|(book, addr)| {
        let prefix = format_external_key(&book, "Sheet1");
        format!("{prefix}!{addr}")
    })
}

fn arb_structured_ref() -> impl Strategy<Value = String> {
    // Include:
    // - simple structured ref
    // - nested structured ref (contains commas inside brackets)
    // - escaped bracket (`]]`) inside nested structured refs
    prop_oneof![
        Just("Table1[Qty]".to_string()),
        Just("Table1[[#Headers],[Qty]]".to_string()),
        Just("Table1[[#Headers],[A]]B]]".to_string()),
    ]
}

fn arb_array_literal() -> impl Strategy<Value = String> {
    prop_oneof![
        // Basic 2x2 array.
        (
            arb_decimal_number(),
            arb_decimal_number(),
            arb_decimal_number(),
            arb_decimal_number()
        )
            .prop_map(|(a, b, c, d)| format!("{{{a},{b};{c},{d}}}")),
        // Array containing a function call to stress comma-disambiguation (arg separators vs array separators).
        (
            arb_decimal_number(),
            arb_decimal_number(),
            arb_decimal_number()
        )
            .prop_map(|(a, b, c)| { format!("{{SUM({a},{b}),{c}}}") }),
    ]
}

fn arb_union_intersection_expr() -> impl Strategy<Value = String> {
    // Intersection uses whitespace; union uses the locale list separator in canonical form (`,`).
    // Use parentheses to ensure commas are treated as the union operator, not function arg separators.
    (arb_ref_operand(), arb_ref_operand(), arb_ref_operand())
        .prop_map(|(a, b, c)| format!("({a},{b}) ({b},{c})"))
}

fn arb_canonical_formula() -> impl Strategy<Value = String> {
    prop_oneof![
        // Function calls + decimal numbers.
        (arb_decimal_number(), arb_decimal_number()).prop_map(|(a, b)| format!("=SUM({a},{b})")),
        // Function name + whitespace before `(`.
        (arb_decimal_number(), arb_decimal_number()).prop_map(|(a, b)| format!("=SUM ({a},{b})")),
        // Boolean literals.
        (
            arb_bool_literal(),
            arb_decimal_number(),
            arb_decimal_number()
        )
            .prop_map(|(cond, a, b)| format!("=IF({cond},{a},{b})")),
        // Error literals.
        (arb_error_literal(), arb_decimal_number())
            .prop_map(|(err, fallback)| { format!("=IFERROR({err},{fallback})") }),
        // Standalone error literal.
        arb_error_literal().prop_map(|err| format!("={err}")),
        // Array literals (including comma ambiguity inside nested calls).
        arb_array_literal().prop_map(|arr| format!("={arr}")),
        // Function call with array literal argument + trailing scalar arg.
        (arb_array_literal(), arb_decimal_number())
            .prop_map(|(arr, scalar)| { format!("=SUM({arr},{scalar})") }),
        // Structured references (brackets + commas that must not be translated).
        (arb_structured_ref(), arb_decimal_number())
            .prop_map(|(sref, n)| { format!("=SUM({sref},{n})") }),
        // External workbook references (brackets + punctuation that must not be translated).
        (arb_external_cell_ref(), arb_decimal_number())
            .prop_map(|(ext, n)| { format!("=SUM({ext},{n})") }),
        // Mix external + structured refs in a single argument list to stress bracket-depth tracking.
        (
            arb_external_cell_ref(),
            arb_structured_ref(),
            arb_decimal_number()
        )
            .prop_map(|(ext, sref, n)| format!("=SUM({ext},{sref},{n})"),),
        // Dotted localized names (CUBE* functions in fr-FR/es-ES), plus decimal punctuation.
        arb_decimal_number().prop_map(|n| { format!("=CUBEVALUE(\"conn\",\"member\",{n})") }),
        // `_xlfn.` prefix handling for translated functions.
        (1u32..=5, 1u32..=5).prop_map(|(rows, cols)| format!("=_xlfn.SEQUENCE({rows},{cols})")),
        // `_xlfn.` prefix + dotted localized function names.
        arb_decimal_number().prop_map(|n| { format!("=_xlfn.CUBEVALUE(\"conn\",\"member\",{n})") }),
        // Union + intersection.
        (arb_union_intersection_expr(), arb_decimal_number())
            .prop_map(|(refs, n)| { format!("=SUM({refs},{n})") }),
    ]
}

#[test]
fn locale_roundtrip_regressions() {
    // These are explicitly crafted to cover tricky edge cases that are easy to regress:
    // - commas inside external workbook brackets must not be translated
    // - escaped brackets (`]]`) inside structured refs must not confuse bracket tracking
    // - comma ambiguity inside arrays containing nested function calls
    let canonical_formulas = [
        "=SUM([Book,1.xlsx]Sheet1!A1,Table1[[#Headers],[Qty]],1.5)",
        "=COUNTA(Table1[[#Headers],[A]]B]])&\"]\"",
        "={SUM(1.5,2.5),3.5}",
    ];

    for canonical in canonical_formulas {
        // Use `unwrap` here so failures show the exact regression string in the panic.
        assert_locale_roundtrip(canonical).unwrap();
    }
}

#[test]
fn proptest_locale_localize_canonicalize_roundtrip() {
    let mut runner = TestRunner::new_with_rng(
        Config {
            cases: CASES,
            failure_persistence: None,
            ..Config::default()
        },
        TestRng::from_seed(RngAlgorithm::ChaCha, &LOCALE_ROUNDTRIP_SEED),
    );

    runner
        .run(&arb_canonical_formula(), |canonical| {
            assert_locale_roundtrip(&canonical)
        })
        .unwrap();
}
