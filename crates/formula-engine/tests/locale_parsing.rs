use formula_engine::value::RecordValue;
use formula_engine::{eval::parse_a1, locale, Engine, ErrorKind, ReferenceStyle, Value};
use std::collections::{HashMap, HashSet};

#[test]
fn canonicalize_and_localize_round_trip_for_de_de() {
    let localized = "=SUMME(1,5;2,5)";
    let canonical = locale::canonicalize_formula(localized, &locale::DE_DE).unwrap();
    assert_eq!(canonical, "=SUM(1.5,2.5)");

    let localized_roundtrip = locale::localize_formula(&canonical, &locale::DE_DE).unwrap();
    assert_eq!(localized_roundtrip, localized);
}

#[test]
fn canonicalize_and_localize_unicode_case_insensitive_function_names_for_de_de() {
    // German translation uses non-ASCII letters (Ä); ensure we do Unicode-aware case-folding.
    for localized in [
        "=zählenwenn(1;\">0\")",
        "=Zählenwenn(1;\">0\")",
        "=ZÄHLENWENN(1;\">0\")",
    ] {
        let canonical = locale::canonicalize_formula(localized, &locale::DE_DE).unwrap();
        assert_eq!(canonical, "=COUNTIF(1,\">0\")");
    }

    // Reverse translation should use the spelling from `src/locale/data/de-DE.tsv`.
    assert_eq!(
        locale::localize_formula("=countif(1,\">0\")", &locale::DE_DE).unwrap(),
        "=ZÄHLENWENN(1;\">0\")"
    );
}

#[test]
fn canonicalize_supports_thousands_and_leading_decimal_in_de_de() {
    let canonical = locale::canonicalize_formula("=SUMME(1.234,56;,5)", &locale::DE_DE).unwrap();
    assert_eq!(canonical, "=SUM(1234.56,.5)");
}

#[test]
fn canonicalize_accepts_canonical_leading_decimal_in_de_de() {
    let canonical = locale::canonicalize_formula("=SUMME(.5;1)", &locale::DE_DE).unwrap();
    assert_eq!(canonical, "=SUM(.5,1)");
}

#[test]
fn canonicalize_accepts_canonical_leading_decimal_in_fr_fr_and_es_es() {
    for (src, loc) in [
        ("=SOMME(.5;1)", &locale::FR_FR),
        ("=SUMA(.5;1)", &locale::ES_ES),
    ] {
        let canonical = locale::canonicalize_formula(src, loc).unwrap();
        assert_eq!(canonical, "=SUM(.5,1)");
    }
}

#[test]
fn canonicalize_and_localize_array_literals_for_de_de() {
    let localized = "=SUMME({1\\2;3\\4})";
    let canonical = locale::canonicalize_formula(localized, &locale::DE_DE).unwrap();
    assert_eq!(canonical, "=SUM({1,2;3,4})");

    let localized_roundtrip = locale::localize_formula(&canonical, &locale::DE_DE).unwrap();
    assert_eq!(localized_roundtrip, localized);
}

#[test]
fn canonicalize_and_localize_unions_for_de_de() {
    let localized = "(A1;B1)";
    let canonical = locale::canonicalize_formula(localized, &locale::DE_DE).unwrap();
    assert_eq!(canonical, "(A1,B1)");

    let localized_roundtrip = locale::localize_formula(&canonical, &locale::DE_DE).unwrap();
    assert_eq!(localized_roundtrip, localized);
}

#[test]
fn translates_xlfn_prefixed_functions() {
    let localized = "=_xlfn.SEQUENZ(1;2)";
    let canonical = locale::canonicalize_formula(localized, &locale::DE_DE).unwrap();
    assert_eq!(canonical, "=_xlfn.SEQUENCE(1,2)");

    let localized_roundtrip = locale::localize_formula(&canonical, &locale::DE_DE).unwrap();
    assert_eq!(localized_roundtrip, localized);
}

#[test]
fn translates_xlfn_prefixed_external_data_functions() {
    // Some Excel files may include `_xlfn.` prefixes; ensure we translate the base function name
    // even when the localized spelling contains dots (e.g. `VALEUR.CUBE`).
    let fr = "=_xlfn.VALEUR.CUBE(\"conn\";\"member\";1,5)";
    let canon = locale::canonicalize_formula(fr, &locale::FR_FR).unwrap();
    assert_eq!(canon, "=_xlfn.CUBEVALUE(\"conn\",\"member\",1.5)");
    assert_eq!(locale::localize_formula(&canon, &locale::FR_FR).unwrap(), fr);

    let es = "=_xlfn.VALOR.CUBO(\"conn\";\"member\";1,5)";
    let canon = locale::canonicalize_formula(es, &locale::ES_ES).unwrap();
    assert_eq!(canon, "=_xlfn.CUBEVALUE(\"conn\",\"member\",1.5)");
    assert_eq!(locale::localize_formula(&canon, &locale::ES_ES).unwrap(), es);
}

#[test]
fn translates_external_data_functions_with_whitespace_before_paren() {
    // Excel tolerates whitespace between the function name and `(`; ensure translation does too.
    for (locale, localized, canonical) in [
        (
            &locale::DE_DE,
            "=CUBEWERT (\"conn\";\"member\";1,5)",
            "=CUBEVALUE (\"conn\",\"member\",1.5)",
        ),
        (
            &locale::FR_FR,
            "=VALEUR.CUBE (\"conn\";\"member\";1,5)",
            "=CUBEVALUE (\"conn\",\"member\",1.5)",
        ),
        (
            &locale::ES_ES,
            "=VALOR.CUBO (\"conn\";\"member\";1,5)",
            "=CUBEVALUE (\"conn\",\"member\",1.5)",
        ),
    ] {
        assert_eq!(
            locale::canonicalize_formula(localized, locale).unwrap(),
            canonical
        );
        assert_eq!(locale::localize_formula(canonical, locale).unwrap(), localized);
    }
}

#[test]
fn canonicalize_and_localize_round_trip_for_fr_fr_and_es_es() {
    let fr = "=SOMME(1,5;2,5)";
    let fr_canon = locale::canonicalize_formula(fr, &locale::FR_FR).unwrap();
    assert_eq!(fr_canon, "=SUM(1.5,2.5)");
    assert_eq!(
        locale::localize_formula(&fr_canon, &locale::FR_FR).unwrap(),
        fr
    );

    let es = "=SUMA(1,5;2,5)";
    let es_canon = locale::canonicalize_formula(es, &locale::ES_ES).unwrap();
    assert_eq!(es_canon, "=SUM(1.5,2.5)");
    assert_eq!(
        locale::localize_formula(&es_canon, &locale::ES_ES).unwrap(),
        es
    );
}

#[test]
fn field_access_function_names_are_not_translated() {
    // The identifier after `.` is a field-access selector, not a function name. Even when it is
    // called with `(`, locale function translation must never rewrite it.

    let de_localized = "=A1.SUMME(1,5;2,5)";
    let de_canon = locale::canonicalize_formula(de_localized, &locale::DE_DE).unwrap();
    assert_eq!(de_canon, "=A1.SUMME(1.5,2.5)");
    assert_eq!(
        locale::localize_formula(&de_canon, &locale::DE_DE).unwrap(),
        de_localized
    );

    let fr_localized = "=A1.SOMME(1,5;2,5)";
    let fr_canon = locale::canonicalize_formula(fr_localized, &locale::FR_FR).unwrap();
    assert_eq!(fr_canon, "=A1.SOMME(1.5,2.5)");
    assert_eq!(
        locale::localize_formula(&fr_canon, &locale::FR_FR).unwrap(),
        fr_localized
    );

    let es_localized = "=A1.SUMA(1,5;2,5)";
    let es_canon = locale::canonicalize_formula(es_localized, &locale::ES_ES).unwrap();
    assert_eq!(es_canon, "=A1.SUMA(1.5,2.5)");
    assert_eq!(
        locale::localize_formula(&es_canon, &locale::ES_ES).unwrap(),
        es_localized
    );
}

#[test]
fn canonicalize_supports_nbsp_thousands_separator_in_fr_fr() {
    // French Excel commonly uses NBSP (U+00A0) for thousands grouping.
    let fr = "=SOMME(1\u{00A0}234,56;0,5)";
    let canon = locale::canonicalize_formula(fr, &locale::FR_FR).unwrap();
    assert_eq!(canon, "=SUM(1234.56,0.5)");
}

#[test]
fn canonicalize_supports_narrow_nbsp_thousands_separator_in_fr_fr() {
    // Some French locales/spreadsheets use narrow NBSP (U+202F) for thousands grouping.
    let fr = "=SOMME(1\u{202F}234,56;0,5)";
    let canon = locale::canonicalize_formula(fr, &locale::FR_FR).unwrap();
    assert_eq!(canon, "=SUM(1234.56,0.5)");
}

#[test]
fn fr_fr_does_not_treat_ascii_spaces_as_thousands_separators() {
    // ASCII spaces are still significant for Excel (whitespace / intersection operator) and must
    // not be silently stripped out of numeric literals.
    let fr = "=SOMME(1 234,56;0,5)";
    let canon = locale::canonicalize_formula(fr, &locale::FR_FR).unwrap();
    assert_eq!(canon, "=SUM(1 234.56,0.5)");
}

#[test]
fn canonicalize_and_localize_additional_function_names_for_fr_fr() {
    fn assert_roundtrip(canonical: &str, localized: &str) {
        assert_eq!(
            locale::canonicalize_formula(localized, &locale::FR_FR).unwrap(),
            canonical
        );
        assert_eq!(
            locale::localize_formula(canonical, &locale::FR_FR).unwrap(),
            localized
        );
    }

    // Common functions with strongly localized spellings.
    assert_roundtrip("=COUNTIF(A1:A3,\">0\")", "=NB.SI(A1:A3;\">0\")");
    assert_roundtrip(
        "=SUMIF(A1:A3,\">0\",B1:B3)",
        "=SOMME.SI(A1:A3;\">0\";B1:B3)",
    );
    assert_roundtrip("=AVERAGEIF(A1:A3,\">0\")", "=MOYENNE.SI(A1:A3;\">0\")");
    assert_roundtrip(
        "=VLOOKUP(1,A1:B3,2,FALSE)",
        "=RECHERCHEV(1;A1:B3;2;FAUX)",
    );
    assert_roundtrip(
        "=HLOOKUP(1,A1:C2,2,FALSE)",
        "=RECHERCHEH(1;A1:C2;2;FAUX)",
    );
    assert_roundtrip("=LEFT(\"abc\",2)", "=GAUCHE(\"abc\";2)");
    assert_roundtrip("=RIGHT(\"abc\",2)", "=DROITE(\"abc\";2)");
    assert_roundtrip("=MID(\"abc\",2,1)", "=STXT(\"abc\";2;1)");
    assert_roundtrip("=LEN(\"abc\")", "=NBCAR(\"abc\")");
    assert_roundtrip("=FIND(\"b\",\"abc\")", "=TROUVE(\"b\";\"abc\")");
    assert_roundtrip("=SEARCH(\"B\",\"abc\")", "=CHERCHE(\"B\";\"abc\")");
    assert_roundtrip("=IFERROR(1/0,0)", "=SIERREUR(1/0;0)");
    assert_roundtrip(
        "=IFS(1=0,\"no\",1=1,\"yes\")",
        "=SI.CONDITIONS(1=0;\"no\";1=1;\"yes\")",
    );

    // TRUE()/FALSE() also exist as zero-arg worksheet functions and are localized in Excel.
    assert_roundtrip("=TRUE()", "=VRAI()");
    assert_roundtrip("=FALSE()", "=FAUX()");

    // `_xlfn.`-prefixed formulas still translate the base function name.
    assert_roundtrip(
        "=_xlfn.XLOOKUP(1,A1:A3,B1:B3)",
        "=_xlfn.RECHERCHEX(1;A1:A3;B1:B3)",
    );
}

#[test]
fn canonicalize_and_localize_true_false_functions_for_de_de_and_es_es() {
    fn assert_roundtrip(locale: &locale::FormulaLocale, canonical: &str, localized: &str) {
        assert_eq!(
            locale::canonicalize_formula(localized, locale).unwrap(),
            canonical
        );
        assert_eq!(locale::localize_formula(canonical, locale).unwrap(), localized);
    }

    // TRUE()/FALSE() also exist as zero-arg worksheet functions and are localized in Excel.
    assert_roundtrip(&locale::DE_DE, "=TRUE()", "=WAHR()");
    assert_roundtrip(&locale::DE_DE, "=FALSE()", "=FALSCH()");
    assert_roundtrip(&locale::ES_ES, "=TRUE()", "=VERDADERO()");
    assert_roundtrip(&locale::ES_ES, "=FALSE()", "=FALSO()");
}

#[test]
fn structured_reference_items_are_not_translated() {
    // Excel keeps structured-reference item keywords (e.g. `#Headers`) and the inner separators
    // inside `Table1[[...],[...]]` canonical (not locale-dependent). We translate only the
    // surrounding formula syntax (function name + argument separators).
    for (canonical, table_spec) in [
        ("=SUM(Table1[#All],1)", "Table1[#All]"),
        ("=SUM(Table1[#Data],1)", "Table1[#Data]"),
        ("=SUM(Table1[#Totals],1)", "Table1[#Totals]"),
        ("=SUM(Table1[[#Headers],[Qty]],1)", "Table1[[#Headers],[Qty]]"),
        ("=SUM(Table1[[#This Row],[Qty]],1)", "Table1[[#This Row],[Qty]]"),
        (
            "=SUM(Table1[[#All],[Col1],[Col2]],1)",
            "Table1[[#All],[Col1],[Col2]]",
        ),
    ] {
        for locale in [&locale::DE_DE, &locale::FR_FR, &locale::ES_ES] {
            let expected = format!(
                "={}({}{}1)",
                locale.localized_function_name("SUM"),
                table_spec,
                locale.config.arg_separator
            );
            let localized = locale::localize_formula(canonical, locale).unwrap();
            assert_eq!(localized, expected);

            let canonical_roundtrip = locale::canonicalize_formula(&localized, locale).unwrap();
            assert_eq!(canonical_roundtrip, canonical);
        }
    }
}

#[test]
fn de_de_translation_table_covers_function_catalog() {
    let mut covered = HashSet::new();
    let tsv = include_str!("../src/locale/data/de-DE.tsv");
    for line in tsv.lines() {
        let line = line.trim();
        if line.is_empty() || line.starts_with('#') {
            continue;
        }
        let (canon, _loc) = line
            .split_once('\t')
            .unwrap_or_else(|| panic!("invalid TSV line in de-DE.tsv: {line:?}"));
        assert!(
            covered.insert(canon),
            "duplicate canonical entry in de-DE.tsv: {canon}"
        );
    }

    for spec in formula_engine::functions::iter_function_specs() {
        let name = spec.name.to_ascii_uppercase();
        assert!(
            covered.contains(name.as_str()),
            "missing de-DE function translation for {name}"
        );
    }

    let expected_count = formula_engine::functions::iter_function_specs().count();
    assert_eq!(
        covered.len(),
        expected_count,
        "de-DE.tsv should contain exactly one entry per function in the engine catalog"
    );
}

#[test]
fn fr_fr_translation_table_covers_function_catalog() {
    let mut covered = HashSet::new();
    let tsv = include_str!("../src/locale/data/fr-FR.tsv");
    for line in tsv.lines() {
        let line = line.trim();
        if line.is_empty() || line.starts_with('#') {
            continue;
        }
        let (canon, _loc) = line
            .split_once('\t')
            .unwrap_or_else(|| panic!("invalid TSV line in fr-FR.tsv: {line:?}"));
        assert!(
            covered.insert(canon),
            "duplicate canonical entry in fr-FR.tsv: {canon}"
        );
    }

    for spec in formula_engine::functions::iter_function_specs() {
        let name = spec.name.to_ascii_uppercase();
        assert!(
            covered.contains(name.as_str()),
            "missing fr-FR function translation for {name}"
        );
    }

    let expected_count = formula_engine::functions::iter_function_specs().count();
    assert_eq!(
        covered.len(),
        expected_count,
        "fr-FR.tsv should contain exactly one entry per function in the engine catalog"
    );
}

#[test]
fn es_es_translation_table_covers_function_catalog() {
    let mut covered = HashSet::new();
    let tsv = include_str!("../src/locale/data/es-ES.tsv");
    for line in tsv.lines() {
        let line = line.trim();
        if line.is_empty() || line.starts_with('#') {
            continue;
        }
        let (canon, _loc) = line
            .split_once('\t')
            .unwrap_or_else(|| panic!("invalid TSV line in es-ES.tsv: {line:?}"));
        assert!(
            covered.insert(canon),
            "duplicate canonical entry in es-ES.tsv: {canon}"
        );
    }

    for spec in formula_engine::functions::iter_function_specs() {
        let name = spec.name.to_ascii_uppercase();
        assert!(
            covered.contains(name.as_str()),
            "missing es-ES function translation for {name}"
        );
    }

    let expected_count = formula_engine::functions::iter_function_specs().count();
    assert_eq!(
        covered.len(),
        expected_count,
        "es-ES.tsv should contain exactly one entry per function in the engine catalog"
    );
}

#[test]
fn canonicalize_and_localize_boolean_literals() {
    let de = "=WENN(WAHR;1;0)";
    let canon = locale::canonicalize_formula(de, &locale::DE_DE).unwrap();
    assert_eq!(canon, "=IF(TRUE,1,0)");
    assert_eq!(
        locale::localize_formula(&canon, &locale::DE_DE).unwrap(),
        de
    );

    let fr = "=SI(VRAI;1;0)";
    let canon = locale::canonicalize_formula(fr, &locale::FR_FR).unwrap();
    assert_eq!(canon, "=IF(TRUE,1,0)");
    assert_eq!(
        locale::localize_formula(&canon, &locale::FR_FR).unwrap(),
        fr
    );

    let es = "=SI(VERDADERO;1;0)";
    let canon = locale::canonicalize_formula(es, &locale::ES_ES).unwrap();
    assert_eq!(canon, "=IF(TRUE,1,0)");
    assert_eq!(
        locale::localize_formula(&canon, &locale::ES_ES).unwrap(),
        es
    );
}

#[test]
fn canonicalize_and_localize_true_false_functions() {
    for (locale, localized_true, localized_false) in [
        (&locale::DE_DE, "=WAHR()", "=FALSCH()"),
        (&locale::FR_FR, "=VRAI()", "=FAUX()"),
        (&locale::ES_ES, "=VERDADERO()", "=FALSO()"),
    ] {
        let canon_true = locale::canonicalize_formula(localized_true, locale).unwrap();
        assert_eq!(canon_true, "=TRUE()");
        assert_eq!(
            locale::localize_formula(&canon_true, locale).unwrap(),
            localized_true
        );

        let canon_false = locale::canonicalize_formula(localized_false, locale).unwrap();
        assert_eq!(canon_false, "=FALSE()");
        assert_eq!(
            locale::localize_formula(&canon_false, locale).unwrap(),
            localized_false
        );
    }
}

#[test]
fn localized_boolean_keywords_are_not_translated_inside_structured_refs() {
    // `WAHR` is the de-DE TRUE keyword, but table names can still be identifiers; separators
    // inside structured refs should never be touched by translation.
    let localized = "=SUMME(WAHR[Col])";
    let canonical = locale::canonicalize_formula(localized, &locale::DE_DE).unwrap();
    assert_eq!(canonical, "=SUM(WAHR[Col])");

    let localized_roundtrip = locale::localize_formula(&canonical, &locale::DE_DE).unwrap();
    assert_eq!(localized_roundtrip, localized);
}

#[test]
fn localized_boolean_keywords_are_not_translated_in_3d_sheet_spans() {
    // In de-DE, `WAHR` is the TRUE keyword, but it can also be a sheet name.
    // Ensure we treat `WAHR:Sheet3!A1` as a 3D sheet span, not a boolean literal.
    let localized = "=SUMME(WAHR:Sheet3!A1)";
    let canonical = locale::canonicalize_formula(localized, &locale::DE_DE).unwrap();
    assert_eq!(canonical, "=SUM(WAHR:Sheet3!A1)");
}

#[test]
fn canonicalize_and_localize_error_literals() {
    let de = "=#WERT!";
    let canon = locale::canonicalize_formula(de, &locale::DE_DE).unwrap();
    assert_eq!(canon, "=#VALUE!");
    assert_eq!(
        locale::localize_formula(&canon, &locale::DE_DE).unwrap(),
        de
    );

    // Non-ASCII localized errors should be translated using Unicode-aware case folding.
    // de-DE: `#ÜBERLAUF!` is the localized spelling for `#SPILL!`.
    let de_spill_variants = ["=#ÜBERLAUF!", "=#Überlauf!", "=#üBeRlAuF!"];
    for src in de_spill_variants {
        let canon = locale::canonicalize_formula(src, &locale::DE_DE).unwrap();
        assert_eq!(canon, "=#SPILL!");
        assert_eq!(
            locale::localize_formula(&canon, &locale::DE_DE).unwrap(),
            "=#ÜBERLAUF!"
        );
    }

    let fr = "=#VALEUR!";
    let canon = locale::canonicalize_formula(fr, &locale::FR_FR).unwrap();
    assert_eq!(canon, "=#VALUE!");
    assert_eq!(
        locale::localize_formula(&canon, &locale::FR_FR).unwrap(),
        fr
    );
}

#[test]
fn canonicalize_normalizes_canonical_error_variants() {
    assert_eq!(
        locale::canonicalize_formula("=#n/a!", &locale::EN_US).unwrap(),
        "=#N/A"
    );
    assert_eq!(
        locale::canonicalize_formula("=#value!", &locale::EN_US).unwrap(),
        "=#VALUE!"
    );
}

#[test]
fn localize_normalizes_canonical_error_variants_before_translation() {
    // Some legacy/stored formulas may contain `#N/A!` instead of canonical `#N/A`.
    // Ensure we normalize before attempting locale error literal lookup.
    let expected = locale::DE_DE
        .localized_error_literal("#N/A")
        .unwrap_or("#N/A");
    assert_eq!(
        locale::localize_formula("=#N/A!", &locale::DE_DE).unwrap(),
        format!("={expected}")
    );
}

// NOTE: Localized spellings in these tests are based on Microsoft Excel's function/error
// translations for the de-DE / fr-FR / es-ES locales. Keep these in sync with
// `src/locale/data/*.tsv` and `src/locale/registry.rs`.
#[test]
fn canonicalize_and_localize_external_data_functions_and_errors_for_de_de() {
    let localized_rtd = "=RTD(\"my.server\";\"topic\";1,5)";
    let canonical_rtd = locale::canonicalize_formula(localized_rtd, &locale::DE_DE).unwrap();
    assert_eq!(canonical_rtd, "=RTD(\"my.server\",\"topic\",1.5)");
    assert_eq!(
        locale::localize_formula(&canonical_rtd, &locale::DE_DE).unwrap(),
        localized_rtd
    );

    let localized_cube = "=CUBEWERT(\"conn\";\"member\";1,5)";
    let canonical_cube = locale::canonicalize_formula(localized_cube, &locale::DE_DE).unwrap();
    assert_eq!(canonical_cube, "=CUBEVALUE(\"conn\",\"member\",1.5)");
    assert_eq!(
        locale::localize_formula(&canonical_cube, &locale::DE_DE).unwrap(),
        localized_cube
    );

    let localized_err = "=#DATEN_ABRUFEN";
    let canonical_err = locale::canonicalize_formula(localized_err, &locale::DE_DE).unwrap();
    assert_eq!(canonical_err, "=#GETTING_DATA");
    assert_eq!(
        locale::localize_formula(&canonical_err, &locale::DE_DE).unwrap(),
        localized_err
    );

    // These error literals currently round-trip unchanged for this locale.
    for err in ["#CONNECT!", "#FIELD!", "#BLOCKED!", "#UNKNOWN!"] {
        let src = format!("={err}");
        let canon = locale::canonicalize_formula(&src, &locale::DE_DE).unwrap();
        assert_eq!(canon, src);
        assert_eq!(locale::localize_formula(&canon, &locale::DE_DE).unwrap(), src);
    }
}

#[test]
fn canonicalize_and_localize_external_data_functions_and_errors_for_fr_fr() {
    let localized_rtd = "=RTD(\"my.server\";\"topic\";1,5)";
    let canonical_rtd = locale::canonicalize_formula(localized_rtd, &locale::FR_FR).unwrap();
    assert_eq!(canonical_rtd, "=RTD(\"my.server\",\"topic\",1.5)");
    assert_eq!(
        locale::localize_formula(&canonical_rtd, &locale::FR_FR).unwrap(),
        localized_rtd
    );

    let localized_cube = "=VALEUR.CUBE(\"conn\";\"member\";1,5)";
    let canonical_cube = locale::canonicalize_formula(localized_cube, &locale::FR_FR).unwrap();
    assert_eq!(canonical_cube, "=CUBEVALUE(\"conn\",\"member\",1.5)");
    assert_eq!(
        locale::localize_formula(&canonical_cube, &locale::FR_FR).unwrap(),
        localized_cube
    );

    let localized_err = "=#OBTENTION_DONNEES";
    let canonical_err = locale::canonicalize_formula(localized_err, &locale::FR_FR).unwrap();
    assert_eq!(canonical_err, "=#GETTING_DATA");
    assert_eq!(
        locale::localize_formula(&canonical_err, &locale::FR_FR).unwrap(),
        localized_err
    );

    // These error literals currently round-trip unchanged for this locale.
    for err in ["#CONNECT!", "#FIELD!", "#BLOCKED!", "#UNKNOWN!"] {
        let src = format!("={err}");
        let canon = locale::canonicalize_formula(&src, &locale::FR_FR).unwrap();
        assert_eq!(canon, src);
        assert_eq!(locale::localize_formula(&canon, &locale::FR_FR).unwrap(), src);
    }
}

#[test]
fn canonicalize_and_localize_external_data_functions_and_errors_for_es_es() {
    let localized_rtd = "=RTD(\"my.server\";\"topic\";1,5)";
    let canonical_rtd = locale::canonicalize_formula(localized_rtd, &locale::ES_ES).unwrap();
    assert_eq!(canonical_rtd, "=RTD(\"my.server\",\"topic\",1.5)");
    assert_eq!(
        locale::localize_formula(&canonical_rtd, &locale::ES_ES).unwrap(),
        localized_rtd
    );

    let localized_cube = "=VALOR.CUBO(\"conn\";\"member\";1,5)";
    let canonical_cube = locale::canonicalize_formula(localized_cube, &locale::ES_ES).unwrap();
    assert_eq!(canonical_cube, "=CUBEVALUE(\"conn\",\"member\",1.5)");
    assert_eq!(
        locale::localize_formula(&canonical_cube, &locale::ES_ES).unwrap(),
        localized_cube
    );

    let localized_err = "=#OBTENIENDO_DATOS";
    let canonical_err = locale::canonicalize_formula(localized_err, &locale::ES_ES).unwrap();
    assert_eq!(canonical_err, "=#GETTING_DATA");
    assert_eq!(
        locale::localize_formula(&canonical_err, &locale::ES_ES).unwrap(),
        localized_err
    );

    // These error literals currently round-trip unchanged for this locale.
    for err in ["#CONNECT!", "#FIELD!", "#BLOCKED!", "#UNKNOWN!"] {
        let src = format!("={err}");
        let canon = locale::canonicalize_formula(&src, &locale::ES_ES).unwrap();
        assert_eq!(canon, src);
        assert_eq!(locale::localize_formula(&canon, &locale::ES_ES).unwrap(), src);
    }
}

#[test]
fn canonicalize_and_localize_all_cube_function_names() {
    fn assert_roundtrip(locale: &locale::FormulaLocale, canonical: &str, localized: &str) {
        assert_eq!(
            locale::canonicalize_formula(localized, locale).unwrap(),
            canonical
        );
        assert_eq!(locale::localize_formula(canonical, locale).unwrap(), localized);
    }

    // de-DE
    assert_roundtrip(
        &locale::DE_DE,
        "=CUBEVALUE(\"conn\",\"member\",1.5)",
        "=CUBEWERT(\"conn\";\"member\";1,5)",
    );
    assert_roundtrip(
        &locale::DE_DE,
        "=CUBEMEMBER(\"conn\",\"member\",\"caption\")",
        "=CUBEMITGLIED(\"conn\";\"member\";\"caption\")",
    );
    assert_roundtrip(
        &locale::DE_DE,
        "=CUBEMEMBERPROPERTY(\"conn\",\"member\",\"prop\")",
        "=CUBEMITGLIEDSEIGENSCHAFT(\"conn\";\"member\";\"prop\")",
    );
    assert_roundtrip(
        &locale::DE_DE,
        "=CUBERANKEDMEMBER(\"conn\",\"set\",3,\"caption\")",
        "=CUBERANGMITGLIED(\"conn\";\"set\";3;\"caption\")",
    );
    assert_roundtrip(
        &locale::DE_DE,
        "=CUBESET(\"conn\",\"set\",\"caption\",1,\"sort\")",
        "=CUBEMENGE(\"conn\";\"set\";\"caption\";1;\"sort\")",
    );
    assert_roundtrip(
        &locale::DE_DE,
        "=CUBESETCOUNT(\"set\")",
        "=CUBEMENGENANZAHL(\"set\")",
    );
    assert_roundtrip(
        &locale::DE_DE,
        "=CUBEKPIMEMBER(\"conn\",\"kpi\",\"property\",\"caption\")",
        "=CUBEKPIMITGLIED(\"conn\";\"kpi\";\"property\";\"caption\")",
    );

    // fr-FR
    assert_roundtrip(
        &locale::FR_FR,
        "=CUBEVALUE(\"conn\",\"member\",1.5)",
        "=VALEUR.CUBE(\"conn\";\"member\";1,5)",
    );
    assert_roundtrip(
        &locale::FR_FR,
        "=CUBEMEMBER(\"conn\",\"member\",\"caption\")",
        "=MEMBRE.CUBE(\"conn\";\"member\";\"caption\")",
    );
    assert_roundtrip(
        &locale::FR_FR,
        "=CUBEMEMBERPROPERTY(\"conn\",\"member\",\"prop\")",
        "=PROPRIETE.MEMBRE.CUBE(\"conn\";\"member\";\"prop\")",
    );
    assert_roundtrip(
        &locale::FR_FR,
        "=CUBERANKEDMEMBER(\"conn\",\"set\",3,\"caption\")",
        "=MEMBRE.RANG.CUBE(\"conn\";\"set\";3;\"caption\")",
    );
    assert_roundtrip(
        &locale::FR_FR,
        "=CUBESET(\"conn\",\"set\",\"caption\",1,\"sort\")",
        "=ENSEMBLE.CUBE(\"conn\";\"set\";\"caption\";1;\"sort\")",
    );
    assert_roundtrip(
        &locale::FR_FR,
        "=CUBESETCOUNT(\"set\")",
        "=NB.ENSEMBLE.CUBE(\"set\")",
    );
    assert_roundtrip(
        &locale::FR_FR,
        "=CUBEKPIMEMBER(\"conn\",\"kpi\",\"property\",\"caption\")",
        "=MEMBREKPI.CUBE(\"conn\";\"kpi\";\"property\";\"caption\")",
    );

    // es-ES
    assert_roundtrip(
        &locale::ES_ES,
        "=CUBEVALUE(\"conn\",\"member\",1.5)",
        "=VALOR.CUBO(\"conn\";\"member\";1,5)",
    );
    assert_roundtrip(
        &locale::ES_ES,
        "=CUBEMEMBER(\"conn\",\"member\",\"caption\")",
        "=MIEMBRO.CUBO(\"conn\";\"member\";\"caption\")",
    );
    assert_roundtrip(
        &locale::ES_ES,
        "=CUBEMEMBERPROPERTY(\"conn\",\"member\",\"prop\")",
        "=PROPIEDAD.MIEMBRO.CUBO(\"conn\";\"member\";\"prop\")",
    );
    assert_roundtrip(
        &locale::ES_ES,
        "=CUBERANKEDMEMBER(\"conn\",\"set\",3,\"caption\")",
        "=MIEMBRO.RANGO.CUBO(\"conn\";\"set\";3;\"caption\")",
    );
    assert_roundtrip(
        &locale::ES_ES,
        "=CUBESET(\"conn\",\"set\",\"caption\",1,\"sort\")",
        "=CONJUNTO.CUBO(\"conn\";\"set\";\"caption\";1;\"sort\")",
    );
    assert_roundtrip(
        &locale::ES_ES,
        "=CUBESETCOUNT(\"set\")",
        "=CONTAR.CONJUNTO.CUBO(\"set\")",
    );
    assert_roundtrip(
        &locale::ES_ES,
        "=CUBEKPIMEMBER(\"conn\",\"kpi\",\"property\",\"caption\")",
        "=MIEMBROKPI.CUBO(\"conn\";\"kpi\";\"property\";\"caption\")",
    );
}

#[test]
fn canonicalize_and_localize_with_style_r1c1_for_external_data_functions() {
    for (locale, localized, canonical) in [
        (
            &locale::DE_DE,
            "=CUBEWERT(\"conn\";R[-4]C[-2];1,5)",
            "=CUBEVALUE(\"conn\",R[-4]C[-2],1.5)",
        ),
        (
            &locale::FR_FR,
            "=VALEUR.CUBE(\"conn\";R[-4]C[-2];1,5)",
            "=CUBEVALUE(\"conn\",R[-4]C[-2],1.5)",
        ),
        (
            &locale::ES_ES,
            "=VALOR.CUBO(\"conn\";R[-4]C[-2];1,5)",
            "=CUBEVALUE(\"conn\",R[-4]C[-2],1.5)",
        ),
    ] {
        let canon = locale::canonicalize_formula_with_style(localized, locale, ReferenceStyle::R1C1)
            .unwrap();
        assert_eq!(canon, canonical);
        let localized_roundtrip =
            locale::localize_formula_with_style(&canon, locale, ReferenceStyle::R1C1).unwrap();
        assert_eq!(localized_roundtrip, localized);
    }
}

#[test]
fn localized_function_names_are_not_translated_in_sheet_prefixes() {
    let de = "=SUMME(CUBEWERT!A1;1)";
    let canon = locale::canonicalize_formula(de, &locale::DE_DE).unwrap();
    assert_eq!(canon, "=SUM(CUBEWERT!A1,1)");
    assert_eq!(locale::localize_formula(&canon, &locale::DE_DE).unwrap(), de);

    let fr = "=SOMME(VALEUR.CUBE!A1;1)";
    let canon = locale::canonicalize_formula(fr, &locale::FR_FR).unwrap();
    assert_eq!(canon, "=SUM(VALEUR.CUBE!A1,1)");
    assert_eq!(locale::localize_formula(&canon, &locale::FR_FR).unwrap(), fr);

    let es = "=SUMA(VALOR.CUBO!A1;1)";
    let canon = locale::canonicalize_formula(es, &locale::ES_ES).unwrap();
    assert_eq!(canon, "=SUM(VALOR.CUBO!A1,1)");
    assert_eq!(locale::localize_formula(&canon, &locale::ES_ES).unwrap(), es);
}

#[test]
fn localized_external_data_function_names_are_not_translated_when_used_as_identifiers() {
    // Function-name translation should only happen for identifiers used in function-call position
    // (`NAME(`). If a workbook has a defined name that happens to match a localized spelling, it
    // must not be rewritten.
    let de = "=CUBEWERT+1";
    let canon = locale::canonicalize_formula(de, &locale::DE_DE).unwrap();
    assert_eq!(canon, de);
    assert_eq!(locale::localize_formula(&canon, &locale::DE_DE).unwrap(), de);

    let fr = "=VALEUR.CUBE+1";
    let canon = locale::canonicalize_formula(fr, &locale::FR_FR).unwrap();
    assert_eq!(canon, fr);
    assert_eq!(locale::localize_formula(&canon, &locale::FR_FR).unwrap(), fr);

    let es = "=VALOR.CUBO+1";
    let canon = locale::canonicalize_formula(es, &locale::ES_ES).unwrap();
    assert_eq!(canon, es);
    assert_eq!(locale::localize_formula(&canon, &locale::ES_ES).unwrap(), es);
}

#[test]
fn engine_accepts_localized_formulas_and_persists_canonical() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula_localized("Sheet1", "A1", "=SUMME(1,5;2,5)", &locale::DE_DE)
        .unwrap();
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(4.0));
    assert_eq!(
        engine.get_cell_formula("Sheet1", "A1"),
        Some("=SUM(1.5,2.5)")
    );

    let displayed = locale::localize_formula(
        engine.get_cell_formula("Sheet1", "A1").unwrap(),
        &locale::DE_DE,
    )
    .unwrap();
    assert_eq!(displayed, "=SUMME(1,5;2,5)");
}

#[test]
fn engine_accepts_localized_true_false_functions_and_persists_canonical() {
    for (locale, localized_true, localized_false) in [
        (&locale::DE_DE, "=WAHR()", "=FALSCH()"),
        (&locale::ES_ES, "=VERDADERO()", "=FALSO()"),
    ] {
        let mut engine = Engine::new();
        engine
            .set_cell_formula_localized("Sheet1", "A1", localized_true, locale)
            .unwrap();
        engine
            .set_cell_formula_localized("Sheet1", "A2", localized_false, locale)
            .unwrap();

        assert_eq!(engine.get_cell_formula("Sheet1", "A1"), Some("=TRUE()"));
        assert_eq!(engine.get_cell_formula("Sheet1", "A2"), Some("=FALSE()"));

        engine.recalculate();
        assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Bool(true));
        assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Bool(false));
    }
}

#[test]
fn engine_get_cell_formula_localized_displays_locale_specific_formula() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=SUM(1.5,2.5)")
        .unwrap();

    assert_eq!(
        engine
            .get_cell_formula_localized("Sheet1", "A1", &locale::DE_DE)
            .as_deref(),
        Some("=SUMME(1,5;2,5)")
    );
}

#[test]
fn engine_get_cell_formula_localized_r1c1_displays_locale_specific_references_and_separators() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "C5", "=SUM(A1,B1)")
        .unwrap();

    assert_eq!(
        engine
            .get_cell_formula_localized_r1c1("Sheet1", "C5", &locale::DE_DE)
            .as_deref(),
        Some("=SUMME(R[-4]C[-2];R[-4]C[-1])")
    );
}

#[test]
fn engine_accepts_localized_external_data_formulas_and_persists_canonical() {
    for (locale, localized_cube, localized_err) in [
        (
            &locale::DE_DE,
            "=CUBEWERT(\"conn\";\"member\";1,5)",
            "=#DATEN_ABRUFEN",
        ),
        (
            &locale::FR_FR,
            "=VALEUR.CUBE(\"conn\";\"member\";1,5)",
            "=#OBTENTION_DONNEES",
        ),
        (
            &locale::ES_ES,
            "=VALOR.CUBO(\"conn\";\"member\";1,5)",
            "=#OBTENIENDO_DATOS",
        ),
    ] {
        let mut engine = Engine::new();

        engine
            .set_cell_formula_localized("Sheet1", "A1", "=RTD(\"my.server\";\"topic\";1,5)", locale)
            .unwrap();
        assert_eq!(
            engine.get_cell_formula("Sheet1", "A1"),
            Some("=RTD(\"my.server\",\"topic\",1.5)")
        );

        engine
            .set_cell_formula_localized("Sheet1", "A2", localized_cube, locale)
            .unwrap();
        assert_eq!(
            engine.get_cell_formula("Sheet1", "A2"),
            Some("=CUBEVALUE(\"conn\",\"member\",1.5)")
        );

        engine
            .set_cell_formula_localized("Sheet1", "A3", localized_err, locale)
            .unwrap();
        assert_eq!(
            engine.get_cell_formula("Sheet1", "A3"),
            Some("=#GETTING_DATA")
        );

        // Without an external data provider, RTD/CUBE* functions should be recognized and return
        // `#N/A` rather than `#NAME?`.
        engine.recalculate();
        assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Error(ErrorKind::NA));
        assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Error(ErrorKind::NA));
        assert_eq!(
            engine.get_cell_value("Sheet1", "A3"),
            Value::Error(ErrorKind::GettingData)
        );
    }
}

#[test]
fn engine_accepts_localized_external_data_r1c1_formulas_and_persists_canonical_a1() {
    for (locale, localized_cube, localized_err) in [
        (
            &locale::DE_DE,
            "=CUBEWERT(\"conn\";R[-5]C[-2];1,5)",
            "=#DATEN_ABRUFEN",
        ),
        (
            &locale::FR_FR,
            "=VALEUR.CUBE(\"conn\";R[-5]C[-2];1,5)",
            "=#OBTENTION_DONNEES",
        ),
        (
            &locale::ES_ES,
            "=VALOR.CUBO(\"conn\";R[-5]C[-2];1,5)",
            "=#OBTENIENDO_DATOS",
        ),
    ] {
        let mut engine = Engine::new();

        // Use an R1C1 reference in the function arguments to ensure the full pipeline
        // (localized separators + function-name translation + R1C1->A1 rewrite) works.
        engine
            .set_cell_formula_localized_r1c1(
                "Sheet1",
                "C5",
                "=RTD(\"my.server\";\"topic\";R[-4]C[-2];1,5)",
                locale,
            )
            .unwrap();
        assert_eq!(
            engine.get_cell_formula("Sheet1", "C5"),
            Some("=RTD(\"my.server\",\"topic\",A1,1.5)")
        );

        engine
            .set_cell_formula_localized_r1c1("Sheet1", "C6", localized_cube, locale)
            .unwrap();
        assert_eq!(
            engine.get_cell_formula("Sheet1", "C6"),
            Some("=CUBEVALUE(\"conn\",A1,1.5)")
        );

        engine
            .set_cell_formula_localized_r1c1("Sheet1", "A1", localized_err, locale)
            .unwrap();
        assert_eq!(engine.get_cell_formula("Sheet1", "A1"), Some("=#GETTING_DATA"));

        engine.recalculate();
        assert_eq!(engine.get_cell_value("Sheet1", "C5"), Value::Error(ErrorKind::NA));
        assert_eq!(engine.get_cell_value("Sheet1", "C6"), Value::Error(ErrorKind::NA));
        assert_eq!(
            engine.get_cell_value("Sheet1", "A1"),
            Value::Error(ErrorKind::GettingData)
        );
    }
}

#[test]
fn engine_accepts_localized_r1c1_formulas_and_persists_canonical_a1() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", 2.0).unwrap();

    engine
        .set_cell_formula_localized_r1c1(
            "Sheet1",
            "C5",
            "=SUMME(R[-4]C[-2];R[-4]C[-1])",
            &locale::DE_DE,
        )
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "C5"), Value::Number(3.0));
    assert_eq!(engine.get_cell_formula("Sheet1", "C5"), Some("=SUM(A1,B1)"));
}

#[test]
fn engine_accepts_localized_r1c1_formulas_with_field_access() {
    let mut engine = Engine::new();
    engine
        .set_cell_value(
            "Sheet1",
            "A1",
            Value::Record(RecordValue::with_fields(
                "Record",
                HashMap::from([("Price".to_string(), Value::Number(10.0))]),
            )),
        )
        .unwrap();

    // This exercises the full pipeline:
    // - localized separators + function name translation (de-DE: `SUMME`, `;`, `1,5`)
    // - R1C1 reference rewriting (`RC[-1]` in B1 -> `A1`)
    // - field access parsing after an R1C1 reference (`RC[-1].Price`)
    engine
        .set_cell_formula_localized_r1c1(
            "Sheet1",
            "B1",
            "=SUMME(RC[-1].Price;1,5)",
            &locale::DE_DE,
        )
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(11.5));
    assert_eq!(
        engine.get_cell_formula("Sheet1", "B1"),
        Some("=SUM(A1.Price,1.5)")
    );
}

#[test]
fn engine_accepts_localized_spilling_formulas() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula_localized("Sheet1", "A1", "=SEQUENZ(2;2)", &locale::DE_DE)
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_formula("Sheet1", "A1"),
        Some("=SEQUENCE(2,2)")
    );
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(4.0));

    assert_eq!(
        engine.spill_range("Sheet1", "A1"),
        Some((parse_a1("A1").unwrap(), parse_a1("B2").unwrap()))
    );

    let localized = locale::localize_formula(
        engine.get_cell_formula("Sheet1", "A1").unwrap(),
        &locale::DE_DE,
    )
    .unwrap();
    assert_eq!(localized, "=SEQUENZ(2;2)");
}
