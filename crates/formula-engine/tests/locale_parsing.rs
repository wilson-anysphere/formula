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
fn canonicalize_and_localize_unicode_case_insensitive_function_names_for_es_es() {
    // Spanish translation uses non-ASCII letters (Ñ); ensure we do Unicode-aware case-folding.
    for localized in ["=año(1)", "=Año(1)", "=AÑO(1)"] {
        let canonical = locale::canonicalize_formula(localized, &locale::ES_ES).unwrap();
        assert_eq!(canonical, "=YEAR(1)");
    }

    // Reverse translation should use the spelling from `src/locale/data/es-ES.tsv`.
    assert_eq!(
        locale::localize_formula("=year(1)", &locale::ES_ES).unwrap(),
        "=AÑO(1)"
    );
}

#[test]
fn canonicalize_and_localize_more_function_names_for_de_de() {
    fn assert_roundtrip(canonical: &str, localized: &str) {
        assert_eq!(
            locale::canonicalize_formula(localized, &locale::DE_DE).unwrap(),
            canonical
        );
        assert_eq!(
            locale::localize_formula(canonical, &locale::DE_DE).unwrap(),
            localized
        );
    }

    // Non-ASCII in localized name.
    assert_roundtrip("=COUNTIF(A1:A3,\">0\")", "=ZÄHLENWENN(A1:A3;\">0\")");

    // Ensure CONCAT (TEXTKETTE) and CONCATENATE (VERKETTEN) remain distinct.
    assert_roundtrip("=CONCAT(\"a\",\"b\")", "=TEXTKETTE(\"a\";\"b\")");
    assert_roundtrip("=CONCATENATE(\"a\",\"b\")", "=VERKETTEN(\"a\";\"b\")");

    assert_roundtrip(
        "=TEXTJOIN(\",\",TRUE,\"a\",\"b\")",
        "=TEXTVERKETTEN(\",\";WAHR;\"a\";\"b\")",
    );
    assert_roundtrip("=VLOOKUP(1,A1:B2,2,FALSE)", "=SVERWEIS(1;A1:B2;2;FALSCH)");
    assert_roundtrip("=HLOOKUP(1,A1:B2,2,FALSE)", "=WVERWEIS(1;A1:B2;2;FALSCH)");
    assert_roundtrip("=IFERROR(1/0,42)", "=WENNFEHLER(1/0;42)");

    // Dotted canonical function name; ensure `.` survives translation.
    assert_roundtrip("=NORM.S.DIST(0,TRUE)", "=NORM.S.VERT(0;WAHR)");

    // Modern `_xlfn.`-prefixed function.
    assert_roundtrip(
        "=_xlfn.XLOOKUP(\"a\",A1:A2,B1:B2)",
        "=_xlfn.XVERWEIS(\"a\";A1:A2;B1:B2)",
    );

    // Plain (non-`_xlfn.`) dynamic-array functions with localized German names.
    assert_roundtrip("=UNIQUE(A1:A3)", "=EINDEUTIG(A1:A3)");
    assert_roundtrip("=SORT(A1:A3)", "=SORTIEREN(A1:A3)");
    assert_roundtrip("=DROP(A1:A3,1)", "=WEGLASSEN(A1:A3;1)");
    assert_roundtrip("=TOCOL(A1:B2,1)", "=ZUSPALTE(A1:B2;1)");
    assert_roundtrip("=TOROW(A1:B2,1)", "=ZUZEILE(A1:B2;1)");
    assert_roundtrip("=XLOOKUP(1,A1:A3,B1:B3)", "=XVERWEIS(1;A1:A3;B1:B3)");
    assert_roundtrip(
        "=IMAGE(\"https://example.com\")",
        "=BILD(\"https://example.com\")",
    );

    // TRUE()/FALSE() as functions (not just boolean literals).
    assert_roundtrip("=TRUE()", "=WAHR()");
    assert_roundtrip("=FALSE()", "=FALSCH()");
}

#[test]
fn de_de_function_translation_table_covers_all_registered_functions() {
    use formula_engine::functions::FunctionSpec;
    let tsv = include_str!("../src/locale/data/de-DE.tsv");
    let mut canon = HashSet::<String>::new();
    let mut localized_to_canon = HashMap::<String, String>::new();
    let mut localized_collisions = Vec::<(String, String, String)>::new();

    for line in tsv.lines() {
        let line = line.trim();
        if line.is_empty() || line.starts_with('#') {
            continue;
        }
        let (canon_name, loc_name) = line.split_once('\t').unwrap_or_else(|| {
            panic!("invalid function translation line (expected TSV): {line:?}")
        });
        assert!(
            canon.insert(canon_name.to_string()),
            "duplicate canonical function translation entry: {canon_name}"
        );
        if let Some(prev) = localized_to_canon.insert(loc_name.to_string(), canon_name.to_string())
        {
            localized_collisions.push((loc_name.to_string(), prev, canon_name.to_string()));
        }
    }

    for spec in inventory::iter::<FunctionSpec> {
        let name = spec.name.to_ascii_uppercase();
        assert!(
            canon.contains(&name),
            "missing de-DE translation entry for function: {name}"
        );
    }

    assert!(
        localized_collisions.is_empty(),
        "de-DE function translations contain localized-name collisions: {localized_collisions:?}"
    );
}

#[test]
fn canonicalize_supports_thousands_and_leading_decimal_in_de_de() {
    let canonical = locale::canonicalize_formula("=SUMME(1.234,56;,5)", &locale::DE_DE).unwrap();
    assert_eq!(canonical, "=SUM(1234.56,.5)");
}

#[test]
fn canonicalize_and_localize_supports_thousands_grouping_in_es_es() {
    let localized = "=SUMA(1.234,56;0,5)";
    let canonical = locale::canonicalize_formula(localized, &locale::ES_ES).unwrap();
    assert_eq!(canonical, "=SUM(1234.56,0.5)");

    let localized_roundtrip = locale::localize_formula(&canonical, &locale::ES_ES).unwrap();
    assert_eq!(localized_roundtrip, localized);
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
fn localize_emits_locale_decimal_separator_for_canonical_leading_decimal() {
    let canonical = "=SUM(.5,1)";
    for (loc, expected_fn) in [
        (&locale::DE_DE, "SUMME"),
        (&locale::FR_FR, "SOMME"),
        (&locale::ES_ES, "SUMA"),
    ] {
        let localized = locale::localize_formula(canonical, loc).unwrap();

        assert!(
            localized.contains(expected_fn),
            "expected localized function name {expected_fn} in {localized:?}"
        );
        // Ensure the decimal separator was localized (no canonical `.5` remains).
        assert!(
            localized.contains(",5"),
            "expected localized decimal separator in {localized:?}"
        );
        assert!(
            !localized.contains(".5"),
            "unexpected canonical decimal separator in {localized:?}"
        );
        // Comma-decimal locales use `;` as argument separator; ensure we rewrite it as well.
        assert!(
            localized.contains(';'),
            "expected localized argument separator in {localized:?}"
        );

        let roundtrip = locale::canonicalize_formula(&localized, loc).unwrap();
        assert_eq!(roundtrip, canonical);
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
fn canonicalize_and_localize_array_literals_for_fr_fr_and_es_es() {
    for (locale, localized) in [
        (&locale::FR_FR, "=SOMME({1\\2;3\\4})"),
        (&locale::ES_ES, "=SUMA({1\\2;3\\4})"),
    ] {
        let canonical = locale::canonicalize_formula(localized, locale).unwrap();
        assert_eq!(canonical, "=SUM({1,2;3,4})");

        let localized_roundtrip = locale::localize_formula(&canonical, locale).unwrap();
        assert_eq!(localized_roundtrip, localized);
    }
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
fn canonicalize_and_localize_unions_for_fr_fr_and_es_es() {
    for locale in [&locale::FR_FR, &locale::ES_ES] {
        let localized = "(A1;B1)";
        let canonical = locale::canonicalize_formula(localized, locale).unwrap();
        assert_eq!(canonical, "(A1,B1)");

        let localized_roundtrip = locale::localize_formula(&canonical, locale).unwrap();
        assert_eq!(localized_roundtrip, localized);
    }
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
    assert_eq!(
        locale::localize_formula(&canon, &locale::FR_FR).unwrap(),
        fr
    );

    let es = "=_xlfn.VALOR.CUBO(\"conn\";\"member\";1,5)";
    let canon = locale::canonicalize_formula(es, &locale::ES_ES).unwrap();
    assert_eq!(canon, "=_xlfn.CUBEVALUE(\"conn\",\"member\",1.5)");
    assert_eq!(
        locale::localize_formula(&canon, &locale::ES_ES).unwrap(),
        es
    );
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
        assert_eq!(
            locale::localize_formula(canonical, locale).unwrap(),
            localized
        );
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
fn canonicalize_and_localize_supports_nbsp_thousands_separator_in_fr_fr() {
    // French Excel commonly uses NBSP (U+00A0) for thousands grouping.
    let fr = "=SOMME(1\u{00A0}234,56;0,5)";
    let canon = locale::canonicalize_formula(fr, &locale::FR_FR).unwrap();
    assert_eq!(canon, "=SUM(1234.56,0.5)");

    // When localizing the canonical form, the engine should re-insert thousands grouping for
    // readability. Accept either NBSP or narrow NBSP depending on the locale configuration.
    let roundtrip = locale::localize_formula(&canon, &locale::FR_FR).unwrap();
    assert!(
        roundtrip == fr || roundtrip == "=SOMME(1\u{202F}234,56;0,5)",
        "unexpected localized roundtrip: {roundtrip:?}"
    );
}

#[test]
fn canonicalize_supports_narrow_nbsp_thousands_separator_in_fr_fr() {
    // Some French locales/spreadsheets use narrow NBSP (U+202F) for thousands grouping.
    let fr = "=SOMME(1\u{202F}234,56;0,5)";
    let canon = locale::canonicalize_formula(fr, &locale::FR_FR).unwrap();
    assert_eq!(canon, "=SUM(1234.56,0.5)");
}

#[test]
fn canonicalize_supports_multiple_nbsp_thousands_separators_in_fr_fr() {
    let fr = "=SOMME(1\u{00A0}234\u{00A0}567,89;0,5)";
    let canon = locale::canonicalize_formula(fr, &locale::FR_FR).unwrap();
    assert_eq!(canon, "=SUM(1234567.89,0.5)");
}

#[test]
fn canonicalize_supports_mixed_nbsp_and_narrow_nbsp_thousands_separators_in_fr_fr() {
    let fr = "=SOMME(1\u{00A0}234\u{202F}567,89;0,5)";
    let canon = locale::canonicalize_formula(fr, &locale::FR_FR).unwrap();
    assert_eq!(canon, "=SUM(1234567.89,0.5)");
}

#[test]
fn localize_inserts_multiple_thousands_separators_in_fr_fr() {
    let canon = "=SUM(1234567.89,0.5)";
    let localized = locale::localize_formula(canon, &locale::FR_FR).unwrap();
    assert!(
        localized == "=SOMME(1\u{00A0}234\u{00A0}567,89;0,5)"
            || localized == "=SOMME(1\u{202F}234\u{202F}567,89;0,5)",
        "unexpected localized output: {localized:?}"
    );
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
    assert_roundtrip("=VLOOKUP(1,A1:B3,2,FALSE)", "=RECHERCHEV(1;A1:B3;2;FAUX)");
    assert_roundtrip("=HLOOKUP(1,A1:C2,2,FALSE)", "=RECHERCHEH(1;A1:C2;2;FAUX)");
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
        assert_eq!(
            locale::localize_formula(canonical, locale).unwrap(),
            localized
        );
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
        (
            "=SUM(Table1[[#Headers],[Qty]],1)",
            "Table1[[#Headers],[Qty]]",
        ),
        (
            "=SUM(Table1[[#This Row],[Qty]],1)",
            "Table1[[#This Row],[Qty]]",
        ),
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
fn structured_reference_items_are_not_translated_in_fr_fr() {
    // Structured-reference internals (`Table1[[...],[...]]`) must always stay canonical: we
    // localize only the surrounding formula syntax (function names + argument separators).
    let canonical = "=SUM(Table1[[#All],[Col1],[Col2]],1)";
    let localized = locale::localize_formula(canonical, &locale::FR_FR).unwrap();
    assert_eq!(localized, "=SOMME(Table1[[#All],[Col1],[Col2]];1)");

    let canonical_roundtrip = locale::canonicalize_formula(&localized, &locale::FR_FR).unwrap();
    assert_eq!(canonical_roundtrip, canonical);
}

#[test]
fn structured_reference_items_are_not_translated_in_es_es() {
    // Structured-reference internals (`Table1[[...],[...]]`) must always stay canonical: we
    // localize only the surrounding formula syntax (function names + argument separators).
    let canonical = "=SUM(Table1[[#All],[Col1],[Col2]],1)";
    let localized = locale::localize_formula(canonical, &locale::ES_ES).unwrap();
    assert_eq!(localized, "=SUMA(Table1[[#All],[Col1],[Col2]];1)");

    let canonical_roundtrip = locale::canonicalize_formula(&localized, &locale::ES_ES).unwrap();
    assert_eq!(canonical_roundtrip, canonical);
}

#[test]
fn structured_reference_separators_roundtrip_for_fr_fr_and_es_es() {
    // Regression test: fr-FR/es-ES use `;` as the function argument separator, but structured
    // reference contents (including `[[#Headers],[Qty]]` comma separators) must remain canonical.
    let canonical = "=SUM(Table1[[#Headers],[Qty]],1)";
    for (locale, expected_localized) in [
        (&locale::FR_FR, "=SOMME(Table1[[#Headers],[Qty]];1)"),
        (&locale::ES_ES, "=SUMA(Table1[[#Headers],[Qty]];1)"),
    ] {
        let localized = locale::localize_formula(canonical, locale).unwrap();
        assert_eq!(localized, expected_localized);

        let canonical_roundtrip = locale::canonicalize_formula(&localized, locale).unwrap();
        assert_eq!(canonical_roundtrip, canonical);
    }
}

#[test]
fn structured_reference_escaped_brackets_are_not_translated() {
    // Excel escapes `]` inside structured references as `]]` (e.g. column name `A]B` is written
    // as `A]]B`). Locale translation must preserve these escapes by treating `[...]` as opaque.
    let canonical = "=SUM(Table1[[#Headers],[A]]B]],1)";
    let localized = locale::localize_formula(canonical, &locale::DE_DE).unwrap();
    assert_eq!(localized, "=SUMME(Table1[[#Headers],[A]]B]];1)");

    let canonical_roundtrip = locale::canonicalize_formula(&localized, &locale::DE_DE).unwrap();
    assert_eq!(canonical_roundtrip, canonical);
}

#[test]
fn external_workbook_prefixes_inside_brackets_are_not_translated() {
    // External workbook references use `[...]` for the workbook name, and the sheet name follows
    // the closing bracket: `[Book.xlsx]Sheet1!A1`.
    //
    // Locale translation must never rewrite workbook/sheet identifiers inside references; only the
    // surrounding formula syntax is localized.
    let canonical = "=SUM([Book1.xlsx]Sheet1!A1,1)";
    let localized = locale::localize_formula(canonical, &locale::DE_DE).unwrap();
    assert_eq!(localized, "=SUMME([Book1.xlsx]Sheet1!A1;1)");

    let canonical_roundtrip = locale::canonicalize_formula(&localized, &locale::DE_DE).unwrap();
    assert_eq!(canonical_roundtrip, canonical);
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
fn locale_error_literal_maps_match_generated_error_tsvs() {
    fn assert_locale(locale: &locale::FormulaLocale, tsv: &str, label: &str) {
        let mut preferred_localized_by_canonical: std::collections::HashMap<String, String> =
            std::collections::HashMap::new();

        for (idx, raw_line) in tsv.lines().enumerate() {
            let line_no = idx + 1;
            let trimmed = raw_line.trim();
            // Error literals themselves start with `#`, so comments are `#` followed by whitespace.
            let is_comment = trimmed == "#"
                || (trimmed.starts_with('#')
                    && trimmed.chars().nth(1).is_some_and(|ch| ch.is_whitespace()));
            if trimmed.is_empty() || is_comment {
                continue;
            }

            let (canonical, localized) = raw_line
                .split_once('\t')
                .unwrap_or_else(|| panic!("invalid TSV line in {label}:{line_no}: {raw_line:?}"));
            let canonical = canonical.trim();
            let localized = localized.trim();
            assert!(
                !canonical.is_empty() && !localized.is_empty(),
                "invalid TSV line in {label}:{line_no} (empty field): {raw_line:?}"
            );

            // Error translation TSVs can include multiple localized spellings for the same canonical
            // literal. The locale registry uses the *first* spelling as the preferred display form
            // (canonical -> localized) and accepts *all* spellings (localized -> canonical).
            let preferred_localized = preferred_localized_by_canonical
                .entry(canonical.to_string())
                .or_insert_with(|| localized.to_string());
            assert_eq!(
                locale.localized_error_literal(canonical).unwrap_or(canonical),
                preferred_localized.as_str(),
                "canonical->localized error translation mismatch for {canonical} in {label}:{line_no}"
            );
            assert_eq!(
                locale.canonical_error_literal(localized).unwrap_or(localized),
                canonical,
                "localized->canonical error translation mismatch for {localized} in {label}:{line_no}"
            );
        }
    }

    assert_locale(
        &locale::DE_DE,
        include_str!("../src/locale/data/de-DE.errors.tsv"),
        "de-DE.errors.tsv",
    );
    assert_locale(
        &locale::FR_FR,
        include_str!("../src/locale/data/fr-FR.errors.tsv"),
        "fr-FR.errors.tsv",
    );
    assert_locale(
        &locale::ES_ES,
        include_str!("../src/locale/data/es-ES.errors.tsv"),
        "es-ES.errors.tsv",
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
fn canonicalize_and_localize_additional_function_translations_for_es_es() {
    fn assert_roundtrip(localized: &str, canonical: &str) {
        assert_eq!(
            locale::canonicalize_formula(localized, &locale::ES_ES).unwrap(),
            canonical
        );
        assert_eq!(
            locale::localize_formula(canonical, &locale::ES_ES).unwrap(),
            localized
        );
    }

    // Common translated worksheet functions.
    assert_roundtrip("=CONTAR.SI(A1:A10;\">0\")", "=COUNTIF(A1:A10,\">0\")");
    assert_roundtrip("=BUSCARV(1;A1:B3;2;FALSO)", "=VLOOKUP(1,A1:B3,2,FALSE)");
    assert_roundtrip("=BUSCARH(1;A1:C2;2;VERDADERO)", "=HLOOKUP(1,A1:C2,2,TRUE)");
    assert_roundtrip(
        "=SUMAR.SI(A1:A10;\">0\";B1:B10)",
        "=SUMIF(A1:A10,\">0\",B1:B10)",
    );
    assert_roundtrip(
        "=SUMAR.SI.CONJUNTO(B1:B10;A1:A10;\">0\";A1:A10;\"<10\")",
        "=SUMIFS(B1:B10,A1:A10,\">0\",A1:A10,\"<10\")",
    );
    assert_roundtrip("=PROMEDIO.SI(A1:A10;\">0\")", "=AVERAGEIF(A1:A10,\">0\")");
    assert_roundtrip(
        "=PROMEDIO.SI.CONJUNTO(B1:B10;A1:A10;\">0\";A1:A10;\"<10\")",
        "=AVERAGEIFS(B1:B10,A1:A10,\">0\",A1:A10,\"<10\")",
    );
    assert_roundtrip("=SI.ERROR(1/0;0)", "=IFERROR(1/0,0)");
    assert_roundtrip(
        "=SI.ND(BUSCARV(1;A1:B2;2;FALSO);0)",
        "=IFNA(VLOOKUP(1,A1:B2,2,FALSE),0)",
    );
    assert_roundtrip(
        "=SI.CONJUNTO(1=1;\"A\";1=2;\"B\")",
        "=IFS(1=1,\"A\",1=2,\"B\")",
    );
    assert_roundtrip("=INDICE(A1:B2;2;1)", "=INDEX(A1:B2,2,1)");
    assert_roundtrip("=COINCIDIR(5;A1:A10;0)", "=MATCH(5,A1:A10,0)");
    assert_roundtrip("=DESREF(A1;1;1)", "=OFFSET(A1,1,1)");
    assert_roundtrip("=INDIRECTO(\"A1\")", "=INDIRECT(\"A1\")");
    assert_roundtrip("=HOY()", "=TODAY()");
    assert_roundtrip(
        "=FECHANUMERO(\"2020-01-01\")",
        "=DATEVALUE(\"2020-01-01\")",
    );
    assert_roundtrip("=HORANUMERO(\"1:00\")", "=TIMEVALUE(\"1:00\")");
    assert_roundtrip("=EXACTO(\"a\";\"b\")", "=EXACT(\"a\",\"b\")");
    assert_roundtrip("=IMAGEN(\"x\")", "=IMAGE(\"x\")");
    assert_roundtrip("=MINUTO(0)", "=MINUTE(0)");
    assert_roundtrip("=RESIDUO(5;2)", "=MOD(5,2)");

    // TRUE()/FALSE() as zero-argument functions (not just boolean literals).
    assert_roundtrip("=VERDADERO()", "=TRUE()");
    assert_roundtrip("=FALSO()", "=FALSE()");

    // `_xlfn.`-prefixed functions should translate their base name.
    assert_roundtrip(
        "=_xlfn.BUSCARX(1;A1:A3;B1:B3)",
        "=_xlfn.XLOOKUP(1,A1:A3,B1:B3)",
    );
}

#[test]
fn true_false_functions_are_case_insensitive() {
    for (locale, true_variants, false_variants, localized_true, localized_false) in [
        (
            &locale::DE_DE,
            ["=wahr()", "=Wahr()", "=WAHR()"],
            ["=falsch()", "=Falsch()", "=FALSCH()"],
            "=WAHR()",
            "=FALSCH()",
        ),
        (
            &locale::FR_FR,
            ["=vrai()", "=Vrai()", "=VRAI()"],
            ["=faux()", "=Faux()", "=FAUX()"],
            "=VRAI()",
            "=FAUX()",
        ),
        (
            &locale::ES_ES,
            ["=verdadero()", "=Verdadero()", "=VERDADERO()"],
            ["=falso()", "=Falso()", "=FALSO()"],
            "=VERDADERO()",
            "=FALSO()",
        ),
    ] {
        for src in true_variants {
            assert_eq!(
                locale::canonicalize_formula(src, locale).unwrap(),
                "=TRUE()"
            );
        }
        for src in false_variants {
            assert_eq!(
                locale::canonicalize_formula(src, locale).unwrap(),
                "=FALSE()"
            );
        }

        // Localization should also accept canonical function names case-insensitively and emit the
        // normalized spelling from the locale TSV.
        assert_eq!(
            locale::localize_formula("=true()", locale).unwrap(),
            localized_true
        );
        assert_eq!(
            locale::localize_formula("=False()", locale).unwrap(),
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
fn localized_boolean_keywords_are_not_translated_in_field_access() {
    // Locales use translated boolean keywords (e.g. `WAHR`, `VRAI`, `VERDADERO`), but those should
    // still be allowed as record field names in dot-access position.
    assert_eq!(
        locale::canonicalize_formula("=A1.WAHR", &locale::DE_DE).unwrap(),
        "=A1.WAHR"
    );
    assert_eq!(
        locale::canonicalize_formula("=A1.VRAI", &locale::FR_FR).unwrap(),
        "=A1.VRAI"
    );
    assert_eq!(
        locale::canonicalize_formula("=A1.VERDADERO", &locale::ES_ES).unwrap(),
        "=A1.VERDADERO"
    );

    // Control: standalone boolean literals should still be translated.
    assert_eq!(
        locale::canonicalize_formula("=WENN(WAHR;1;0)", &locale::DE_DE).unwrap(),
        "=IF(TRUE,1,0)"
    );
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
fn field_access_selectors_are_not_translated() {
    assert_eq!(
        locale::canonicalize_formula("=A1.WAHR", &locale::DE_DE).unwrap(),
        "=A1.WAHR"
    );
    assert_eq!(
        locale::localize_formula("=A1.TRUE", &locale::DE_DE).unwrap(),
        "=A1.TRUE"
    );
    assert_eq!(
        locale::canonicalize_formula("=A1.SUMME(1;2)", &locale::DE_DE).unwrap(),
        "=A1.SUMME(1,2)"
    );
    assert_eq!(
        locale::localize_formula("=A1.SUM(1,2)", &locale::DE_DE).unwrap(),
        "=A1.SUM(1;2)"
    );
}

#[test]
fn canonicalize_and_localize_error_literals_for_all_locales() {
    struct Case<'a> {
        kind: ErrorKind,
        localized_preferred: &'a str,
        localized_variants: &'a [&'a str],
    }
    fn assert_error_round_trip(locale: &locale::FormulaLocale, case: &Case<'_>) {
        let canonical = case.kind.as_code();
        // localized -> canonical
        for &localized in case.localized_variants {
            let src = format!("={localized}");
            let canon = locale::canonicalize_formula(&src, locale).unwrap();
            assert_eq!(canon, format!("={canonical}"));
        }

        // canonical -> localized (preferred)
        let src = format!("={canonical}");
        let localized = locale::localize_formula(&src, locale).unwrap();
        assert_eq!(localized, format!("={}", case.localized_preferred));

        // `#N/A!` is an accepted alias for `#N/A`; ensure it localizes to the preferred form.
        if canonical.eq_ignore_ascii_case("#N/A") {
            let localized = locale::localize_formula("=#N/A!", locale).unwrap();
            assert_eq!(localized, format!("={}", case.localized_preferred));
        }
    }

    // de-DE
    for case in [
        Case {
            kind: ErrorKind::Null,
            localized_preferred: "#NULL!",
            localized_variants: &["#NULL!"],
        },
        Case {
            kind: ErrorKind::Div0,
            localized_preferred: "#DIV/0!",
            localized_variants: &["#DIV/0!"],
        },
        Case {
            kind: ErrorKind::Value,
            localized_preferred: "#WERT!",
            localized_variants: &["#WERT!"],
        },
        Case {
            kind: ErrorKind::Ref,
            localized_preferred: "#BEZUG!",
            localized_variants: &["#BEZUG!"],
        },
        Case {
            kind: ErrorKind::Name,
            localized_preferred: "#NAME?",
            localized_variants: &["#NAME?"],
        },
        Case {
            kind: ErrorKind::Num,
            localized_preferred: "#ZAHL!",
            localized_variants: &["#ZAHL!"],
        },
        Case {
            kind: ErrorKind::NA,
            localized_preferred: "#NV",
            localized_variants: &["#NV", "#N/A", "#N/A!"],
        },
        Case {
            kind: ErrorKind::GettingData,
            localized_preferred: "#DATEN_ABRUFEN",
            localized_variants: &["#DATEN_ABRUFEN"],
        },
        Case {
            kind: ErrorKind::Spill,
            localized_preferred: "#ÜBERLAUF!",
            localized_variants: &["#ÜBERLAUF!", "#Überlauf!", "#üBeRlAuF!"],
        },
        Case {
            kind: ErrorKind::Calc,
            localized_preferred: "#KALK!",
            localized_variants: &["#KALK!", "#CALC!"],
        },
        Case {
            kind: ErrorKind::Field,
            localized_preferred: "#FIELD!",
            localized_variants: &["#FIELD!"],
        },
        Case {
            kind: ErrorKind::Connect,
            localized_preferred: "#CONNECT!",
            localized_variants: &["#CONNECT!"],
        },
        Case {
            kind: ErrorKind::Blocked,
            localized_preferred: "#BLOCKED!",
            localized_variants: &["#BLOCKED!"],
        },
        Case {
            kind: ErrorKind::Unknown,
            localized_preferred: "#UNKNOWN!",
            localized_variants: &["#UNKNOWN!"],
        },
    ] {
        assert_error_round_trip(&locale::DE_DE, &case);
    }

    // fr-FR
    for case in [
        Case {
            kind: ErrorKind::Null,
            localized_preferred: "#NUL!",
            localized_variants: &["#NUL!", "#NULL!"],
        },
        Case {
            kind: ErrorKind::Div0,
            localized_preferred: "#DIV/0!",
            localized_variants: &["#DIV/0!"],
        },
        Case {
            kind: ErrorKind::Value,
            localized_preferred: "#VALEUR!",
            localized_variants: &["#VALEUR!"],
        },
        Case {
            kind: ErrorKind::Ref,
            localized_preferred: "#REF!",
            localized_variants: &["#REF!"],
        },
        Case {
            kind: ErrorKind::Name,
            localized_preferred: "#NOM?",
            localized_variants: &["#NOM?"],
        },
        Case {
            kind: ErrorKind::Num,
            localized_preferred: "#NOMBRE!",
            localized_variants: &["#NOMBRE!"],
        },
        Case {
            kind: ErrorKind::NA,
            localized_preferred: "#N/A",
            localized_variants: &["#N/A", "#N/A!"],
        },
        Case {
            kind: ErrorKind::GettingData,
            localized_preferred: "#OBTENTION_DONNEES",
            localized_variants: &["#OBTENTION_DONNEES"],
        },
        Case {
            kind: ErrorKind::Spill,
            localized_preferred: "#PROPAGATION!",
            localized_variants: &["#PROPAGATION!", "#DEVERSEMENT!"],
        },
        Case {
            kind: ErrorKind::Calc,
            localized_preferred: "#CALCUL!",
            localized_variants: &["#CALCUL!", "#CALC!"],
        },
        Case {
            kind: ErrorKind::Field,
            localized_preferred: "#CHAMP!",
            localized_variants: &["#CHAMP!", "#FIELD!"],
        },
        Case {
            kind: ErrorKind::Connect,
            localized_preferred: "#CONNEXION!",
            localized_variants: &["#CONNEXION!", "#CONNECT!"],
        },
        Case {
            kind: ErrorKind::Blocked,
            localized_preferred: "#BLOQUE!",
            localized_variants: &["#BLOQUE!", "#BLOCKED!"],
        },
        Case {
            kind: ErrorKind::Unknown,
            localized_preferred: "#INCONNU!",
            localized_variants: &["#INCONNU!", "#UNKNOWN!"],
        },
    ] {
        assert_error_round_trip(&locale::FR_FR, &case);
    }

    // es-ES (includes inverted punctuation variants)
    for case in [
        Case {
            kind: ErrorKind::Null,
            localized_preferred: "#¡NULO!",
            localized_variants: &["#¡NULO!", "#NULO!"],
        },
        Case {
            kind: ErrorKind::Div0,
            localized_preferred: "#¡DIV/0!",
            localized_variants: &["#¡DIV/0!", "#DIV/0!"],
        },
        Case {
            kind: ErrorKind::Value,
            localized_preferred: "#¡VALOR!",
            localized_variants: &["#¡VALOR!", "#VALOR!"],
        },
        Case {
            kind: ErrorKind::Ref,
            localized_preferred: "#¡REF!",
            localized_variants: &["#¡REF!", "#REF!"],
        },
        Case {
            kind: ErrorKind::Name,
            localized_preferred: "#¿NOMBRE?",
            localized_variants: &["#¿NOMBRE?", "#NOMBRE?"],
        },
        Case {
            kind: ErrorKind::Num,
            localized_preferred: "#¡NUM!",
            localized_variants: &["#¡NUM!", "#NUM!"],
        },
        Case {
            kind: ErrorKind::NA,
            localized_preferred: "#N/A",
            localized_variants: &["#N/A", "#N/A!"],
        },
        Case {
            kind: ErrorKind::GettingData,
            localized_preferred: "#OBTENIENDO_DATOS",
            localized_variants: &["#OBTENIENDO_DATOS"],
        },
        Case {
            kind: ErrorKind::Spill,
            localized_preferred: "#¡DESBORDAMIENTO!",
            localized_variants: &["#¡DESBORDAMIENTO!", "#DESBORDAMIENTO!"],
        },
        Case {
            kind: ErrorKind::Calc,
            localized_preferred: "#¡CALC!",
            localized_variants: &["#¡CALC!", "#CALC!"],
        },
        Case {
            kind: ErrorKind::Field,
            localized_preferred: "#¡CAMPO!",
            localized_variants: &["#¡CAMPO!", "#CAMPO!"],
        },
        Case {
            kind: ErrorKind::Connect,
            localized_preferred: "#¡CONECTAR!",
            localized_variants: &["#¡CONECTAR!", "#CONECTAR!"],
        },
        Case {
            kind: ErrorKind::Blocked,
            localized_preferred: "#¡BLOQUEADO!",
            localized_variants: &["#¡BLOQUEADO!", "#BLOQUEADO!"],
        },
        Case {
            kind: ErrorKind::Unknown,
            localized_preferred: "#¡DESCONOCIDO!",
            localized_variants: &["#¡DESCONOCIDO!", "#DESCONOCIDO!"],
        },
    ] {
        assert_error_round_trip(&locale::ES_ES, &case);
    }

    let es_spill_variants = [
        "=#¡DESBORDAMIENTO!",
        "=#¡desbordamiento!",
        "=#¡DeSbOrDaMiEnTo!",
    ];
    for src in es_spill_variants {
        let canon = locale::canonicalize_formula(src, &locale::ES_ES).unwrap();
        assert_eq!(canon, "=#SPILL!");
        assert_eq!(
            locale::localize_formula(&canon, &locale::ES_ES).unwrap(),
            "=#¡DESBORDAMIENTO!"
        );
    }
}

#[test]
fn canonicalize_and_localize_inverted_punctuation_error_literals_for_es_es() {
    for (localized_variants, canonical) in [
        (&["=#¡VALOR!", "=#¡valor!"][..], "=#VALUE!"),
        (&["=#¿NOMBRE?", "=#¿nombre?"][..], "=#NAME?"),
    ] {
        for &localized in localized_variants {
            let canon = locale::canonicalize_formula(localized, &locale::ES_ES).unwrap();
            assert_eq!(canon, canonical);
            // Ensure canonical -> localized prefers the inverted punctuation spelling.
            assert_eq!(
                locale::localize_formula(&canon, &locale::ES_ES).unwrap(),
                localized_variants[0]
            );
        }
    }
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
// `src/locale/data/*.tsv` and `src/locale/data/*.errors.tsv`.
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
}

#[test]
fn canonicalize_and_localize_all_cube_function_names() {
    fn assert_roundtrip(locale: &locale::FormulaLocale, canonical: &str, localized: &str) {
        assert_eq!(
            locale::canonicalize_formula(localized, locale).unwrap(),
            canonical
        );
        assert_eq!(
            locale::localize_formula(canonical, locale).unwrap(),
            localized
        );
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
        "=CUBEELEMENT(\"conn\";\"member\";\"caption\")",
    );
    assert_roundtrip(
        &locale::DE_DE,
        "=CUBEMEMBERPROPERTY(\"conn\",\"member\",\"prop\")",
        "=CUBEELEMENTEIGENSCHAFT(\"conn\";\"member\";\"prop\")",
    );
    assert_roundtrip(
        &locale::DE_DE,
        "=CUBERANKEDMEMBER(\"conn\",\"set\",3,\"caption\")",
        "=CUBERANGELEMENT(\"conn\";\"set\";3;\"caption\")",
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
        "=CUBEKPIELEMENT(\"conn\";\"kpi\";\"property\";\"caption\")",
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
        let canon =
            locale::canonicalize_formula_with_style(localized, locale, ReferenceStyle::R1C1)
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
    assert_eq!(
        locale::localize_formula(&canon, &locale::DE_DE).unwrap(),
        de
    );

    let fr = "=SOMME(VALEUR.CUBE!A1;1)";
    let canon = locale::canonicalize_formula(fr, &locale::FR_FR).unwrap();
    assert_eq!(canon, "=SUM(VALEUR.CUBE!A1,1)");
    assert_eq!(
        locale::localize_formula(&canon, &locale::FR_FR).unwrap(),
        fr
    );

    let es = "=SUMA(VALOR.CUBO!A1;1)";
    let canon = locale::canonicalize_formula(es, &locale::ES_ES).unwrap();
    assert_eq!(canon, "=SUM(VALOR.CUBO!A1,1)");
    assert_eq!(
        locale::localize_formula(&canon, &locale::ES_ES).unwrap(),
        es
    );
}

#[test]
fn localized_external_data_function_names_are_not_translated_when_used_as_identifiers() {
    // Function-name translation should only happen for identifiers used in function-call position
    // (`NAME(`). If a workbook has a defined name that happens to match a localized spelling, it
    // must not be rewritten.
    let de = "=CUBEWERT+1";
    let canon = locale::canonicalize_formula(de, &locale::DE_DE).unwrap();
    assert_eq!(canon, de);
    assert_eq!(
        locale::localize_formula(&canon, &locale::DE_DE).unwrap(),
        de
    );

    let fr = "=VALEUR.CUBE+1";
    let canon = locale::canonicalize_formula(fr, &locale::FR_FR).unwrap();
    assert_eq!(canon, fr);
    assert_eq!(
        locale::localize_formula(&canon, &locale::FR_FR).unwrap(),
        fr
    );

    let es = "=VALOR.CUBO+1";
    let canon = locale::canonicalize_formula(es, &locale::ES_ES).unwrap();
    assert_eq!(canon, es);
    assert_eq!(
        locale::localize_formula(&canon, &locale::ES_ES).unwrap(),
        es
    );
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
        assert_eq!(
            engine.get_cell_value("Sheet1", "A1"),
            Value::Error(ErrorKind::NA)
        );
        assert_eq!(
            engine.get_cell_value("Sheet1", "A2"),
            Value::Error(ErrorKind::NA)
        );
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
        assert_eq!(
            engine.get_cell_formula("Sheet1", "A1"),
            Some("=#GETTING_DATA")
        );

        engine.recalculate();
        assert_eq!(
            engine.get_cell_value("Sheet1", "C5"),
            Value::Error(ErrorKind::NA)
        );
        assert_eq!(
            engine.get_cell_value("Sheet1", "C6"),
            Value::Error(ErrorKind::NA)
        );
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
        .set_cell_formula_localized_r1c1("Sheet1", "B1", "=SUMME(RC[-1].Price;1,5)", &locale::DE_DE)
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
