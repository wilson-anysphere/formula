use std::collections::HashMap;
use std::fs;
use std::io::Write;
use std::path::{Path, PathBuf};

use calamine::{open_workbook_auto, Reader};
use formula_xlsb::format::{format_a1, format_hex};
use formula_xlsb::rgce::decode_rgce;
use formula_xlsb::{Formula, XlsbWorkbook};
use pretty_assertions::assert_eq;

mod fixture_builder;
use fixture_builder::XlsbFixtureBuilder;

type CellCoord = (u32, u32);
type SheetFormulas<T> = HashMap<String, HashMap<CellCoord, T>>;

#[test]
fn formulas_match_calamine_for_all_fixtures() {
    let fixtures_dir = Path::new(concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures"));
    let mut fixtures = discover_xlsb_fixtures(fixtures_dir);
    fixtures.sort();

    assert!(
        !fixtures.is_empty(),
        "no .xlsb fixtures found under {}",
        fixtures_dir.display()
    );

    for fixture in fixtures {
        compare_fixture(&fixture);
    }
}

#[test]
fn formulas_match_calamine_for_generated_fixture() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_sheet_name("Sheet1");

    builder.set_cell_number(0, 0, 1.0);
    builder.set_cell_number(0, 1, 2.0);
    builder.set_cell_formula_num(
        0,
        2,
        3.0,
        // `A1+B1`
        vec![
            0x24, 0, 0, 0, 0, 0x00, 0xC0, // A1
            0x24, 0, 0, 0, 0, 0x01, 0xC0, // B1
            0x03, // +
        ],
        Vec::new(),
    );

    builder.set_cell_formula_num(
        1,
        0,
        -3.0,
        // `-(A1+B1)`
        vec![
            0x24, 0, 0, 0, 0, 0x00, 0xC0, // A1
            0x24, 0, 0, 0, 0, 0x01, 0xC0, // B1
            0x03, // +
            0x15, // (..)
            0x13, // unary -
        ],
        Vec::new(),
    );

    builder.set_cell_formula_num(
        1,
        1,
        3.0,
        // `1+2`
        vec![0x1E, 0x01, 0x00, 0x1E, 0x02, 0x00, 0x03],
        Vec::new(),
    );

    builder.set_cell_formula_num(
        1,
        2,
        0.02,
        // `2%`
        vec![0x1E, 0x02, 0x00, 0x14],
        Vec::new(),
    );

    builder.set_cell_formula_bool(
        2,
        0,
        true,
        // `1=1`
        vec![0x1E, 0x01, 0x00, 0x1E, 0x01, 0x00, 0x0B],
    );

    builder.set_cell_formula_err(
        2,
        1,
        0x07, // #DIV/0!
        // `1/0`
        vec![0x1E, 0x01, 0x00, 0x1E, 0x00, 0x00, 0x06],
    );

    builder.set_cell_formula_str(
        2,
        2,
        "AB",
        // `"A"&"B"`
        vec![
            0x17, 0x01, 0x00, 0x41, 0x00, // "A"
            0x17, 0x01, 0x00, 0x42, 0x00, // "B"
            0x08, // &
        ],
    );

    builder.set_cell_formula_bool(
        3,
        0,
        true,
        // `TRUE`
        vec![0x1D, 0x01],
    );

    let mut ptg_num_plus_int = Vec::new();
    ptg_num_plus_int.push(0x1F);
    ptg_num_plus_int.extend_from_slice(&1.5f64.to_le_bytes());
    ptg_num_plus_int.extend_from_slice(&[0x1E, 0x02, 0x00, 0x03]); // 2, +
    builder.set_cell_formula_num(3, 1, 3.5, ptg_num_plus_int, Vec::new());

    builder.set_cell_formula_err(
        3,
        2,
        0x00, // #NULL!
        // `A1 B1` (intersection)
        vec![
            0x24, 0, 0, 0, 0, 0x00, 0xC0, // A1
            0x24, 0, 0, 0, 0, 0x01, 0xC0, // B1
            0x0F, // intersection
        ],
    );

    builder.set_cell_formula_num(
        3,
        3,
        3.0,
        // `$A$1+$B1`
        vec![
            0x24, 0, 0, 0, 0, 0x00, 0x00, // $A$1
            0x24, 0, 0, 0, 0, 0x01, 0x40, // $B1
            0x03, // +
        ],
        Vec::new(),
    );

    builder.set_cell_formula_str(
        4,
        0,
        "A\"B",
        // `"A""B"` (string containing a quote)
        vec![
            0x17, 0x04, 0x00, 0x41, 0x00, 0x22, 0x00, 0x22, 0x00, 0x42, 0x00, // "A\"\"B"
        ],
    );

    builder.set_cell_formula_err(
        4,
        1,
        0x2A, // #N/A
        // `#N/A` (PtgErr constant)
        vec![0x1C, 0x2A],
    );

    builder.set_cell_formula_bool(
        4,
        2,
        false,
        // `FALSE`
        vec![0x1D, 0x00],
    );

    builder.set_cell_formula_num(
        4,
        3,
        1.0,
        // `+1`
        vec![0x1E, 0x01, 0x00, 0x12],
        Vec::new(),
    );

    let bytes = builder.build_bytes();
    let mut tmp = tempfile::Builder::new()
        .prefix("formula_xlsb_generated_")
        .suffix(".xlsb")
        .tempfile()
        .expect("create temp xlsb");
    tmp.write_all(&bytes).expect("write temp xlsb");

    compare_fixture(tmp.path());
}

fn discover_xlsb_fixtures(fixtures_dir: &Path) -> Vec<PathBuf> {
    let mut out = Vec::new();
    discover_xlsb_fixtures_inner(fixtures_dir, &mut out);
    out
}

fn discover_xlsb_fixtures_inner(dir: &Path, out: &mut Vec<PathBuf>) {
    let entries = fs::read_dir(dir)
        .unwrap_or_else(|err| panic!("read fixtures dir {}: {err}", dir.display()));

    for entry in entries {
        let entry = entry.unwrap_or_else(|err| panic!("read fixtures dir entry in {}: {err}", dir.display()));
        let path = entry.path();
        if path.is_dir() {
            discover_xlsb_fixtures_inner(&path, out);
            continue;
        }

        let is_xlsb = path
            .extension()
            .and_then(|ext| ext.to_str())
            .map(|ext| ext.eq_ignore_ascii_case("xlsb"))
            .unwrap_or(false);
        if is_xlsb {
            out.push(path);
        }
    }
}

fn compare_fixture(path: &Path) {
    let calamine = read_calamine_formulas(path);
    let xlsb = read_formula_xlsb_formulas(path);

    for (sheet_name, cal_sheet) in calamine {
        let Some(xlsb_sheet) = xlsb.get(&sheet_name) else {
            panic!(
                "fixture {}: sheet {sheet_name:?} present in calamine but missing in formula-xlsb",
                path.display()
            );
        };

        for ((row, col), cal_formula) in cal_sheet {
            let Some(decoded) = xlsb_sheet.get(&(row, col)) else {
                panic!(
                    "fixture {}: sheet {sheet_name:?} cell {} has a formula per calamine but formula-xlsb did not report one",
                    path.display(),
                    format_a1(row, col)
                );
            };

            let decoded_text = match decoded.text.as_deref() {
                Some(text) => text,
                None => match decode_rgce(&decoded.rgce) {
                    Ok(unexpected) => {
                        panic!(
                            "fixture {}: sheet {sheet_name:?} cell {}: formula_xlsb returned Formula::text=None, but decode_rgce succeeded ({unexpected:?}); rgce={}",
                            path.display(),
                            format_a1(row, col),
                            format_hex(&decoded.rgce)
                        );
                    }
                    Err(err) => {
                        panic!(
                            "fixture {}: sheet {sheet_name:?} cell {}: formula-xlsb could not decode rgce (expected formula {cal_formula:?}): {err}; rgce={}",
                            path.display(),
                            format_a1(row, col),
                            format_hex(&decoded.rgce)
                        );
                    }
                },
            };

            let cal_norm = normalize_formula_for_compare(&cal_formula);
            let decoded_norm = normalize_formula_for_compare(decoded_text);
            assert_eq!(
                decoded_norm,
                cal_norm,
                "fixture {} sheet {sheet_name} cell {}\ncalamine: {cal_formula}\nformula-xlsb: {decoded_text}\nrgce={}",
                path.display(),
                format_a1(row, col),
                format_hex(&decoded.rgce)
            );
        }
    }
}

fn read_calamine_formulas(path: &Path) -> SheetFormulas<String> {
    let mut workbook = open_workbook_auto(path)
        .unwrap_or_else(|err| panic!("fixture {}: calamine open_workbook_auto failed: {err}", path.display()));

    let sheet_names = workbook.sheet_names().to_owned();
    let mut out: SheetFormulas<String> = HashMap::new();

    for sheet_name in sheet_names {
        let formula_range = match workbook.worksheet_formula(&sheet_name) {
            Ok(range) => range,
            Err(err) => {
                // Some workbooks/sheets may have no formula records. Treat that as an empty map so
                // fixtures without formulas can still be included.
                eprintln!(
                    "fixture {}: calamine failed to read formulas for sheet {sheet_name:?}: {err}; treating as no formulas",
                    path.display()
                );
                out.insert(sheet_name, HashMap::new());
                continue;
            }
        };

        let start = formula_range.start().unwrap_or((0, 0));
        let mut formulas = HashMap::new();

        for (row, col, formula) in formula_range.used_cells() {
            if formula.trim().is_empty() {
                continue;
            }

            let row: u32 = row
                .try_into()
                .unwrap_or_else(|_| panic!("fixture {}: row index overflow", path.display()));
            let col: u32 = col
                .try_into()
                .unwrap_or_else(|_| panic!("fixture {}: col index overflow", path.display()));

            let row = start
                .0
                .checked_add(row)
                .unwrap_or_else(|| panic!("fixture {}: row index overflow", path.display()));
            let col = start
                .1
                .checked_add(col)
                .unwrap_or_else(|| panic!("fixture {}: col index overflow", path.display()));

            formulas.insert((row, col), formula.clone());
        }

        out.insert(sheet_name, formulas);
    }

    out
}

fn read_formula_xlsb_formulas(path: &Path) -> SheetFormulas<Formula> {
    let wb = XlsbWorkbook::open(path)
        .unwrap_or_else(|err| panic!("fixture {}: formula-xlsb open failed: {err}", path.display()));

    let mut out: SheetFormulas<Formula> = HashMap::new();
    for (sheet_idx, sheet_meta) in wb.sheet_metas().iter().enumerate() {
        let sheet_name = sheet_meta.name.clone();
        let mut formulas: HashMap<CellCoord, Formula> = HashMap::new();

        wb.for_each_cell(sheet_idx, |cell| {
            if let Some(formula) = cell.formula {
                formulas.insert((cell.row, cell.col), formula);
            }
        })
        .unwrap_or_else(|err| {
            panic!(
                "fixture {}: formula-xlsb failed to stream sheet {sheet_name:?}: {err}",
                path.display()
            )
        });

        out.insert(sheet_name, formulas);
    }

    out
}

fn normalize_formula_for_compare(formula: &str) -> String {
    let trimmed = formula.trim();
    let trimmed = trimmed.strip_prefix('=').unwrap_or(trimmed).trim();

    let mut out = String::with_capacity(trimmed.len() + 1);
    let mut in_string = false;
    let mut in_quoted_ident = false;
    let mut pending_ws = false;
    let mut prev_emitted: Option<char> = None;
    let mut chars = trimmed.chars().peekable();

    while let Some(ch) = chars.next() {
        if ch == '"' && !in_quoted_ident {
            if pending_ws && !in_string {
                if should_keep_space(prev_emitted, ch) {
                    out.push(' ');
                }
                pending_ws = false;
            }
            out.push(ch);
            prev_emitted = Some(ch);
            if in_string {
                if chars.peek() == Some(&'"') {
                    // Escaped quote inside a string literal.
                    out.push('"');
                    chars.next();
                    prev_emitted = Some('"');
                } else {
                    in_string = false;
                }
            } else {
                in_string = true;
            }
            continue;
        }

        if ch == '\'' && !in_string {
            if pending_ws && !in_quoted_ident {
                if should_keep_space(prev_emitted, ch) {
                    out.push(' ');
                }
                pending_ws = false;
            }
            out.push(ch);
            prev_emitted = Some(ch);
            if in_quoted_ident {
                if chars.peek() == Some(&'\'') {
                    // Escaped quote inside a quoted sheet/workbook identifier.
                    out.push('\'');
                    chars.next();
                    prev_emitted = Some('\'');
                } else {
                    in_quoted_ident = false;
                }
            } else {
                in_quoted_ident = true;
            }
            continue;
        }

        if !in_string && !in_quoted_ident && ch.is_whitespace() {
            pending_ws = true;
            continue;
        }

        if pending_ws && !in_string && !in_quoted_ident {
            if should_keep_space(prev_emitted, ch) {
                out.push(' ');
            }
            pending_ws = false;
        }

        if in_string || in_quoted_ident {
            out.push(ch);
            prev_emitted = Some(ch);
        } else {
            let ch = ch.to_ascii_uppercase();
            out.push(ch);
            prev_emitted = Some(ch);
        }
    }

    format!("={out}")
}

fn should_keep_space(prev: Option<char>, next: char) -> bool {
    let Some(prev) = prev else {
        return false;
    };
    !is_space_insensitive_delimiter(prev) && !is_space_insensitive_delimiter(next)
}

fn is_space_insensitive_delimiter(ch: char) -> bool {
    matches!(
        ch,
        '(' | ')' | ',' | '+' | '-' | '*' | '/' | '^' | '&' | '=' | '<' | '>' | ':' | '!' | '%' | '{' | '}' | '[' | ']' | '@'
    )
}

#[test]
fn normalize_formula_preserves_intersection_space() {
    assert_eq!(normalize_formula_for_compare("=A1   B1"), "=A1 B1");
}

#[test]
fn normalize_formula_drops_trivial_whitespace() {
    assert_eq!(normalize_formula_for_compare("= A1 +  B1 "), "=A1+B1");
    assert_eq!(normalize_formula_for_compare("=A1 , B1"), "=A1,B1");
    assert_eq!(normalize_formula_for_compare("=A1 : B1"), "=A1:B1");
    assert_eq!(normalize_formula_for_compare("=A1 ! B1"), "=A1!B1");
}

#[test]
fn normalize_formula_preserves_string_literals_and_quoted_identifiers() {
    assert_eq!(
        normalize_formula_for_compare("= \"a  b\" & \"C\""),
        "=\"a  b\"&\"C\""
    );
    assert_eq!(
        normalize_formula_for_compare("='My Sheet' ! a1"),
        "='My Sheet'!A1"
    );
}

#[test]
fn normalize_formula_is_case_insensitive_outside_literals() {
    assert_eq!(normalize_formula_for_compare("=sum(a1)"), "=SUM(A1)");
    assert_eq!(normalize_formula_for_compare("=Sheet1!a1"), "=SHEET1!A1");
}

#[test]
fn normalize_formula_preserves_case_inside_literals() {
    assert_eq!(normalize_formula_for_compare("=\"aBc\""), "=\"aBc\"");
    assert_eq!(
        normalize_formula_for_compare("='MiXeD Sheet'!A1"),
        "='MiXeD Sheet'!A1"
    );
}
