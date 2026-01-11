use std::collections::HashMap;
use std::fs;
use std::io::Write;
use std::path::{Path, PathBuf};

use calamine::{open_workbook_auto, Reader};
use formula_xlsb::XlsbWorkbook;
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
                    a1_notation(row, col)
                );
            };

            let Some(decoded_text) = decoded.as_deref() else {
                panic!(
                    "fixture {}: sheet {sheet_name:?} cell {} has a formula but formula-xlsb could not decode it",
                    path.display(),
                    a1_notation(row, col)
                );
            };

            let cal_norm = normalize_formula_for_compare(&cal_formula);
            let decoded_norm = normalize_formula_for_compare(decoded_text);
            assert_eq!(
                decoded_norm,
                cal_norm,
                "fixture {} sheet {sheet_name} cell {}",
                path.display(),
                a1_notation(row, col)
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

fn read_formula_xlsb_formulas(path: &Path) -> SheetFormulas<Option<String>> {
    let wb = XlsbWorkbook::open(path)
        .unwrap_or_else(|err| panic!("fixture {}: formula-xlsb open failed: {err}", path.display()));

    let mut out: SheetFormulas<Option<String>> = HashMap::new();
    for (sheet_idx, sheet_meta) in wb.sheet_metas().iter().enumerate() {
        let sheet_name = sheet_meta.name.clone();
        let mut formulas: HashMap<CellCoord, Option<String>> = HashMap::new();

        wb.for_each_cell(sheet_idx, |cell| {
            if let Some(formula) = cell.formula {
                formulas.insert((cell.row, cell.col), formula.text);
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
    let mut chars = trimmed.chars().peekable();

    while let Some(ch) = chars.next() {
        if ch == '"' && !in_quoted_ident {
            out.push(ch);
            if in_string {
                if chars.peek() == Some(&'"') {
                    // Escaped quote inside a string literal.
                    out.push('"');
                    chars.next();
                } else {
                    in_string = false;
                }
            } else {
                in_string = true;
            }
            continue;
        }

        if ch == '\'' && !in_string {
            out.push(ch);
            if in_quoted_ident {
                if chars.peek() == Some(&'\'') {
                    // Escaped quote inside a quoted sheet/workbook identifier.
                    out.push('\'');
                    chars.next();
                } else {
                    in_quoted_ident = false;
                }
            } else {
                in_quoted_ident = true;
            }
            continue;
        }

        if !in_string && !in_quoted_ident && ch.is_whitespace() {
            continue;
        }

        if in_string || in_quoted_ident {
            out.push(ch);
        } else {
            out.push(ch.to_ascii_uppercase());
        }
    }

    format!("={out}")
}

fn a1_notation(row: u32, col: u32) -> String {
    format!("{}{}", col_label(col), row + 1)
}

fn col_label(mut col: u32) -> String {
    col += 1; // Excel column labels are 1-based.

    let mut buf = Vec::new();
    while col > 0 {
        let rem = ((col - 1) % 26) as u8;
        buf.push((b'A' + rem) as char);
        col = (col - 1) / 26;
    }
    buf.iter().rev().collect()
}
