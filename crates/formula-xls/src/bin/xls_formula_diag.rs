use std::path::Path;

use formula_xls::diagnostics::{collect_xls_formula_diagnostics, SheetFormulaDiagnostics};

fn print_sheet(sheet: &SheetFormulaDiagnostics) {
    let s = &sheet.stats;
    println!("Sheet: {:?} (offset {})", sheet.name, sheet.offset);
    println!("  FORMULA records: {}", s.formula_records);
    println!("    rgce starts with PtgExp: {}", s.formula_ptgexp);
    println!("    rgce starts with PtgTbl: {}", s.formula_ptgtbl);
    println!("  SHRFMLA records: {}", s.shrfmla_records);
    println!("  ARRAY records: {}", s.array_records);
    println!("  TABLE records: {}", s.table_records);
    println!(
        "  unresolved: PtgExp={}, PtgTbl={}",
        s.unresolved_ptgexp, s.unresolved_ptgtbl
    );
    if s.record_parse_errors != 0 || s.payload_parse_errors != 0 {
        println!(
            "  parse errors: record={}, payload={}",
            s.record_parse_errors, s.payload_parse_errors
        );
    }
    if !sheet.errors.is_empty() {
        println!("  errors:");
        for err in &sheet.errors {
            println!("    - {err}");
        }
    }
}

fn main() {
    let mut args = std::env::args().skip(1).collect::<Vec<_>>();
    if args.is_empty() || args.iter().any(|a| a == "-h" || a == "--help") {
        eprintln!("Usage: xls-formula-diag <workbook.xls> [more.xls ...]");
        std::process::exit(if args.is_empty() { 2 } else { 0 });
    }

    let mut exit_code = 0;

    // Support multiple inputs for quick corpus triage.
    for path in std::mem::take(&mut args) {
        println!("== {path} ==");
        match collect_xls_formula_diagnostics(Path::new(&path)) {
            Ok(diag) => {
                if !diag.errors.is_empty() {
                    exit_code = 1;
                    println!("workbook errors:");
                    for err in &diag.errors {
                        println!("  - {err}");
                    }
                }

                for sheet in &diag.sheets {
                    print_sheet(sheet);
                }

                let totals = diag.totals();
                println!("-- totals --");
                println!("  FORMULA records: {}", totals.formula_records);
                println!("    rgce starts with PtgExp: {}", totals.formula_ptgexp);
                println!("    rgce starts with PtgTbl: {}", totals.formula_ptgtbl);
                println!("  SHRFMLA records: {}", totals.shrfmla_records);
                println!("  ARRAY records: {}", totals.array_records);
                println!("  TABLE records: {}", totals.table_records);
                println!(
                    "  unresolved: PtgExp={}, PtgTbl={}",
                    totals.unresolved_ptgexp, totals.unresolved_ptgtbl
                );
                if totals.record_parse_errors != 0 || totals.payload_parse_errors != 0 {
                    exit_code = 1;
                    println!(
                        "  parse errors: record={}, payload={}",
                        totals.record_parse_errors, totals.payload_parse_errors
                    );
                }

                if diag.has_errors() {
                    exit_code = 1;
                }
            }
            Err(err) => {
                exit_code = 1;
                eprintln!("fatal error: {err}");
            }
        }
    }

    if exit_code != 0 {
        std::process::exit(exit_code);
    }
}
