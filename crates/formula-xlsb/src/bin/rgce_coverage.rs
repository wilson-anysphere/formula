// CLI binaries are not supported on `wasm32-unknown-unknown` (no filesystem / process args).
#[cfg(target_arch = "wasm32")]
fn main() {}

#[cfg(not(target_arch = "wasm32"))]
use std::collections::BTreeMap;
#[cfg(not(target_arch = "wasm32"))]
use std::env;
#[cfg(not(target_arch = "wasm32"))]
use std::io::{self, Write};
#[cfg(not(target_arch = "wasm32"))]
use std::ops::ControlFlow;
#[cfg(not(target_arch = "wasm32"))]
use std::path::PathBuf;

#[cfg(not(target_arch = "wasm32"))]
use formula_model::sheet_name_eq_case_insensitive;
#[cfg(not(target_arch = "wasm32"))]
use formula_xlsb::format::{format_a1, format_hex};
#[cfg(not(target_arch = "wasm32"))]
use formula_xlsb::rgce::{decode_rgce_with_context_and_rgcb_and_base, CellCoord};
#[cfg(not(target_arch = "wasm32"))]
use formula_xlsb::{Formula, SheetMeta};
#[cfg(not(target_arch = "wasm32"))]
use serde::Serialize;

#[cfg(not(target_arch = "wasm32"))]
#[path = "../xlsb_cli_open.rs"]
mod xlsb_cli_open;
#[cfg(not(target_arch = "wasm32"))]
use xlsb_cli_open::open_xlsb_workbook;

#[cfg(not(target_arch = "wasm32"))]
#[derive(Debug)]
struct Args {
    path: PathBuf,
    sheet: Option<String>,
    max: Option<usize>,
    password: Option<String>,
}

#[cfg(not(target_arch = "wasm32"))]
impl Args {
    fn parse() -> Result<Self, io::Error> {
        let mut path: Option<PathBuf> = None;
        let mut sheet: Option<String> = None;
        let mut max: Option<usize> = None;
        let mut password: Option<String> = None;

        let mut it = env::args().skip(1).peekable();
        while let Some(arg) = it.next() {
            match arg.as_str() {
                "-h" | "--help" => {
                    print_usage();
                    std::process::exit(0);
                }
                "--sheet" => {
                    let value = it.next().ok_or_else(|| {
                        io::Error::new(io::ErrorKind::InvalidInput, "--sheet expects <name|index>")
                    })?;
                    sheet = Some(value);
                }
                "--max" => {
                    let value =
                        it.next()
                            .ok_or_else(|| io::Error::new(io::ErrorKind::InvalidInput, "--max expects <n>"))?;
                    let n: usize = value.parse().map_err(|_| {
                        io::Error::new(io::ErrorKind::InvalidInput, format!("invalid --max value: {value}"))
                    })?;
                    max = Some(n);
                }
                "--password" => {
                    let value = it.next().ok_or_else(|| {
                        io::Error::new(io::ErrorKind::InvalidInput, "--password expects <pw>")
                    })?;
                    password = Some(value);
                }
                _ if arg.starts_with("--sheet=") => {
                    sheet = Some(arg["--sheet=".len()..].to_string());
                }
                _ if arg.starts_with("--max=") => {
                    let value = &arg["--max=".len()..];
                    let n: usize = value.parse().map_err(|_| {
                        io::Error::new(io::ErrorKind::InvalidInput, format!("invalid --max value: {value}"))
                    })?;
                    max = Some(n);
                }
                _ if arg.starts_with("--password=") => {
                    password = Some(arg["--password=".len()..].to_string());
                }
                _ if arg.starts_with('-') => {
                    return Err(io::Error::new(
                        io::ErrorKind::InvalidInput,
                        format!("unknown option: {arg}"),
                    ));
                }
                _ => {
                    if path.is_some() {
                        return Err(io::Error::new(
                            io::ErrorKind::InvalidInput,
                            format!("unexpected argument: {arg}"),
                        ));
                    }
                    path = Some(PathBuf::from(arg));
                }
            }
        }

        let path = path.ok_or_else(|| io::Error::new(io::ErrorKind::InvalidInput, "missing <path>"))?;

        Ok(Self {
            path,
            sheet,
            max,
            password,
        })
    }
}

#[cfg(not(target_arch = "wasm32"))]
fn print_usage() {
    println!(
        "\
rgce_coverage: emit JSONL coverage data for formula rgce decoding

Usage:
  rgce_coverage <path.xlsb> [--password <pw>] [--sheet <name|index>] [--max <n>]

Options:
  --password <pw>          Password for Office-encrypted XLSB (OLE EncryptedPackage wrapper)
  --sheet <name|index>   Only scan a single worksheet (match by name or 0-based index)
  --max <n>              Limit scanned formula cells (across all selected sheets)
"
    );
}

#[cfg(not(target_arch = "wasm32"))]
#[derive(Debug, Serialize)]
struct FormulaCellLine {
    sheet: String,
    a1: String,
    rgce_hex: String,
    rgcb_hex_len: usize,
    decoded: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    ptg: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    offset: Option<usize>,
}

#[cfg(not(target_arch = "wasm32"))]
#[derive(Debug, Serialize)]
struct SummaryLine {
    kind: &'static str,
    formulas_total: usize,
    decoded_ok: usize,
    decoded_failed: usize,
    failures_by_ptg: BTreeMap<String, usize>,
}

#[cfg(not(target_arch = "wasm32"))]
fn main() {
    if let Err(err) = run() {
        eprintln!("rgce_coverage: {err}");
        std::process::exit(1);
    }
}

#[cfg(not(target_arch = "wasm32"))]
fn run() -> Result<(), Box<dyn std::error::Error>> {
    let args = Args::parse()?;
    let wb = open_xlsb_workbook(&args.path, args.password.as_deref())?;

    let selected = resolve_sheets(wb.sheet_metas(), args.sheet.as_deref())?;

    let ctx = wb.workbook_context();
    let mut stdout = io::BufWriter::new(io::stdout());

    let mut formulas_total = 0usize;
    let mut decoded_ok = 0usize;
    let mut decoded_failed = 0usize;
    let mut failures_by_ptg: BTreeMap<String, usize> = BTreeMap::new();

    if args.max == Some(0) {
        let summary = SummaryLine {
            kind: "summary",
            formulas_total,
            decoded_ok,
            decoded_failed,
            failures_by_ptg,
        };
        serde_json::to_writer(&mut stdout, &summary)?;
        writeln!(&mut stdout)?;
        return Ok(());
    }

    let mut stop_all = false;

    for sheet_index in selected {
        if stop_all {
            break;
        }
        let sheet_name = wb.sheet_metas()[sheet_index].name.clone();

        wb.for_each_cell_control_flow(sheet_index, |cell| {
            if stop_all {
                return ControlFlow::Break(());
            }

            let Some(Formula { rgce, extra, text, .. }) = cell.formula else {
                return ControlFlow::Continue(());
            };

            if let Some(max) = args.max {
                if formulas_total >= max {
                    stop_all = true;
                    return ControlFlow::Break(());
                }
            }

            formulas_total += 1;
            let addr = format_a1(cell.row, cell.col);
            let rgce_hex = format_hex(&rgce);
            let base = CellCoord::new(cell.row, cell.col);

            let mut line = FormulaCellLine {
                sheet: sheet_name.clone(),
                a1: addr,
                rgce_hex,
                rgcb_hex_len: extra.len(),
                decoded: text.clone(),
                ptg: None,
                offset: None,
            };

            if line.decoded.is_some() {
                decoded_ok += 1;
            } else {
                decoded_failed += 1;
                // Re-run the decoder so we can surface `ptg` + `offset` for coverage analysis.
                match decode_rgce_with_context_and_rgcb_and_base(&rgce, &extra, ctx, base) {
                    Ok(decoded) => {
                        decoded_ok += 1;
                        decoded_failed = decoded_failed.saturating_sub(1);
                        line.decoded = Some(decoded);
                    }
                    Err(err) => {
                        let ptg_hex = err
                            .ptg()
                            .map(|b| format!("0x{b:02X}"))
                            .unwrap_or_else(|| "null".to_string());
                        *failures_by_ptg.entry(ptg_hex.clone()).or_insert(0) += 1;
                        line.ptg = err.ptg().map(|b| format!("0x{b:02X}"));
                        line.offset = Some(err.offset());
                    }
                }
            }

            serde_json::to_writer(&mut stdout, &line).expect("serialize json");
            writeln!(&mut stdout).expect("write newline");

            if let Some(max) = args.max {
                if formulas_total >= max {
                    stop_all = true;
                    return ControlFlow::Break(());
                }
            }

            ControlFlow::Continue(())
        })?;
    }

    let summary = SummaryLine {
        kind: "summary",
        formulas_total,
        decoded_ok,
        decoded_failed,
        failures_by_ptg,
    };
    serde_json::to_writer(&mut stdout, &summary)?;
    writeln!(&mut stdout)?;

    Ok(())
}

#[cfg(not(target_arch = "wasm32"))]
fn resolve_sheets(sheets: &[SheetMeta], selector: Option<&str>) -> Result<Vec<usize>, io::Error> {
    let Some(selector) = selector else {
        return Ok((0..sheets.len()).collect());
    };

    if let Ok(idx) = selector.parse::<usize>() {
        if idx >= sheets.len() {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                format!("sheet index out of bounds: {idx} (sheets={})", sheets.len()),
            ));
        }
        return Ok(vec![idx]);
    }

    if let Some((idx, _)) = sheets
        .iter()
        .enumerate()
        .find(|(_, s)| sheet_name_eq_case_insensitive(&s.name, selector))
    {
        return Ok(vec![idx]);
    }

    Err(io::Error::new(
        io::ErrorKind::InvalidInput,
        format!("sheet not found: {selector}"),
    ))
}
