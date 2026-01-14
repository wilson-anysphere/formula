// CLI binaries are not supported on `wasm32-unknown-unknown` (no filesystem / process args).
#[cfg(target_arch = "wasm32")]
fn main() {}

#[cfg(not(target_arch = "wasm32"))]
use std::env;
#[cfg(not(target_arch = "wasm32"))]
use std::io;
#[cfg(not(target_arch = "wasm32"))]
use std::path::PathBuf;

#[cfg(not(target_arch = "wasm32"))]
use formula_model::sheet_name_eq_case_insensitive;
#[cfg(not(target_arch = "wasm32"))]
use formula_xlsb::errors::xlsb_error_display;
#[cfg(not(target_arch = "wasm32"))]
use formula_xlsb::format::{format_a1, format_hex};
#[cfg(not(target_arch = "wasm32"))]
use formula_xlsb::rgce::decode_rgce_with_rgcb;
#[cfg(not(target_arch = "wasm32"))]
use formula_xlsb::{CellValue, Formula, SheetMeta, XlsbWorkbook};

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
    formulas_only: bool,
    max: Option<usize>,
    rgce: bool,
    password: Option<String>,
}

#[cfg(not(target_arch = "wasm32"))]
impl Args {
    fn parse() -> Result<Self, io::Error> {
        let mut path: Option<PathBuf> = None;
        let mut sheet: Option<String> = None;
        let mut formulas_only = false;
        let mut max: Option<usize> = None;
        let mut rgce = false;
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
                "--formulas-only" => formulas_only = true,
                "--max" => {
                    let value = it
                        .next()
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
                "--rgce" => rgce = true,
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
            formulas_only,
            max,
            rgce,
            password,
        })
    }
}

#[cfg(not(target_arch = "wasm32"))]
fn print_usage() {
    println!(
        "\
xlsb_dump: dump XLSB worksheet cells and formula rgce bytes (diagnostics)

Usage:
  xlsb_dump <path> [--password <pw>] [--sheet <name|index>] [--formulas-only] [--max <n>] [--rgce]

Options:
  --password <pw>          Password for Office-encrypted XLSB (OLE EncryptedPackage wrapper)
  --sheet <name|index>   Only dump a single worksheet (match by name or 0-based index)
  --formulas-only        Only print cells with formula payloads
  --max <n>              Limit printed cells per sheet
  --rgce                 Print rgce hex even when formula text decodes successfully
"
    );
}

#[cfg(not(target_arch = "wasm32"))]
fn main() {
    if let Err(err) = run() {
        eprintln!("xlsb_dump: {err}");
        std::process::exit(1);
    }
}

#[cfg(not(target_arch = "wasm32"))]
fn run() -> Result<(), Box<dyn std::error::Error>> {
    let args = Args::parse()?;
    let wb = open_xlsb_workbook(&args.path, args.password.as_deref())?;

    println!("Workbook: {}", args.path.display());
    println!("Sheets:");
    for (idx, meta) in wb.sheet_metas().iter().enumerate() {
        println!("  [{idx}] {}  part={}", meta.name, meta.part_path);
    }

    let selected = resolve_sheets(wb.sheet_metas(), args.sheet.as_deref())?;
    for sheet_index in selected {
        dump_sheet(&wb, sheet_index, &args)?;
    }

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

#[cfg(not(target_arch = "wasm32"))]
fn dump_sheet(wb: &XlsbWorkbook, sheet_index: usize, args: &Args) -> Result<(), formula_xlsb::Error> {
    let meta = &wb.sheet_metas()[sheet_index];
    println!();
    println!("-- Sheet[{sheet_index}] {} ({})", meta.name, meta.part_path);

    let mut printed = 0usize;
    wb.for_each_cell(sheet_index, |cell| {
        if args.formulas_only && cell.formula.is_none() {
            return;
        }

        if let Some(max) = args.max {
            if printed >= max {
                return;
            }
        }

        printed += 1;
        print_cell(cell, args);
    })?;

    Ok(())
}

#[cfg(not(target_arch = "wasm32"))]
fn print_cell(cell: formula_xlsb::Cell, args: &Args) {
    let addr = format_a1(cell.row, cell.col);
    let value = format_cell_value(&cell.value);

    match cell.formula {
        None => println!("{addr}: {value}"),
        Some(Formula { rgce, text, extra, .. }) => {
            match text {
                Some(text) => {
                    if args.rgce {
                        println!("{addr}: {value}  formula={text}  rgce={}", format_hex(&rgce));
                    } else {
                        println!("{addr}: {value}  formula={text}");
                    }
                }
                None => match decode_rgce_with_rgcb(&rgce, &extra) {
                    Ok(decoded) => {
                        // Should be rare (the parser uses the same decoder), but keep output sensible.
                        println!(
                            "{addr}: {value}  formula={decoded}  rgce={}",
                            format_hex(&rgce)
                        );
                    }
                    Err(err) => {
                        let ptg = err
                            .ptg()
                            .map(|b| format!("0x{b:02X}"))
                            .unwrap_or_else(|| "<none>".to_string());
                        println!(
                            "{addr}: {value}  rgce={}  decode_error=({err}; ptg={ptg}, offset={})",
                            format_hex(&rgce),
                            err.offset()
                        );
                    }
                },
            }
        }
    }
}

#[cfg(not(target_arch = "wasm32"))]
fn format_cell_value(value: &CellValue) -> String {
    match value {
        CellValue::Blank => "BLANK".to_string(),
        CellValue::Number(n) => n.to_string(),
        CellValue::Bool(v) => {
            if *v {
                "TRUE".to_string()
            } else {
                "FALSE".to_string()
            }
        }
        CellValue::Error(code) => xlsb_error_display(*code),
        CellValue::Text(s) => format!("{s:?}"),
    }
}
