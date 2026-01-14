use crate::calc_settings::CalculationMode;
use crate::eval::{parse_a1, CellAddr};
use crate::functions::{FunctionContext, Reference, SheetId};
use crate::{ErrorKind, Value};
use formula_format::cell_format_code;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum InfoType {
    Recalc,
    System,
    Directory,
    NumFile,
    Origin,
    OSVersion,
    Release,
    Version,
    MemAvail,
    TotMem,
}

fn parse_info_type(key: &str) -> Option<InfoType> {
    match key.trim().to_ascii_lowercase().as_str() {
        "recalc" => Some(InfoType::Recalc),
        "system" => Some(InfoType::System),
        "directory" => Some(InfoType::Directory),
        "numfile" => Some(InfoType::NumFile),
        "origin" => Some(InfoType::Origin),
        "osversion" => Some(InfoType::OSVersion),
        "release" => Some(InfoType::Release),
        "version" => Some(InfoType::Version),
        "memavail" => Some(InfoType::MemAvail),
        "totmem" => Some(InfoType::TotMem),
        _ => None,
    }
}

/// Excel INFO(type_text) worksheet information function.
pub fn info(ctx: &dyn FunctionContext, type_text: &str) -> Value {
    let Some(info_type) = parse_info_type(type_text) else {
        // Unrecognized type_text.
        return Value::Error(ErrorKind::Value);
    };

    match info_type {
        // Deterministic & commonly used values.
        InfoType::Recalc => match ctx.calculation_mode() {
            CalculationMode::Automatic => Value::Text("Automatic".to_string()),
            CalculationMode::AutomaticNoTable => {
                Value::Text("Automatic except for tables".to_string())
            }
            CalculationMode::Manual => Value::Text("Manual".to_string()),
        },
        InfoType::System => Value::Text(ctx.info_system().unwrap_or("pcdos").to_string()),
        InfoType::NumFile => Value::Number(ctx.sheet_count() as f64),
        InfoType::Directory => ctx
            .info_directory()
            .map(|s| Value::Text(s.to_string()))
            .unwrap_or(Value::Error(ErrorKind::NA)),
        InfoType::Origin => ctx
            .info_origin()
            .map(|s| Value::Text(s.to_string()))
            .unwrap_or(Value::Error(ErrorKind::NA)),
        InfoType::OSVersion => ctx
            .info_osversion()
            .map(|s| Value::Text(s.to_string()))
            .unwrap_or(Value::Error(ErrorKind::NA)),
        InfoType::Release => ctx
            .info_release()
            .map(|s| Value::Text(s.to_string()))
            .unwrap_or(Value::Error(ErrorKind::NA)),
        InfoType::Version => ctx
            .info_version()
            .map(|s| Value::Text(s.to_string()))
            .unwrap_or(Value::Error(ErrorKind::NA)),
        InfoType::MemAvail => ctx
            .info_memavail()
            .filter(|n| n.is_finite())
            .map(Value::Number)
            .unwrap_or(Value::Error(ErrorKind::NA)),
        InfoType::TotMem => ctx
            .info_totmem()
            .filter(|n| n.is_finite())
            .map(Value::Number)
            .unwrap_or(Value::Error(ErrorKind::NA)),
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum CellInfoType {
    Address,
    Col,
    Row,
    Contents,
    Type,
    Format,
    Color,
    Parentheses,
    Width,
    Protect,
    Prefix,
    Filename,
}

fn parse_cell_info_type(key: &str) -> Option<CellInfoType> {
    match key.trim().to_ascii_lowercase().as_str() {
        "address" => Some(CellInfoType::Address),
        "col" => Some(CellInfoType::Col),
        "row" => Some(CellInfoType::Row),
        "contents" => Some(CellInfoType::Contents),
        "type" => Some(CellInfoType::Type),
        "format" => Some(CellInfoType::Format),
        "color" => Some(CellInfoType::Color),
        "parentheses" => Some(CellInfoType::Parentheses),
        "width" => Some(CellInfoType::Width),
        "protect" => Some(CellInfoType::Protect),
        "prefix" => Some(CellInfoType::Prefix),
        // Excel returns an empty string for `CELL("filename")` until the workbook is saved.
        // This engine does not currently track workbook file metadata, so always return "".
        "filename" => Some(CellInfoType::Filename),
        // Notes:
        // - `CELL("width")`/`CELL("protect")`/`CELL("prefix")` return best-effort defaults because
        //   this engine does not currently track those properties.
        // - `CELL("color")`/`CELL("parentheses")`/`CELL("format")` are implemented based on the
        //   cell number format string, but do not consider conditional formatting rules.
        _ => None,
    }
}

fn effective_style_id(ctx: &dyn FunctionContext, sheet_id: &SheetId, addr: CellAddr) -> u32 {
    let style_id = ctx.cell_style_id(sheet_id, addr);
    if style_id != 0 {
        return style_id;
    }

    // If no cell-level style is present, fall back to row then column defaults (matching Excel
    // style precedence).
    if let Some(style_id) = ctx.row_style_id(sheet_id, addr.row) {
        return style_id;
    }
    if let Some(props) = ctx.col_properties(sheet_id, addr.col) {
        if let Some(style_id) = props.style_id {
            return style_id;
        }
    }

    0
}

fn is_ident_cont_char(c: char) -> bool {
    matches!(c, '$' | '_' | '\\' | '.' | 'A'..='Z' | 'a'..='z' | '0'..='9')
}

fn quote_sheet_name(name: &str) -> String {
    if name.is_empty() {
        return String::new();
    }

    let starts_like_number = matches!(name.chars().next(), Some('0'..='9' | '.'));
    let starts_like_r1c1 = matches!(name.chars().next(), Some('R' | 'r' | 'C' | 'c'))
        && matches!(name.chars().nth(1), Some('0'..='9' | '['));
    let looks_like_a1 = parse_a1(name).is_ok();
    let needs_quote = starts_like_number
        || starts_like_r1c1
        || looks_like_a1
        || name.chars().any(|c| !is_ident_cont_char(c));

    if !needs_quote {
        return name.to_string();
    }

    let escaped = name.replace('\'', "''");
    format!("'{escaped}'")
}

fn abs_a1(addr: CellAddr) -> String {
    let a1 = addr.to_a1();
    let split = a1.find(|c: char| c.is_ascii_digit()).unwrap_or(a1.len());
    let (col, row) = a1.split_at(split);
    format!("${col}${row}")
}

fn cell_number_format<'a>(
    ctx: &'a dyn FunctionContext,
    sheet_id: &SheetId,
    addr: CellAddr,
) -> Option<&'a str> {
    let style_id = effective_style_id(ctx, sheet_id, addr);
    ctx.style_table()
        .and_then(|styles| styles.get(style_id))
        .and_then(|style| style.number_format.as_deref())
}

fn format_options_for_cell(ctx: &dyn FunctionContext) -> formula_format::FormatOptions {
    use crate::date::ExcelDateSystem;
    use formula_format::DateSystem;

    formula_format::FormatOptions {
        locale: ctx.value_locale().separators,
        date_system: match ctx.date_system() {
            ExcelDateSystem::Excel1900 { .. } => DateSystem::Excel1900,
            ExcelDateSystem::Excel1904 => DateSystem::Excel1904,
        },
    }
}

/// Excel CELL(info_type, [reference]) worksheet information function.
pub fn cell(ctx: &dyn FunctionContext, info_type: &str, reference: Option<Reference>) -> Value {
    let info_type = info_type.trim();
    if info_type.is_empty() {
        return Value::Error(ErrorKind::Value);
    }

    let Some(info_type) = parse_cell_info_type(info_type) else {
        return Value::Error(ErrorKind::Value);
    };

    // Track whether the caller explicitly provided a reference argument.
    //
    // When `reference` is omitted, Excel uses the "current cell" as the implicit reference.
    // That implicit self-reference should not be recorded as a dynamic dependency; otherwise,
    // formulas that contain INDIRECT/OFFSET elsewhere (thus enabling dynamic tracing) can
    // accidentally record a self-edge and become a circular reference.
    let reference_provided = reference.is_some();
    let reference = reference.unwrap_or_else(|| Reference {
        sheet_id: SheetId::Local(ctx.current_sheet_id()),
        start: ctx.current_cell_addr(),
        end: ctx.current_cell_addr(),
    });
    let reference = reference.normalized();
    let addr = reference.start;

    let record_explicit_cell = |ctx: &dyn FunctionContext| -> Reference {
        let cell_ref = Reference {
            sheet_id: reference.sheet_id.clone(),
            start: addr,
            end: addr,
        };
        if reference_provided {
            ctx.record_reference(&cell_ref);
        }
        cell_ref
    };

    match info_type {
        CellInfoType::Address => {
            let abs = abs_a1(addr);
            let include_sheet = match &reference.sheet_id {
                SheetId::Local(id) => *id != ctx.current_sheet_id(),
                SheetId::External(_) => true,
            };
            if !include_sheet {
                return Value::Text(abs);
            }

            let sheet_name = match &reference.sheet_id {
                SheetId::Local(id) => ctx.sheet_name(*id).map(|s| s.to_string()),
                SheetId::External(key) => Some(key.clone()),
            };

            match sheet_name {
                Some(name) if !name.is_empty() => {
                    Value::Text(format!("{}!{abs}", quote_sheet_name(&name)))
                }
                _ => Value::Text(abs),
            }
        }
        CellInfoType::Col => Value::Number((u64::from(addr.col) + 1) as f64),
        CellInfoType::Row => Value::Number((u64::from(addr.row) + 1) as f64),
        CellInfoType::Contents => {
            let cell_ref = record_explicit_cell(ctx);

            if let Some(formula) = ctx.get_cell_formula(&cell_ref.sheet_id, addr) {
                let mut out = formula.to_string();
                if !out.trim_start().starts_with('=') {
                    out.insert(0, '=');
                }
                return Value::Text(out);
            }

            match ctx.get_cell_value(&cell_ref.sheet_id, addr) {
                // Excel treats a blank cell as 0 when returning its contents.
                Value::Blank => Value::Number(0.0),
                other => other,
            }
        }
        CellInfoType::Type => {
            let cell_ref = record_explicit_cell(ctx);

            if ctx.get_cell_formula(&cell_ref.sheet_id, addr).is_some() {
                return Value::Text("v".to_string());
            }

            let code = match ctx.get_cell_value(&cell_ref.sheet_id, addr) {
                Value::Blank => "b",
                Value::Text(_) => "l",
                _ => "v",
            };
            Value::Text(code.to_string())
        }
        CellInfoType::Format => {
            let cell_ref = record_explicit_cell(ctx);
            let fmt = cell_number_format(ctx, &cell_ref.sheet_id, addr);
            Value::Text(cell_format_code(fmt))
        }
        CellInfoType::Color => {
            let cell_ref = record_explicit_cell(ctx);
            let format_code = cell_number_format(ctx, &cell_ref.sheet_id, addr);
            let options = format_options_for_cell(ctx);
            let info = formula_format::cell_format_info(format_code, &options);
            Value::Number(info.color as f64)
        }
        CellInfoType::Parentheses => {
            let cell_ref = record_explicit_cell(ctx);
            let format_code = cell_number_format(ctx, &cell_ref.sheet_id, addr);
            let options = format_options_for_cell(ctx);
            let info = formula_format::cell_format_info(format_code, &options);
            Value::Number(info.parentheses as f64)
        }
        CellInfoType::Width => {
            // `CELL("width")` consults column metadata but should avoid recording an implicit
            // self-reference when `reference` is omitted (to prevent dynamic-deps cycles).
            let _cell_ref = record_explicit_cell(ctx);

            // This engine does not currently track per-column widths. Excel's default column width
            // is 8.43 "character" units.
            Value::Number(8.43)
        }
        CellInfoType::Protect => {
            // `CELL("protect")` consults cell protection metadata but should avoid recording an
            // implicit self-reference when `reference` is omitted (to prevent dynamic-deps cycles).
            let _cell_ref = record_explicit_cell(ctx);

            // Excel's default cell style is locked, so return 1 ("locked").
            Value::Number(1.0)
        }
        CellInfoType::Prefix => {
            // `CELL("prefix")` consults alignment/prefix metadata but should avoid recording an
            // implicit self-reference when `reference` is omitted (to prevent dynamic-deps cycles).
            let _cell_ref = record_explicit_cell(ctx);

             // This engine does not currently track per-cell alignment/prefix formatting, so return
             // the empty string.
             Value::Text(String::new())
         }
        CellInfoType::Filename => Value::Text(String::new()),
    }
}
