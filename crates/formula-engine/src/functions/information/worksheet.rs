use crate::calc_settings::CalculationMode;
use crate::date::ExcelDateSystem;
use crate::eval::{parse_a1, CellAddr};
use crate::functions::{FunctionContext, Reference, SheetId};
use crate::{ErrorKind, Value};
use formula_format::cell_format_code;
use formula_model::HorizontalAlignment;

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

fn workbook_dir_for_excel(dir: &str) -> String {
    if dir.is_empty() {
        return String::new();
    }
    if dir.ends_with('/') || dir.ends_with('\\') {
        return dir.to_string();
    }

    // Excel returns directory strings with a trailing path separator. We don't want to probe the
    // OS, so infer the separator from the host-supplied directory string.
    let last_slash = dir.rfind('/');
    let last_backslash = dir.rfind('\\');
    let sep = match (last_slash, last_backslash) {
        (Some(i), Some(j)) => {
            if i > j {
                '/'
            } else {
                '\\'
            }
        }
        (Some(_), None) => '/',
        (None, Some(_)) => '\\',
        (None, None) => '/',
    };

    let mut out = String::with_capacity(dir.len() + 1);
    out.push_str(dir);
    out.push(sep);
    out
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
        InfoType::Directory => {
            // Prefer the host-provided `INFO("directory")` override if available.
            if let Some(dir) = ctx.info_directory().filter(|d| !d.is_empty()) {
                Value::Text(workbook_dir_for_excel(dir))
            } else {
                // Otherwise fall back to workbook file metadata. Excel returns `#N/A` until the
                // workbook has been saved, which we model as having a `workbook_filename`.
                match (ctx.workbook_filename(), ctx.workbook_directory()) {
                    (Some(_), Some(dir)) if !dir.is_empty() => {
                        Value::Text(workbook_dir_for_excel(dir))
                    }
                    _ => Value::Error(ErrorKind::NA),
                }
            }
        }
        InfoType::Origin => {
            // `INFO("origin")` is UI/view-state driven in Excel (the top-left visible cell).
            //
            // Prefer structured view metadata when provided by the host, but fall back to the
            // legacy string-based `EngineInfo.origin`/`origin_by_sheet` plumbing for compatibility.
            if let Some(origin) = ctx.sheet_origin_cell(ctx.current_sheet_id()) {
                return Value::Text(abs_a1(origin));
            }
            if let Some(origin) = ctx.info_origin().map(str::trim).filter(|s| !s.is_empty()) {
                // If the legacy value looks like an A1 reference, normalize it to Excel's absolute
                // A1 form (`$A$1`). Otherwise return it verbatim for backward compatibility.
                if let Ok(addr) = parse_a1(origin) {
                    return Value::Text(abs_a1(addr));
                }
                return Value::Text(origin.to_string());
            }

            // Excel defaults to the top-left visible cell when no origin is provided.
            Value::Text(abs_a1(CellAddr { row: 0, col: 0 }))
        }
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
        "filename" => Some(CellInfoType::Filename),
        // Notes:
        // - `CELL("width")` consults per-column metadata when available (defaulting to 8.43) and
        //   encodes whether the width is explicitly set using Excel's `+0.1` convention.
        // - `CELL("prefix")` consults the cell's effective horizontal alignment.
        // - `CELL("protect")` consults the cell's effective protection formatting.
        // - `CELL("color")`/`CELL("parentheses")`/`CELL("format")` are implemented based on the
        //   cell number format string, but do not consider conditional formatting rules.
        _ => None,
    }
}

fn style_layer_ids(ctx: &dyn FunctionContext, sheet_id: &SheetId, addr: CellAddr) -> [u32; 5] {
    let col_style_id = ctx
        .col_properties(sheet_id, addr.col)
        .and_then(|props| props.style_id)
        .unwrap_or(0);

    // Style precedence matches the DocumentController layering:
    //   sheet < col < row < range-run < cell
    [
        ctx.cell_style_id(sheet_id, addr),
        ctx.format_run_style_id(sheet_id, addr),
        ctx.row_style_id(sheet_id, addr.row).unwrap_or(0),
        col_style_id,
        ctx.sheet_default_style_id(sheet_id).unwrap_or(0),
    ]
}

fn resolve_locked(ctx: &dyn FunctionContext, sheet_id: &SheetId, addr: CellAddr) -> bool {
    let Some(styles) = ctx.style_table() else {
        // Excel defaults to locked cells.
        return true;
    };

    for style_id in style_layer_ids(ctx, sheet_id, addr) {
        if let Some(style) = styles.get(style_id) {
            if let Some(protection) = &style.protection {
                return protection.locked;
            }
        }
    }

    // Excel defaults to locked cells.
    true
}

fn resolve_horizontal_alignment(
    ctx: &dyn FunctionContext,
    sheet_id: &SheetId,
    addr: CellAddr,
) -> HorizontalAlignment {
    ctx.cell_horizontal_alignment(sheet_id, addr)
        .unwrap_or(HorizontalAlignment::General)
}
fn resolve_number_format<'a>(
    ctx: &'a dyn FunctionContext,
    sheet_id: &SheetId,
    addr: CellAddr,
) -> Option<&'a str> {
    if let Some(fmt) = ctx.get_cell_number_format(sheet_id, addr) {
        return Some(fmt);
    }

    let styles = ctx.style_table()?;

    // Style precedence matches the DocumentController layering:
    //   sheet < col < row < range-run < cell
    //
    // When a style does not specify a number format (`number_format=None`), treat it as "inherit"
    // so lower-precedence layers can contribute the number format.
    for style_id in style_layer_ids(ctx, sheet_id, addr) {
        if let Some(style) = styles.get(style_id) {
            if let Some(fmt) = style.number_format.as_deref() {
                return Some(fmt);
            }
        }
    }

    None
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
    ctx.get_cell_number_format(sheet_id, addr)
        .or_else(|| resolve_number_format(ctx, sheet_id, addr))
}

fn format_options_for_cell(ctx: &dyn FunctionContext) -> formula_format::FormatOptions {
    formula_format::FormatOptions {
        locale: ctx.value_locale().separators,
        date_system: match ctx.date_system() {
            ExcelDateSystem::Excel1900 { .. } => formula_format::DateSystem::Excel1900,
            ExcelDateSystem::Excel1904 => formula_format::DateSystem::Excel1904,
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
        let is_self_reference = matches!(&cell_ref.sheet_id, SheetId::Local(id) if *id == ctx.current_sheet_id())
            && addr == ctx.current_cell_addr();
        if reference_provided && !is_self_reference {
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
            // `CELL("width")` consults per-column metadata only.
            //
            // The reference argument is used for its *address* (column) only, not its value, so we
            // intentionally do not record it as a cell-value dependency. This keeps the dependency
            // graph aligned with Excel (e.g. `CELL("width",A1)` is not a circular dependency in
            // A1) and prevents spurious dynamic dependencies when the reference is produced by
            // `INDIRECT`/`OFFSET`.
            //
            // Mirror `get_cell_value` bounds behavior: out-of-bounds references should surface
            // `#REF!` rather than defaulting to a sheet/Excel width fallback.
            let (rows, cols) = ctx.sheet_dimensions(&reference.sheet_id);
            if addr.row >= rows || addr.col >= cols {
                return Value::Error(ErrorKind::Ref);
            }

            // Excel returns a number where the integer part is the column width (in characters),
            // rounded down, and the first decimal digit is `0` when the column uses the sheet
            // default width or `1` when it uses an explicit per-column override.
            //
            // Hidden columns always return `0`, regardless of the stored width.
            //
            // Column widths are stored in Excel "character" units (OOXML `col/@width`).
            const EXCEL_STANDARD_COL_WIDTH: f64 = 8.43;
            const EXCEL_EXPLICIT_WIDTH_MARKER: f64 = 0.1;

            let props = ctx.col_properties(&reference.sheet_id, addr.col);
            if props.as_ref().is_some_and(|p| p.hidden) {
                return Value::Number(0.0);
            }

            let (width, is_custom) = match props.and_then(|p| p.width) {
                Some(w) => (w as f64, true),
                None => match ctx.sheet_default_col_width(&reference.sheet_id) {
                    Some(w) => (w as f64, false),
                    None => (EXCEL_STANDARD_COL_WIDTH, false),
                },
            };

            let chars = width.floor();
            let flag = if is_custom {
                EXCEL_EXPLICIT_WIDTH_MARKER
            } else {
                0.0
            };
            Value::Number(chars + flag)
        }
        CellInfoType::Protect => {
            // `CELL("protect")` consults cell protection metadata but should avoid recording an
            // implicit self-reference when `reference` is omitted (to prevent dynamic-deps cycles).
            let cell_ref = record_explicit_cell(ctx);

            // Mirror `get_cell_value` bounds behavior: out-of-bounds references should surface
            // `#REF!` rather than defaulting to "locked".
            let (rows, cols) = ctx.sheet_dimensions(&cell_ref.sheet_id);
            if addr.row >= rows || addr.col >= cols {
                return Value::Error(ErrorKind::Ref);
            }

            // Excel's `CELL("protect")` reports the cell's locked formatting state:
            // - All cells default to locked (`1`).
            // - Formatting can explicitly set locked=FALSE (`0`).
            // - The result does *not* depend on whether sheet protection is enabled.
            let locked = resolve_locked(ctx, &cell_ref.sheet_id, addr);
            Value::Number(if locked { 1.0 } else { 0.0 })
        }
        CellInfoType::Prefix => {
            use formula_model::HorizontalAlignment;

            // `CELL("prefix")` consults alignment/prefix metadata but should avoid recording an
            // implicit self-reference when `reference` is omitted (to prevent dynamic-deps cycles).
            let cell_ref = record_explicit_cell(ctx);

            // Mirror `get_cell_value` bounds behavior: out-of-bounds references should surface
            // `#REF!` rather than defaulting to an empty prefix.
            let (rows, cols) = ctx.sheet_dimensions(&cell_ref.sheet_id);
            if addr.row >= rows || addr.col >= cols {
                return Value::Error(ErrorKind::Ref);
            }

            let horizontal = resolve_horizontal_alignment(ctx, &cell_ref.sheet_id, addr);
            let prefix = match horizontal {
                HorizontalAlignment::Left => "'",
                HorizontalAlignment::Center => "^",
                HorizontalAlignment::Right => "\"",
                HorizontalAlignment::Fill => "\\",
                HorizontalAlignment::General | HorizontalAlignment::Justify => "",
            };

            Value::Text(prefix.to_string())
        }
        CellInfoType::Filename => {
            // `CELL("filename")` depends on workbook metadata, but keep the dynamic dependency trace
            // behavior consistent with other CELL variants: record the reference argument only when
            // it is explicitly provided (to avoid implicit self-edges when reference is omitted).
            let cell_ref = record_explicit_cell(ctx);

            // Excel returns "" until the workbook has a known filename (i.e. it has been saved).
            let Some(filename) = ctx.workbook_filename().filter(|s| !s.is_empty()) else {
                return Value::Text(String::new());
            };

            // Use the sheet containing the referenced cell (not necessarily the current sheet).
            let sheet_name = match &cell_ref.sheet_id {
                SheetId::Local(id) => ctx.sheet_name(*id).unwrap_or_default(),
                // We don't have separate workbook file metadata for external workbooks. The
                // canonical external sheet key already includes `[Book.xlsx]Sheet`, so return it
                // directly (no directory prefix).
                SheetId::External(key) => return Value::Text(key.clone()),
            };

            match ctx.workbook_directory().filter(|s| !s.is_empty()) {
                Some(dir) => Value::Text(format!(
                    "{}[{filename}]{sheet_name}",
                    workbook_dir_for_excel(dir)
                )),
                None => Value::Text(format!("[{filename}]{sheet_name}")),
            }
        }
    }
}
