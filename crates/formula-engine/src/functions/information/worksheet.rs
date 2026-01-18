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
    let key = key.trim();
    if key.eq_ignore_ascii_case("recalc") {
        return Some(InfoType::Recalc);
    }
    if key.eq_ignore_ascii_case("system") {
        return Some(InfoType::System);
    }
    if key.eq_ignore_ascii_case("directory") {
        return Some(InfoType::Directory);
    }
    if key.eq_ignore_ascii_case("numfile") {
        return Some(InfoType::NumFile);
    }
    if key.eq_ignore_ascii_case("origin") {
        return Some(InfoType::Origin);
    }
    if key.eq_ignore_ascii_case("osversion") {
        return Some(InfoType::OSVersion);
    }
    if key.eq_ignore_ascii_case("release") {
        return Some(InfoType::Release);
    }
    if key.eq_ignore_ascii_case("version") {
        return Some(InfoType::Version);
    }
    if key.eq_ignore_ascii_case("memavail") {
        return Some(InfoType::MemAvail);
    }
    if key.eq_ignore_ascii_case("totmem") {
        return Some(InfoType::TotMem);
    }
    None
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

    let out_len = dir.len() + 1;
    let mut out = String::new();
    if out.try_reserve_exact(out_len).is_err() {
        debug_assert!(false, "allocation failed (workbook_dir_for_excel, len={out_len})");
        return String::new();
    }
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
    let key = key.trim();
    if key.eq_ignore_ascii_case("address") {
        return Some(CellInfoType::Address);
    }
    if key.eq_ignore_ascii_case("col") {
        return Some(CellInfoType::Col);
    }
    if key.eq_ignore_ascii_case("row") {
        return Some(CellInfoType::Row);
    }
    if key.eq_ignore_ascii_case("contents") {
        return Some(CellInfoType::Contents);
    }
    if key.eq_ignore_ascii_case("type") {
        return Some(CellInfoType::Type);
    }
    if key.eq_ignore_ascii_case("format") {
        return Some(CellInfoType::Format);
    }
    if key.eq_ignore_ascii_case("color") {
        return Some(CellInfoType::Color);
    }
    if key.eq_ignore_ascii_case("parentheses") {
        return Some(CellInfoType::Parentheses);
    }
    if key.eq_ignore_ascii_case("width") {
        return Some(CellInfoType::Width);
    }
    if key.eq_ignore_ascii_case("protect") {
        return Some(CellInfoType::Protect);
    }
    if key.eq_ignore_ascii_case("prefix") {
        return Some(CellInfoType::Prefix);
    }
    // Excel returns an empty string for `CELL("filename")` until the workbook is saved.
    if key.eq_ignore_ascii_case("filename") {
        return Some(CellInfoType::Filename);
    }
    // Notes:
    // - `CELL("width")` consults per-column metadata when available (defaulting to 8.43) and
    //   encodes whether the width is explicitly set using Excel's `+0.1` convention.
    // - `CELL("prefix")` consults the cell's effective horizontal alignment.
    // - `CELL("protect")` consults the cell's effective protection formatting.
    // - `CELL("color")`/`CELL("parentheses")`/`CELL("format")` are implemented based on the
    //   cell number format string, but do not consider conditional formatting rules.
    None
}

fn quote_sheet_name(name: &str) -> String {
    if name.is_empty() {
        return String::new();
    }
    let out_len = name.len().saturating_add(2);
    let mut out = String::new();
    if out.try_reserve_exact(out_len).is_err() {
        debug_assert!(false, "allocation failed (quote_sheet_name, len={out_len})");
        return String::new();
    }
    formula_model::push_sheet_name_a1(&mut out, name);
    out
}

fn abs_a1(addr: CellAddr) -> String {
    let mut out = String::new();
    formula_model::push_a1_cell_ref(addr.row, addr.col, true, true, &mut out);
    out
}

fn cell_number_format<'a>(
    ctx: &'a dyn FunctionContext,
    sheet_id: &SheetId,
    addr: CellAddr,
) -> Option<&'a str> {
    ctx.get_cell_number_format(sheet_id, addr)
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
            // Mirror `get_cell_value` bounds behavior: out-of-bounds references should surface
            // `#REF!` rather than fabricating an absolute A1 string.
            let (rows, cols) = ctx.sheet_dimensions(&reference.sheet_id);
            if addr.row >= rows || addr.col >= cols {
                return Value::Error(ErrorKind::Ref);
            }

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
        CellInfoType::Col => {
            let (rows, cols) = ctx.sheet_dimensions(&reference.sheet_id);
            if addr.row >= rows || addr.col >= cols {
                return Value::Error(ErrorKind::Ref);
            }
            Value::Number((u64::from(addr.col) + 1) as f64)
        }
        CellInfoType::Row => {
            let (rows, cols) = ctx.sheet_dimensions(&reference.sheet_id);
            if addr.row >= rows || addr.col >= cols {
                return Value::Error(ErrorKind::Ref);
            }
            Value::Number((u64::from(addr.row) + 1) as f64)
        }
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

            // Mirror `get_cell_value` bounds behavior: out-of-bounds references should surface
            // `#REF!`.
            let (rows, cols) = ctx.sheet_dimensions(&cell_ref.sheet_id);
            if addr.row >= rows || addr.col >= cols {
                return Value::Error(ErrorKind::Ref);
            }

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
            let (rows, cols) = ctx.sheet_dimensions(&cell_ref.sheet_id);
            if addr.row >= rows || addr.col >= cols {
                return Value::Error(ErrorKind::Ref);
            }
            let fmt = cell_number_format(ctx, &cell_ref.sheet_id, addr);
            Value::Text(cell_format_code(fmt))
        }
        CellInfoType::Color => {
            let cell_ref = record_explicit_cell(ctx);
            let (rows, cols) = ctx.sheet_dimensions(&cell_ref.sheet_id);
            if addr.row >= rows || addr.col >= cols {
                return Value::Error(ErrorKind::Ref);
            }
            let format_code = cell_number_format(ctx, &cell_ref.sheet_id, addr);
            let options = format_options_for_cell(ctx);
            let info = formula_format::cell_format_info(format_code, &options);
            Value::Number(info.color as f64)
        }
        CellInfoType::Parentheses => {
            let cell_ref = record_explicit_cell(ctx);
            let (rows, cols) = ctx.sheet_dimensions(&cell_ref.sheet_id);
            if addr.row >= rows || addr.col >= cols {
                return Value::Error(ErrorKind::Ref);
            }
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
            let sheet_id = &reference.sheet_id;
            let (rows, cols) = ctx.sheet_dimensions(sheet_id);
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

            let props = ctx.col_properties(sheet_id, addr.col);
            if props.as_ref().is_some_and(|p| p.hidden) {
                return Value::Number(0.0);
            }

            let (width, is_custom) = match props.and_then(|p| p.width) {
                Some(w) => (w as f64, true),
                None => match ctx.sheet_default_col_width(sheet_id) {
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
            let style = ctx.effective_cell_style(&cell_ref.sheet_id, addr);
            Value::Number(if style.locked { 1.0 } else { 0.0 })
        }
        CellInfoType::Prefix => {
            // `CELL("prefix")` consults alignment/prefix metadata but should avoid recording an
            // implicit self-reference when `reference` is omitted (to prevent dynamic-deps cycles).
            let cell_ref = record_explicit_cell(ctx);
            // Mirror `get_cell_value` bounds behavior: out-of-bounds references should surface
            // `#REF!` rather than defaulting to an empty prefix.
            let (rows, cols) = ctx.sheet_dimensions(&cell_ref.sheet_id);
            if addr.row >= rows || addr.col >= cols {
                return Value::Error(ErrorKind::Ref);
            }

            let style = ctx.effective_cell_style(&cell_ref.sheet_id, addr);
            let prefix = match style.alignment_horizontal {
                Some(HorizontalAlignment::Left) => "'",
                Some(HorizontalAlignment::Center) => "^",
                Some(HorizontalAlignment::Right) => "\"",
                Some(HorizontalAlignment::Fill) => "\\",
                // `General`, `Justify`, or unset alignment.
                _ => "",
            };

            Value::Text(prefix.to_string())
        }
        CellInfoType::Filename => {
            // `CELL("filename")` depends on workbook metadata, but keep the dynamic dependency trace
            // behavior consistent with other CELL variants: record the reference argument only when
            // it is explicitly provided (to avoid implicit self-edges when reference is omitted).
            let cell_ref = record_explicit_cell(ctx);

            // Mirror `get_cell_value` bounds behavior: out-of-bounds references should surface
            // `#REF!` rather than returning the workbook filename for an invalid address.
            let (rows, cols) = ctx.sheet_dimensions(&cell_ref.sheet_id);
            if addr.row >= rows || addr.col >= cols {
                return Value::Error(ErrorKind::Ref);
            }
            // Use the sheet containing the referenced cell (not necessarily the current sheet).
            let sheet_name = match &cell_ref.sheet_id {
                SheetId::Local(id) => ctx.sheet_name(*id).unwrap_or_default(),
                // We don't have separate workbook file metadata for external workbooks. The
                // canonical external sheet key already includes `[Book.xlsx]Sheet`, so return it
                // directly (no directory prefix).
                SheetId::External(key) => return Value::Text(key.clone()),
            };

            // Excel returns "" until the workbook has a known filename (i.e. it has been saved).
            let Some(filename) = ctx.workbook_filename().filter(|s| !s.is_empty()) else {
                return Value::Text(String::new());
            };

            let key = crate::external_refs::format_external_key(filename, &sheet_name);
            match ctx.workbook_directory().filter(|s| !s.is_empty()) {
                Some(dir) => Value::Text(format!("{}{}", workbook_dir_for_excel(dir), key)),
                None => Value::Text(key),
            }
        }
    }
}
