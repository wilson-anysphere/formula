use std::borrow::Cow;
use std::collections::{BTreeMap, BTreeSet};
use std::io::{Read, Seek};
use std::path::PathBuf;

use clap::{Parser, Subcommand};
use formula_office_crypto::{decrypt_encrypted_package_ole, is_encrypted_ooxml_ole};
use formula_vba_runtime::{
    parse_program, row_col_to_a1, ExecutionResult, Spreadsheet, VbaError, VbaRuntime, VbaValue,
};
use quick_xml::events::Event;
use quick_xml::Reader as XmlReader;
use serde::{Deserialize, Serialize};
use zip::ZipArchive;

#[derive(Debug, Parser)]
#[command(name = "formula-vba-oracle-cli")]
#[command(about = "Execute VBA macros via formula-vba-runtime and emit deterministic JSON diffs.")]
struct Cli {
    #[command(subcommand)]
    command: Command,
}

#[derive(Debug, Subcommand)]
enum Command {
    /// Execute a macro and emit a deterministic JSON report.
    Run(RunArgs),
    /// Extract sheet names + VBA modules + procedure list from an XLSM / oracle workbook JSON payload.
    Extract(ExtractArgs),
}

#[derive(Debug, Parser)]
struct RunArgs {
    /// Macro/procedure name to execute (e.g. `Main`).
    #[arg(long = "macro")]
    macro_name: String,

    /// Optional input file path. If omitted, reads bytes from stdin.
    #[arg(long)]
    input: Option<PathBuf>,

    /// Input format hint: `auto`, `json`, or `xlsm`.
    #[arg(long, default_value = "auto")]
    format: String,

    /// Password for encrypted (OLE `EncryptedPackage`) XLSM inputs.
    ///
    /// If the input workbook is encrypted, `--password` is required.
    #[arg(long, conflicts_with = "password_file")]
    password: Option<String>,

    /// Read password for encrypted XLSM inputs from a file (first line).
    ///
    /// This is useful when the password contains shell-sensitive characters.
    #[arg(long = "password-file", conflicts_with = "password")]
    password_file: Option<PathBuf>,

    /// Macro arguments as a JSON array (e.g. `[1, \"foo\"]`).
    #[arg(long)]
    args: Option<String>,
}

#[derive(Debug, Parser)]
struct ExtractArgs {
    /// Optional input file path. If omitted, reads bytes from stdin.
    #[arg(long)]
    input: Option<PathBuf>,

    /// Input format hint: `auto`, `json`, or `xlsm`.
    #[arg(long, default_value = "auto")]
    format: String,

    /// Password for encrypted (OLE `EncryptedPackage`) XLSM inputs.
    ///
    /// If the input workbook is encrypted, `--password` is required.
    #[arg(long, conflicts_with = "password_file")]
    password: Option<String>,

    /// Read password for encrypted XLSM inputs from a file (first line).
    ///
    /// This is useful when the password contains shell-sensitive characters.
    #[arg(long = "password-file", conflicts_with = "password")]
    password_file: Option<PathBuf>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct OracleWorkbook {
    #[serde(default)]
    schema_version: u32,
    #[serde(default)]
    active_sheet: Option<String>,
    #[serde(default)]
    sheets: Vec<OracleSheet>,
    #[serde(default)]
    vba_modules: Vec<VbaModule>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct OracleSheet {
    name: String,
    #[serde(default)]
    cells: BTreeMap<String, OracleCell>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct OracleCell {
    #[serde(default)]
    value: Option<serde_json::Value>,
    #[serde(default)]
    formula: Option<String>,
    #[serde(default)]
    format: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct VbaModule {
    name: String,
    code: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct OracleCellSnapshot {
    value: Option<serde_json::Value>,
    formula: Option<String>,
    format: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct OracleCellDiff {
    before: OracleCellSnapshot,
    after: OracleCellSnapshot,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct RunReport {
    ok: bool,
    macro_name: String,
    logs: Vec<String>,
    warnings: Vec<String>,
    error: Option<String>,
    exit_status: i32,
    /// Deterministic cell diffs, grouped by sheet name then A1 address.
    cell_diffs: BTreeMap<String, BTreeMap<String, OracleCellDiff>>,
    workbook_after: OracleWorkbook,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct ExtractReport {
    ok: bool,
    error: Option<String>,
    workbook: OracleWorkbook,
    procedures: Vec<ProcedureSummary>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct ProcedureSummary {
    name: String,
    kind: String,
}

#[derive(Debug, Clone)]
struct CellState {
    value: VbaValue,
    formula: Option<String>,
    format: Option<String>,
}

impl Default for CellState {
    fn default() -> Self {
        Self {
            value: VbaValue::Empty,
            formula: None,
            format: None,
        }
    }
}

#[cfg(test)]
mod leading_slash_zip_entries_tests {
    use super::{extract_from_xlsm, OracleWorkbook};
    use std::io::{Cursor, Read, Write};

    use zip::write::FileOptions;
    use zip::{ZipArchive, ZipWriter};

    fn rewrite_zip_with_leading_slash_entry_names(bytes: &[u8]) -> Vec<u8> {
        let mut input = ZipArchive::new(Cursor::new(bytes)).expect("read input zip");

        let mut output = ZipWriter::new(Cursor::new(Vec::<u8>::new()));
        let base_options = FileOptions::<()>::default();

        for i in 0..input.len() {
            let mut entry = input.by_index(i).expect("open zip entry");
            let name = entry.name().to_string();
            let new_name = if name.starts_with('/') {
                name
            } else {
                format!("/{name}")
            };

            // Do not trust `ZipFile::size()` for allocation; ZIP metadata is untrusted and can
            // advertise enormous uncompressed sizes (zip-bomb style OOM).
            let mut contents = Vec::new();
            entry.read_to_end(&mut contents).expect("read entry bytes");

            let options = base_options.compression_method(entry.compression());

            if entry.is_dir() {
                output
                    .add_directory(new_name, options)
                    .expect("add directory");
            } else {
                output.start_file(new_name, options).expect("start file");
                output.write_all(&contents).expect("write file");
            }
        }

        output.finish().expect("finish zip").into_inner()
    }

    #[test]
    fn extract_from_xlsm_tolerates_leading_slash_zip_entry_names() {
        let fixture_path = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../fixtures/xlsx/macros/basic.xlsm"
        );
        let base = std::fs::read(fixture_path).expect("read basic.xlsm fixture");
        let bytes = rewrite_zip_with_leading_slash_entry_names(&base);

        let (workbook, procedures) = extract_from_xlsm(&bytes).expect("extract");

        assert_eq!(workbook.active_sheet.as_deref(), Some("Sheet1"));
        assert_eq!(workbook.sheets.len(), 1);
        assert_eq!(workbook.sheets[0].name, "Sheet1");
        assert!(
            !workbook.vba_modules.is_empty(),
            "expected at least one VBA module"
        );
        assert!(
            !procedures.is_empty(),
            "expected at least one discovered procedure"
        );

        // Avoid unused warnings when this test is the only reference in this module.
        let _ = OracleWorkbook {
            schema_version: 0,
            active_sheet: None,
            sheets: Vec::new(),
            vba_modules: Vec::new(),
        };
    }
}

#[cfg(test)]
mod encrypted_xlsm_tests {
    use super::{extract_from_xlsm, is_ole_encrypted_ooxml, maybe_decrypt_encrypted_xlsm};
    use std::io::Write;

    #[test]
    fn extract_from_encrypted_xlsm_with_password() {
        let fixture_path = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../fixtures/xlsx/macros/basic.xlsm"
        );
        let base = std::fs::read(fixture_path).expect("read basic.xlsm fixture");

        let password = "Password1234_";

        // Encrypt the XLSM bytes into the OLE `EncryptedPackage` wrapper.
        let mut rng = rand::rng();
        let cursor = std::io::Cursor::new(Vec::<u8>::new());
        let mut writer =
            ms_offcrypto_writer::Ecma376AgileWriter::create(&mut rng, password, cursor)
                .expect("create encryptor");
        writer.write_all(&base).expect("write plaintext");
        let cursor = writer.into_inner().expect("finalize encryptor");
        let encrypted = cursor.into_inner();

        assert!(
            is_ole_encrypted_ooxml(&encrypted),
            "expected ms-offcrypto-writer to emit an OLE EncryptedPackage wrapper"
        );

        let decrypted =
            maybe_decrypt_encrypted_xlsm(&encrypted, Some(password)).expect("decrypt fixture");

        let (workbook, procedures) =
            extract_from_xlsm(decrypted.as_ref()).expect("extract decrypted");

        assert!(
            !workbook.vba_modules.is_empty(),
            "expected at least one VBA module"
        );
        assert!(
            !procedures.is_empty(),
            "expected at least one discovered procedure"
        );
    }

    #[test]
    fn encrypted_xlsm_requires_password() {
        let fixture_path = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../fixtures/xlsx/macros/basic.xlsm"
        );
        let base = std::fs::read(fixture_path).expect("read basic.xlsm fixture");

        let password = "Password1234_";
        let mut rng = rand::rng();
        let cursor = std::io::Cursor::new(Vec::<u8>::new());
        let mut writer =
            ms_offcrypto_writer::Ecma376AgileWriter::create(&mut rng, password, cursor)
                .expect("create encryptor");
        writer.write_all(&base).expect("write plaintext");
        let cursor = writer.into_inner().expect("finalize encryptor");
        let encrypted = cursor.into_inner();

        let err = maybe_decrypt_encrypted_xlsm(&encrypted, None).expect_err("expected error");
        assert!(
            err.to_lowercase().contains("password required"),
            "expected a clear password-required error, got: {err}"
        );
    }
}

#[derive(Debug, Clone)]
struct SheetState {
    name: String,
    cells: BTreeMap<(u32, u32), CellState>, // (row,col) 1-based
}

#[derive(Debug, Clone)]
struct ExecutionWorkbook {
    sheets: Vec<SheetState>,
    active_sheet: usize,
    active_cell: (u32, u32),
    logs: Vec<String>,
}

impl ExecutionWorkbook {
    fn from_oracle_workbook(wb: &OracleWorkbook) -> Result<Self, VbaError> {
        let mut sheets = Vec::new();
        for sheet in &wb.sheets {
            let mut cells = BTreeMap::new();
            for (addr, cell) in &sheet.cells {
                let (row, col) = formula_vba_runtime::a1_to_row_col(addr)?;
                let value = json_value_to_vba(cell.value.as_ref());
                let state = CellState {
                    value,
                    formula: cell.formula.clone(),
                    format: cell.format.clone(),
                };
                if !matches!(state.value, VbaValue::Empty)
                    || state.formula.is_some()
                    || state.format.is_some()
                {
                    cells.insert((row, col), state);
                }
            }
            sheets.push(SheetState {
                name: sheet.name.clone(),
                cells,
            });
        }

        if sheets.is_empty() {
            sheets.push(SheetState {
                name: "Sheet1".to_string(),
                cells: BTreeMap::new(),
            });
        }

        let active_sheet = wb
            .active_sheet
            .as_deref()
            .and_then(|name| {
                sheets
                    .iter()
                    .position(|s| s.name.eq_ignore_ascii_case(name))
            })
            .unwrap_or(0);

        Ok(Self {
            sheets,
            active_sheet,
            active_cell: (1, 1),
            logs: Vec::new(),
        })
    }

    fn to_oracle_workbook(&self, vba_modules: Vec<VbaModule>) -> OracleWorkbook {
        let active_sheet = self.sheets.get(self.active_sheet).map(|s| s.name.clone());
        let mut sheets = Vec::new();
        for sheet in &self.sheets {
            let mut cells: BTreeMap<String, OracleCell> = BTreeMap::new();
            for ((row, col), state) in &sheet.cells {
                let addr = row_col_to_a1(*row, *col).unwrap_or_else(|_| format!("R{row}C{col}"));
                let value = vba_value_to_json(&state.value);
                cells.insert(
                    addr,
                    OracleCell {
                        value,
                        formula: state.formula.clone(),
                        format: state.format.clone(),
                    },
                );
            }
            sheets.push(OracleSheet {
                name: sheet.name.clone(),
                cells,
            });
        }

        OracleWorkbook {
            schema_version: 1,
            active_sheet,
            sheets,
            vba_modules,
        }
    }

    fn snapshot(&self) -> BTreeMap<String, BTreeMap<String, OracleCellSnapshot>> {
        let mut out = BTreeMap::new();
        for sheet in &self.sheets {
            let mut cells = BTreeMap::new();
            for ((row, col), state) in &sheet.cells {
                let addr = row_col_to_a1(*row, *col).unwrap_or_else(|_| format!("R{row}C{col}"));
                cells.insert(
                    addr,
                    OracleCellSnapshot {
                        value: vba_value_to_json(&state.value),
                        formula: state.formula.clone(),
                        format: state.format.clone(),
                    },
                );
            }
            out.insert(sheet.name.clone(), cells);
        }
        out
    }
}

impl Spreadsheet for ExecutionWorkbook {
    fn sheet_count(&self) -> usize {
        self.sheets.len()
    }

    fn sheet_name(&self, sheet: usize) -> Option<&str> {
        self.sheets.get(sheet).map(|s| s.name.as_str())
    }

    fn sheet_index(&self, name: &str) -> Option<usize> {
        self.sheets
            .iter()
            .position(|s| s.name.eq_ignore_ascii_case(name))
    }

    fn active_sheet(&self) -> usize {
        self.active_sheet
    }

    fn set_active_sheet(&mut self, sheet: usize) -> Result<(), VbaError> {
        if sheet >= self.sheets.len() {
            return Err(VbaError::Runtime(format!(
                "Sheet index out of range: {sheet}"
            )));
        }
        self.active_sheet = sheet;
        Ok(())
    }

    fn active_cell(&self) -> (u32, u32) {
        self.active_cell
    }

    fn set_active_cell(&mut self, row: u32, col: u32) -> Result<(), VbaError> {
        if row == 0 || col == 0 {
            return Err(VbaError::Runtime("ActiveCell is 1-based".to_string()));
        }
        self.active_cell = (row, col);
        Ok(())
    }

    fn get_cell_value(&self, sheet: usize, row: u32, col: u32) -> Result<VbaValue, VbaError> {
        let sh = self
            .sheets
            .get(sheet)
            .ok_or_else(|| VbaError::Runtime(format!("Unknown sheet index: {sheet}")))?;
        Ok(sh
            .cells
            .get(&(row, col))
            .map(|c| c.value.clone())
            .unwrap_or(VbaValue::Empty))
    }

    fn set_cell_value(
        &mut self,
        sheet: usize,
        row: u32,
        col: u32,
        value: VbaValue,
    ) -> Result<(), VbaError> {
        let sh = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| VbaError::Runtime(format!("Unknown sheet index: {sheet}")))?;
        let cell = sh.cells.entry((row, col)).or_default();
        cell.value = value;
        cell.formula = None;
        Ok(())
    }

    fn get_cell_formula(
        &self,
        sheet: usize,
        row: u32,
        col: u32,
    ) -> Result<Option<String>, VbaError> {
        let sh = self
            .sheets
            .get(sheet)
            .ok_or_else(|| VbaError::Runtime(format!("Unknown sheet index: {sheet}")))?;
        Ok(sh.cells.get(&(row, col)).and_then(|c| c.formula.clone()))
    }

    fn set_cell_formula(
        &mut self,
        sheet: usize,
        row: u32,
        col: u32,
        formula: String,
    ) -> Result<(), VbaError> {
        let sh = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| VbaError::Runtime(format!("Unknown sheet index: {sheet}")))?;
        let cell = sh.cells.entry((row, col)).or_default();
        cell.formula = Some(formula);
        Ok(())
    }

    fn clear_cell_contents(&mut self, sheet: usize, row: u32, col: u32) -> Result<(), VbaError> {
        let sh = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| VbaError::Runtime(format!("Unknown sheet index: {sheet}")))?;
        if let Some(cell) = sh.cells.get_mut(&(row, col)) {
            cell.value = VbaValue::Empty;
            cell.formula = None;
            if cell.format.is_none() {
                sh.cells.remove(&(row, col));
            }
        }
        Ok(())
    }

    fn log(&mut self, message: String) {
        self.logs.push(message);
    }

    fn last_used_row_in_column(&self, sheet: usize, col: u32, start_row: u32) -> Option<u32> {
        let sh = self.sheets.get(sheet)?;
        sh.cells
            .iter()
            .filter_map(|(&(row, c), cell)| {
                if c != col || row > start_row {
                    return None;
                }
                if !matches!(cell.value, VbaValue::Empty) || cell.formula.is_some() {
                    Some(row)
                } else {
                    None
                }
            })
            .max()
    }

    fn next_used_row_in_column(&self, sheet: usize, col: u32, start_row: u32) -> Option<u32> {
        let sh = self.sheets.get(sheet)?;
        sh.cells
            .iter()
            .filter_map(|(&(row, c), cell)| {
                if c != col || row < start_row {
                    return None;
                }
                if !matches!(cell.value, VbaValue::Empty) || cell.formula.is_some() {
                    Some(row)
                } else {
                    None
                }
            })
            .min()
    }

    fn last_used_col_in_row(&self, sheet: usize, row: u32, start_col: u32) -> Option<u32> {
        let sh = self.sheets.get(sheet)?;
        sh.cells
            .iter()
            .filter_map(|(&(r, col), cell)| {
                if r != row || col > start_col {
                    return None;
                }
                if !matches!(cell.value, VbaValue::Empty) || cell.formula.is_some() {
                    Some(col)
                } else {
                    None
                }
            })
            .max()
    }

    fn next_used_col_in_row(&self, sheet: usize, row: u32, start_col: u32) -> Option<u32> {
        let sh = self.sheets.get(sheet)?;
        sh.cells
            .iter()
            .filter_map(|(&(r, col), cell)| {
                if r != row || col < start_col {
                    return None;
                }
                if !matches!(cell.value, VbaValue::Empty) || cell.formula.is_some() {
                    Some(col)
                } else {
                    None
                }
            })
            .min()
    }

    fn used_cells_in_range(
        &self,
        range: formula_vba_runtime::VbaRangeRef,
    ) -> Option<Vec<(u32, u32)>> {
        let sh = self.sheets.get(range.sheet)?;
        let mut out = Vec::new();
        for (&(row, col), cell) in &sh.cells {
            if row < range.start_row
                || row > range.end_row
                || col < range.start_col
                || col > range.end_col
            {
                continue;
            }
            if matches!(cell.value, VbaValue::Empty) && cell.formula.is_none() {
                continue;
            }
            out.push((row, col));
        }
        Some(out)
    }
}

fn json_value_to_vba(value: Option<&serde_json::Value>) -> VbaValue {
    match value {
        None | Some(serde_json::Value::Null) => VbaValue::Empty,
        Some(serde_json::Value::Bool(b)) => VbaValue::Boolean(*b),
        Some(serde_json::Value::Number(n)) => VbaValue::Double(n.as_f64().unwrap_or(0.0)),
        Some(serde_json::Value::String(s)) => VbaValue::String(s.clone()),
        Some(other) => VbaValue::String(other.to_string()),
    }
}

fn vba_value_to_json(value: &VbaValue) -> Option<serde_json::Value> {
    match value {
        VbaValue::Empty | VbaValue::Null => None,
        VbaValue::Boolean(b) => Some(serde_json::Value::Bool(*b)),
        VbaValue::Double(n) => Some(serde_json::Value::Number(
            serde_json::Number::from_f64(*n).unwrap_or_else(|| serde_json::Number::from(0)),
        )),
        VbaValue::String(s) => Some(serde_json::Value::String(s.clone())),
        other => Some(serde_json::Value::String(other.to_string_lossy())),
    }
}

fn read_all_input(input: &Option<PathBuf>) -> Result<Vec<u8>, String> {
    if let Some(path) = input {
        std::fs::read(path).map_err(|e| format!("Failed to read {}: {e}", path.display()))
    } else {
        let mut buf = Vec::new();
        std::io::stdin()
            .read_to_end(&mut buf)
            .map_err(|e| format!("Failed to read stdin: {e}"))?;
        Ok(buf)
    }
}

fn is_ole_encrypted_ooxml(bytes: &[u8]) -> bool {
    is_encrypted_ooxml_ole(bytes)
}

fn maybe_decrypt_encrypted_xlsm<'a>(
    bytes: &'a [u8],
    password: Option<&str>,
) -> Result<Cow<'a, [u8]>, String> {
    if !is_ole_encrypted_ooxml(bytes) {
        return Ok(Cow::Borrowed(bytes));
    }

    let password = password.ok_or_else(|| {
        "password required for encrypted workbook (use --password or --password-file)".to_string()
    })?;
    decrypt_encrypted_package_ole(bytes, password)
        .map(Cow::Owned)
        .map_err(|e| format!("Failed to decrypt workbook: {e}"))
}

fn resolve_password(
    password: &Option<String>,
    password_file: &Option<PathBuf>,
) -> Result<Option<String>, String> {
    if let Some(value) = password.clone() {
        return Ok(Some(value));
    }
    let Some(path) = password_file else {
        return Ok(None);
    };
    let contents = std::fs::read_to_string(path)
        .map_err(|e| format!("Failed to read password file {}: {e}", path.display()))?;
    let pw = contents.lines().next().unwrap_or("").trim().to_string();
    if pw.is_empty() {
        return Err(format!(
            "Password file {} is empty (expected password on first line)",
            path.display()
        ));
    }
    Ok(Some(pw))
}

fn detect_format(bytes: &[u8], hint: &str) -> Result<InputFormat, String> {
    match hint {
        "auto" => {
            if bytes.starts_with(b"PK") || is_ole_encrypted_ooxml(bytes) {
                Ok(InputFormat::Xlsm)
            } else {
                Ok(InputFormat::Json)
            }
        }
        "json" => Ok(InputFormat::Json),
        "xlsm" => Ok(InputFormat::Xlsm),
        other => Err(format!("Unknown format: {other} (expected auto|json|xlsm)")),
    }
}

enum InputFormat {
    Json,
    Xlsm,
}

fn parse_oracle_workbook_json(bytes: &[u8]) -> Result<OracleWorkbook, String> {
    serde_json::from_slice(bytes).map_err(|e| format!("Failed to parse oracle workbook JSON: {e}"))
}

fn extract_from_xlsm(bytes: &[u8]) -> Result<(OracleWorkbook, Vec<ProcedureSummary>), String> {
    let cursor = std::io::Cursor::new(bytes);
    let mut zip = ZipArchive::new(cursor).map_err(|e| format!("Invalid XLSM zip: {e}"))?;

    let sheet_names =
        read_sheet_names_from_workbook_xml(&mut zip).unwrap_or_else(|| vec!["Sheet1".to_string()]);
    let vba_modules = read_vba_modules_from_xlsm(&mut zip)?;

    let workbook = OracleWorkbook {
        schema_version: 1,
        active_sheet: sheet_names.first().cloned(),
        sheets: sheet_names
            .into_iter()
            .map(|name| OracleSheet {
                name,
                cells: BTreeMap::new(),
            })
            .collect(),
        vba_modules: vba_modules.clone(),
    };

    let procedures = list_procedures(&vba_modules)?;
    Ok((workbook, procedures))
}

fn find_zip_entry_case_insensitive<R: Read + Seek>(
    zip: &ZipArchive<R>,
    name: &str,
) -> Option<String> {
    let target = name.trim_start_matches('/').replace('\\', "/");

    for candidate in zip.file_names() {
        let mut normalized = candidate.trim_start_matches('/');
        let replaced;
        if normalized.contains('\\') {
            replaced = normalized.replace('\\', "/");
            normalized = &replaced;
        }

        if normalized.eq_ignore_ascii_case(&target) {
            return Some(candidate.to_string());
        }
    }

    None
}

fn read_zip_entry_bytes<R: Read + Seek>(
    zip: &mut ZipArchive<R>,
    name: &str,
) -> Result<Option<Vec<u8>>, String> {
    let read_entry = |zip: &mut ZipArchive<R>, entry_name: &str| -> Result<Vec<u8>, String> {
        let mut entry = zip
            .by_name(entry_name)
            .map_err(|e| format!("Failed to open zip entry {entry_name}: {e}"))?;
        // Do not trust `ZipFile::size()` for allocation; ZIP metadata is untrusted and can
        // advertise enormous uncompressed sizes (zip-bomb style OOM).
        let mut buf = Vec::new();
        entry
            .read_to_end(&mut buf)
            .map_err(|e| format!("Failed to read zip entry {entry_name}: {e}"))?;
        Ok(buf)
    };

    // Fast path: exact entry name.
    match zip.by_name(name) {
        Ok(mut entry) => {
            // Do not trust `ZipFile::size()` for allocation; ZIP metadata is untrusted and can
            // advertise enormous uncompressed sizes (zip-bomb style OOM).
            let mut buf = Vec::new();
            entry
                .read_to_end(&mut buf)
                .map_err(|e| format!("Failed to read zip entry {name}: {e}"))?;
            return Ok(Some(buf));
        }
        Err(zip::result::ZipError::FileNotFound) => {}
        Err(e) => return Err(format!("Failed to open zip entry {name}: {e}")),
    }

    // Fallback: tolerate a leading `/` and case-only mismatches.
    let Some(actual) = find_zip_entry_case_insensitive(zip, name) else {
        return Ok(None);
    };
    Ok(Some(read_entry(zip, &actual)?))
}

fn read_vba_modules_from_xlsm(
    zip: &mut ZipArchive<std::io::Cursor<&[u8]>>,
) -> Result<Vec<VbaModule>, String> {
    let Some(buf) = read_zip_entry_bytes(zip, "xl/vbaProject.bin")? else {
        return Err("Missing xl/vbaProject.bin".to_string());
    };

    let project = formula_vba::VBAProject::parse(&buf)
        .map_err(|e| format!("Failed to parse vbaProject.bin: {e}"))?;
    Ok(project
        .modules
        .into_iter()
        .map(|m| VbaModule {
            name: m.name,
            code: m.code,
        })
        .collect())
}

fn read_sheet_names_from_workbook_xml(
    zip: &mut ZipArchive<std::io::Cursor<&[u8]>>,
) -> Option<Vec<String>> {
    let buf = read_zip_entry_bytes(zip, "xl/workbook.xml")
        .ok()
        .flatten()?;

    let mut reader = XmlReader::from_reader(buf.as_slice());
    reader.config_mut().trim_text(true);
    let mut temp = Vec::new();

    let mut names = Vec::new();
    loop {
        match reader.read_event_into(&mut temp).ok()? {
            Event::Start(e) | Event::Empty(e) if e.name().as_ref() == b"sheet" => {
                for attr in e.attributes().flatten() {
                    if attr.key.as_ref() == b"name" {
                        if let Ok(val) = attr.unescape_value() {
                            names.push(val.into_owned());
                        }
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        temp.clear();
    }

    if names.is_empty() {
        None
    } else {
        Some(names)
    }
}

fn list_procedures(modules: &[VbaModule]) -> Result<Vec<ProcedureSummary>, String> {
    let combined = modules
        .iter()
        .map(|m| m.code.as_str())
        .collect::<Vec<_>>()
        .join("\n\n");
    let program = parse_program(&combined).map_err(|e| e.to_string())?;
    let mut procedures = program
        .procedures
        .values()
        .map(|p| ProcedureSummary {
            name: p.name.clone(),
            kind: format!("{:?}", p.kind).to_ascii_lowercase(),
        })
        .collect::<Vec<_>>();
    procedures.sort_by(|a, b| a.name.cmp(&b.name));
    Ok(procedures)
}

fn parse_args_json(args: &Option<String>) -> Result<Vec<VbaValue>, String> {
    let Some(raw) = args else {
        return Ok(Vec::new());
    };
    let json: serde_json::Value =
        serde_json::from_str(raw).map_err(|e| format!("Invalid --args JSON: {e}"))?;
    let arr = json
        .as_array()
        .ok_or_else(|| "--args must be a JSON array".to_string())?;
    Ok(arr.iter().map(|v| json_value_to_vba(Some(v))).collect())
}

fn run_macro(workbook: &OracleWorkbook, macro_name: &str, args: &[VbaValue]) -> RunReport {
    let warnings = Vec::new();
    let mut error = None;
    let mut exit_status = 0;

    let modules = workbook.vba_modules.clone();
    let combined = modules
        .iter()
        .map(|m| m.code.as_str())
        .collect::<Vec<_>>()
        .join("\n\n");

    let program = match parse_program(&combined) {
        Ok(p) => p,
        Err(e) => {
            error = Some(e.to_string());
            exit_status = 1;
            return RunReport {
                ok: false,
                macro_name: macro_name.to_string(),
                logs: Vec::new(),
                warnings,
                error,
                exit_status,
                cell_diffs: BTreeMap::new(),
                workbook_after: workbook.clone(),
            };
        }
    };

    let runtime = VbaRuntime::new(program);

    let before_exec = match ExecutionWorkbook::from_oracle_workbook(workbook) {
        Ok(w) => w,
        Err(e) => {
            error = Some(e.to_string());
            exit_status = 1;
            return RunReport {
                ok: false,
                macro_name: macro_name.to_string(),
                logs: Vec::new(),
                warnings,
                error,
                exit_status,
                cell_diffs: BTreeMap::new(),
                workbook_after: workbook.clone(),
            };
        }
    };

    let mut after_exec = before_exec.clone();
    let exec_result: Result<ExecutionResult, VbaError> =
        runtime.execute(&mut after_exec, macro_name, args);
    let logs = after_exec.logs.clone();

    let ok = match exec_result {
        Ok(_) => true,
        Err(e) => {
            error = Some(e.to_string());
            exit_status = 1;
            false
        }
    };

    let cell_diffs = diff_execution_workbooks(&before_exec, &after_exec);
    let workbook_after = after_exec.to_oracle_workbook(modules);

    RunReport {
        ok,
        macro_name: macro_name.to_string(),
        logs,
        warnings,
        error,
        exit_status,
        cell_diffs,
        workbook_after,
    }
}

fn diff_execution_workbooks(
    before: &ExecutionWorkbook,
    after: &ExecutionWorkbook,
) -> BTreeMap<String, BTreeMap<String, OracleCellDiff>> {
    let before_snap = before.snapshot();
    let after_snap = after.snapshot();

    let mut out = BTreeMap::new();
    let sheet_names: BTreeSet<String> = before_snap
        .keys()
        .chain(after_snap.keys())
        .cloned()
        .collect();

    for sheet_name in sheet_names {
        let before_cells = before_snap.get(&sheet_name).cloned().unwrap_or_default();
        let after_cells = after_snap.get(&sheet_name).cloned().unwrap_or_default();
        let addr_set: BTreeSet<String> = before_cells
            .keys()
            .chain(after_cells.keys())
            .cloned()
            .collect();

        let mut sheet_diffs: BTreeMap<String, OracleCellDiff> = BTreeMap::new();
        for addr in addr_set {
            let before_cell = before_cells
                .get(&addr)
                .cloned()
                .unwrap_or(OracleCellSnapshot {
                    value: None,
                    formula: None,
                    format: None,
                });
            let after_cell = after_cells
                .get(&addr)
                .cloned()
                .unwrap_or(OracleCellSnapshot {
                    value: None,
                    formula: None,
                    format: None,
                });

            if before_cell.value == after_cell.value
                && before_cell.formula == after_cell.formula
                && before_cell.format == after_cell.format
            {
                continue;
            }

            sheet_diffs.insert(
                addr,
                OracleCellDiff {
                    before: before_cell,
                    after: after_cell,
                },
            );
        }

        if !sheet_diffs.is_empty() {
            out.insert(sheet_name, sheet_diffs);
        }
    }

    out
}

fn main() {
    let cli = Cli::parse();

    match cli.command {
        Command::Run(args) => {
            let bytes = match read_all_input(&args.input) {
                Ok(b) => b,
                Err(e) => {
                    let report = RunReport {
                        ok: false,
                        macro_name: args.macro_name,
                        logs: Vec::new(),
                        warnings: Vec::new(),
                        error: Some(e),
                        exit_status: 1,
                        cell_diffs: BTreeMap::new(),
                        workbook_after: OracleWorkbook {
                            schema_version: 1,
                            active_sheet: None,
                            sheets: Vec::new(),
                            vba_modules: Vec::new(),
                        },
                    };
                    println!("{}", serde_json::to_string(&report).unwrap());
                    std::process::exit(1);
                }
            };

            let format = match detect_format(&bytes, &args.format) {
                Ok(f) => f,
                Err(e) => {
                    let report = RunReport {
                        ok: false,
                        macro_name: args.macro_name,
                        logs: Vec::new(),
                        warnings: Vec::new(),
                        error: Some(e),
                        exit_status: 1,
                        cell_diffs: BTreeMap::new(),
                        workbook_after: OracleWorkbook {
                            schema_version: 1,
                            active_sheet: None,
                            sheets: Vec::new(),
                            vba_modules: Vec::new(),
                        },
                    };
                    println!("{}", serde_json::to_string(&report).unwrap());
                    std::process::exit(1);
                }
            };

            let (workbook, procedures) = match format {
                InputFormat::Json => match parse_oracle_workbook_json(&bytes) {
                    Ok(wb) => {
                        let procedures = list_procedures(&wb.vba_modules).unwrap_or_default();
                        (wb, procedures)
                    }
                    Err(e) => {
                        let report = RunReport {
                            ok: false,
                            macro_name: args.macro_name,
                            logs: Vec::new(),
                            warnings: Vec::new(),
                            error: Some(e),
                            exit_status: 1,
                            cell_diffs: BTreeMap::new(),
                            workbook_after: OracleWorkbook {
                                schema_version: 1,
                                active_sheet: None,
                                sheets: Vec::new(),
                                vba_modules: Vec::new(),
                            },
                        };
                        println!("{}", serde_json::to_string(&report).unwrap());
                        std::process::exit(1);
                    }
                },
                InputFormat::Xlsm => {
                    let password = match resolve_password(&args.password, &args.password_file) {
                        Ok(pw) => pw,
                        Err(e) => {
                            let report = RunReport {
                                ok: false,
                                macro_name: args.macro_name,
                                logs: Vec::new(),
                                warnings: Vec::new(),
                                error: Some(e),
                                exit_status: 1,
                                cell_diffs: BTreeMap::new(),
                                workbook_after: OracleWorkbook {
                                    schema_version: 1,
                                    active_sheet: None,
                                    sheets: Vec::new(),
                                    vba_modules: Vec::new(),
                                },
                            };
                            println!("{}", serde_json::to_string(&report).unwrap());
                            std::process::exit(1);
                        }
                    };

                    let decrypted = match maybe_decrypt_encrypted_xlsm(&bytes, password.as_deref()) {
                        Ok(b) => b,
                        Err(e) => {
                            let report = RunReport {
                                ok: false,
                                macro_name: args.macro_name,
                                logs: Vec::new(),
                                warnings: Vec::new(),
                                error: Some(e),
                                exit_status: 1,
                                cell_diffs: BTreeMap::new(),
                                workbook_after: OracleWorkbook {
                                    schema_version: 1,
                                    active_sheet: None,
                                    sheets: Vec::new(),
                                    vba_modules: Vec::new(),
                                },
                            };
                            println!("{}", serde_json::to_string(&report).unwrap());
                            std::process::exit(1);
                        }
                    };

                    match extract_from_xlsm(decrypted.as_ref()) {
                        Ok((wb, procs)) => (wb, procs),
                        Err(e) => {
                            let report = RunReport {
                                ok: false,
                                macro_name: args.macro_name,
                                logs: Vec::new(),
                                warnings: Vec::new(),
                                error: Some(e),
                                exit_status: 1,
                                cell_diffs: BTreeMap::new(),
                                workbook_after: OracleWorkbook {
                                    schema_version: 1,
                                    active_sheet: None,
                                    sheets: Vec::new(),
                                    vba_modules: Vec::new(),
                                },
                            };
                            println!("{}", serde_json::to_string(&report).unwrap());
                            std::process::exit(1);
                        }
                    }
                }
            };

            let args_values = match parse_args_json(&args.args) {
                Ok(v) => v,
                Err(e) => {
                    let report = RunReport {
                        ok: false,
                        macro_name: args.macro_name,
                        logs: Vec::new(),
                        warnings: Vec::new(),
                        error: Some(e),
                        exit_status: 1,
                        cell_diffs: BTreeMap::new(),
                        workbook_after: workbook,
                    };
                    println!("{}", serde_json::to_string(&report).unwrap());
                    std::process::exit(1);
                }
            };

            // Warn if macro isn't even in the parsed program; this helps humans, but we keep running
            // so we still get a deterministic error from the runtime.
            if !procedures
                .iter()
                .any(|p| p.name.eq_ignore_ascii_case(&args.macro_name))
            {
                // Don't treat as hard error; VBA is case-insensitive and runtime error message is good enough.
            }

            let report = run_macro(&workbook, &args.macro_name, &args_values);
            println!("{}", serde_json::to_string(&report).unwrap());
            if !report.ok {
                std::process::exit(report.exit_status);
            }
        }
        Command::Extract(args) => {
            let bytes = match read_all_input(&args.input) {
                Ok(b) => b,
                Err(e) => {
                    let report = ExtractReport {
                        ok: false,
                        error: Some(e),
                        workbook: OracleWorkbook {
                            schema_version: 1,
                            active_sheet: None,
                            sheets: Vec::new(),
                            vba_modules: Vec::new(),
                        },
                        procedures: Vec::new(),
                    };
                    println!("{}", serde_json::to_string(&report).unwrap());
                    std::process::exit(1);
                }
            };

            let format = match detect_format(&bytes, &args.format) {
                Ok(f) => f,
                Err(e) => {
                    let report = ExtractReport {
                        ok: false,
                        error: Some(e),
                        workbook: OracleWorkbook {
                            schema_version: 1,
                            active_sheet: None,
                            sheets: Vec::new(),
                            vba_modules: Vec::new(),
                        },
                        procedures: Vec::new(),
                    };
                    println!("{}", serde_json::to_string(&report).unwrap());
                    std::process::exit(1);
                }
            };

            let report = match format {
                InputFormat::Json => match parse_oracle_workbook_json(&bytes) {
                    Ok(workbook) => match list_procedures(&workbook.vba_modules) {
                        Ok(procedures) => ExtractReport {
                            ok: true,
                            error: None,
                            workbook,
                            procedures,
                        },
                        Err(e) => ExtractReport {
                            ok: false,
                            error: Some(e),
                            workbook,
                            procedures: Vec::new(),
                        },
                    },
                    Err(e) => ExtractReport {
                        ok: false,
                        error: Some(e),
                        workbook: OracleWorkbook {
                            schema_version: 1,
                            active_sheet: None,
                            sheets: Vec::new(),
                            vba_modules: Vec::new(),
                        },
                        procedures: Vec::new(),
                    },
                },
                InputFormat::Xlsm => {
                    match resolve_password(&args.password, &args.password_file) {
                        Ok(password) => {
                            match maybe_decrypt_encrypted_xlsm(&bytes, password.as_deref()) {
                                Ok(decrypted) => match extract_from_xlsm(decrypted.as_ref()) {
                                    Ok((workbook, procedures)) => ExtractReport {
                                        ok: true,
                                        error: None,
                                        workbook,
                                        procedures,
                                    },
                                    Err(e) => ExtractReport {
                                        ok: false,
                                        error: Some(e),
                                        workbook: OracleWorkbook {
                                            schema_version: 1,
                                            active_sheet: None,
                                            sheets: Vec::new(),
                                            vba_modules: Vec::new(),
                                        },
                                        procedures: Vec::new(),
                                    },
                                },
                                Err(e) => ExtractReport {
                                    ok: false,
                                    error: Some(e),
                                    workbook: OracleWorkbook {
                                        schema_version: 1,
                                        active_sheet: None,
                                        sheets: Vec::new(),
                                        vba_modules: Vec::new(),
                                    },
                                    procedures: Vec::new(),
                                },
                            }
                        }
                        Err(e) => ExtractReport {
                            ok: false,
                            error: Some(e),
                            workbook: OracleWorkbook {
                                schema_version: 1,
                                active_sheet: None,
                                sheets: Vec::new(),
                                vba_modules: Vec::new(),
                            },
                            procedures: Vec::new(),
                        },
                    }
                }
            };

            println!("{}", serde_json::to_string(&report).unwrap());
            if !report.ok {
                std::process::exit(1);
            }
        }
    }
}
