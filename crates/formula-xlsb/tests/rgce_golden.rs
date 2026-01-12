use std::collections::BTreeMap;
use std::fs;
use std::path::{Path, PathBuf};
use std::sync::OnceLock;

use formula_xlsb::rgce::{decode_rgce_with_context_and_rgcb_and_base, CellCoord};
use formula_xlsb::workbook_context::WorkbookContext;
use formula_xlsb::XlsbWorkbook;
use pretty_assertions::assert_eq;
use serde::Deserialize;

#[derive(Debug, Deserialize)]
struct GoldenCase {
    name: String,
    rgce_hex: String,
    #[serde(default)]
    rgcb_hex: String,
    expected: String,
    #[serde(default)]
    ctx: CaseContext,
    #[serde(default)]
    ctx_fixture: Option<String>,
    #[serde(default)]
    base: Option<BaseCell>,
}

#[derive(Debug, Default, Deserialize)]
struct CaseContext {
    #[serde(default)]
    extern_sheets: Vec<ExternSheetEntry>,
    #[serde(default)]
    workbook_names: Vec<WorkbookNameEntry>,
    #[serde(default)]
    sheet_names: Vec<SheetNameEntry>,
}

#[derive(Debug, Deserialize)]
struct ExternSheetEntry {
    first: String,
    last: String,
    ixti: u16,
}

#[derive(Debug, Deserialize)]
struct WorkbookNameEntry {
    name: String,
    index: u32,
}

#[derive(Debug, Deserialize)]
struct SheetNameEntry {
    sheet: String,
    name: String,
    index: u32,
}

#[derive(Debug, Deserialize)]
struct BaseCell {
    row: u32,
    col: u32,
}

fn parse_hex_bytes(raw: &str) -> Result<Vec<u8>, String> {
    let raw = raw.trim();
    if raw.is_empty() {
        return Ok(Vec::new());
    }
    raw.split_whitespace()
        .map(|tok| {
            let tok = tok.strip_prefix("0x").unwrap_or(tok);
            u8::from_str_radix(tok, 16).map_err(|err| format!("invalid hex byte {tok:?}: {err}"))
        })
        .collect()
}

fn load_fixture_context(rel_path: &str) -> WorkbookContext {
    static CACHE: OnceLock<std::sync::Mutex<BTreeMap<PathBuf, WorkbookContext>>> = OnceLock::new();
    let cache = CACHE.get_or_init(|| std::sync::Mutex::new(BTreeMap::new()));

    let manifest_dir = Path::new(env!("CARGO_MANIFEST_DIR"));
    let path = manifest_dir.join(rel_path);

    if let Some(ctx) = cache.lock().expect("lock").get(&path).cloned() {
        return ctx;
    }

    let wb = XlsbWorkbook::open(&path).expect("open fixture workbook");
    let ctx = wb.workbook_context().clone();
    cache
        .lock()
        .expect("lock")
        .insert(path, ctx.clone());
    ctx
}

#[test]
fn rgce_golden_cases_decode_exactly() {
    let manifest_dir = Path::new(env!("CARGO_MANIFEST_DIR"));
    let corpus_path = manifest_dir.join("tests/rgce_golden_cases.jsonl");
    let corpus = fs::read_to_string(&corpus_path).expect("read golden corpus");

    for (idx, line) in corpus.lines().enumerate() {
        let line = line.trim();
        if line.is_empty() {
            continue;
        }

        let case: GoldenCase =
            serde_json::from_str(line).unwrap_or_else(|err| panic!("invalid corpus JSONL line {}: {err}", idx + 1));

        let rgce = parse_hex_bytes(&case.rgce_hex)
            .unwrap_or_else(|err| panic!("case {}: invalid rgce_hex: {err}", case.name));
        let rgcb = parse_hex_bytes(&case.rgcb_hex)
            .unwrap_or_else(|err| panic!("case {}: invalid rgcb_hex: {err}", case.name));

        let mut ctx = match &case.ctx_fixture {
            Some(rel) => load_fixture_context(rel),
            None => WorkbookContext::default(),
        };
        for entry in &case.ctx.extern_sheets {
            ctx.add_extern_sheet(entry.first.clone(), entry.last.clone(), entry.ixti);
        }
        for entry in &case.ctx.workbook_names {
            ctx.add_workbook_name(entry.name.clone(), entry.index);
        }
        for entry in &case.ctx.sheet_names {
            ctx.add_sheet_name(entry.sheet.clone(), entry.name.clone(), entry.index);
        }

        let base = case
            .base
            .as_ref()
            .map(|b| CellCoord::new(b.row, b.col))
            .unwrap_or_else(|| CellCoord::new(0, 0));

        let decoded =
            decode_rgce_with_context_and_rgcb_and_base(&rgce, &rgcb, &ctx, base).unwrap_or_else(|err| {
                panic!(
                    "case {} failed to decode: {err} (ptg={:?}, offset={})",
                    case.name,
                    err.ptg(),
                    err.offset()
                )
            });

        assert_eq!(decoded, case.expected, "case {} (line {})", case.name, idx + 1);
    }
}
