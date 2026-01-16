#![no_main]

use chrono::TimeZone;
use libfuzzer_sys::fuzz_target;

use formula_engine::eval::{CellAddr, EvalContext, Evaluator, RecalcContext, ResolvedName, ValueResolver};
use formula_engine::{ErrorKind, Value};

/// Keep evaluation fuzzing bounded: we want to exercise many inputs quickly, without allowing
/// pathological inputs to drive very large allocations.
const MAX_EVAL_FORMULA_CHARS: usize = 2_048;
const MAX_INPUT_BYTES: usize = MAX_EVAL_FORMULA_CHARS * 4; // max UTF-8 bytes per char
/// Keep in sync with `formula_engine::eval::evaluator::MAX_MATERIALIZED_ARRAY_CELLS`.
const MAX_MATERIALIZED_ARRAY_CELLS: usize = 5_000_000;

fn truncate_to_chars(s: &str, max_chars: usize) -> &str {
    let mut count = 0usize;
    for (idx, _) in s.char_indices() {
        if count == max_chars {
            return &s[..idx];
        }
        count += 1;
    }
    s
}

#[derive(Debug, Clone)]
struct FuzzResolver {
    sheet_names: [&'static str; 3],
    rows: u32,
    cols: u32,
}

impl Default for FuzzResolver {
    fn default() -> Self {
        Self {
            sheet_names: ["Sheet1", "Sheet2", "Weird Sheet"],
            // Keep dimensions small enough to keep fuzzing fast and avoid materializing very large
            // arrays from plain range references.
            rows: 1_024,
            cols: 256,
        }
    }
}

impl FuzzResolver {
    fn sheet_id_by_name(&self, name: &str) -> Option<usize> {
        let name = name.trim();
        self.sheet_names
            .iter()
            .position(|s| s.eq_ignore_ascii_case(name))
    }

    fn stable_cell_number(&self, sheet_id: usize, addr: CellAddr) -> f64 {
        // A small reversible-ish mixer to create deterministic, varied cell values without
        // allocating per-cell strings.
        let mut x = (sheet_id as u64).wrapping_mul(0x9e3779b97f4a7c15);
        x ^= (addr.row as u64).wrapping_mul(0xbf58476d1ce4e5b9);
        x ^= (addr.col as u64).wrapping_mul(0x94d049bb133111eb);
        x ^= x >> 27;
        ((x % 10_000) as f64) / 10.0
    }
}

impl ValueResolver for FuzzResolver {
    fn sheet_exists(&self, sheet_id: usize) -> bool {
        sheet_id < self.sheet_names.len()
    }

    fn sheet_count(&self) -> usize {
        self.sheet_names.len()
    }

    fn sheet_dimensions(&self, _sheet_id: usize) -> (u32, u32) {
        (self.rows, self.cols)
    }

    fn get_cell_value(&self, sheet_id: usize, addr: CellAddr) -> Value {
        // A few fixed "interesting" cells to exercise non-numeric code paths without turning
        // large materializations into giant string allocations.
        if addr.row == 0 && addr.col == 0 {
            return Value::Text("hello".to_string());
        }
        if addr.row == 0 && addr.col == 1 {
            return Value::Bool(true);
        }
        if addr.row == 0 && addr.col == 2 {
            return Value::Error(ErrorKind::Div0);
        }
        Value::Number(self.stable_cell_number(sheet_id, addr))
    }

    fn sheet_name(&self, sheet_id: usize) -> Option<&str> {
        self.sheet_names.get(sheet_id).copied()
    }

    fn sheet_id(&self, name: &str) -> Option<usize> {
        self.sheet_id_by_name(name)
    }

    fn get_external_value(&self, sheet: &str, addr: CellAddr) -> Option<Value> {
        // Deterministic but stable external values.
        let mut hash: u64 = 0xcbf29ce484222325;
        for b in sheet.as_bytes() {
            hash ^= *b as u64;
            hash = hash.wrapping_mul(0x100000001b3);
        }
        hash ^= addr.row as u64;
        hash = hash.wrapping_mul(0x100000001b3);
        hash ^= addr.col as u64;
        Some(Value::Number(((hash % 10_000) as f64) / 100.0))
    }

    fn resolve_structured_ref(
        &self,
        _ctx: EvalContext,
        _sref: &formula_engine::structured_refs::StructuredRef,
    ) -> Result<Vec<(usize, CellAddr, CellAddr)>, ErrorKind> {
        Ok(Vec::new())
    }

    fn resolve_name(&self, sheet_id: usize, name: &str) -> Option<ResolvedName> {
        if !self.sheet_exists(sheet_id) {
            return None;
        }
        let name = name.trim();
        if name.eq_ignore_ascii_case("PI") {
            return Some(ResolvedName::Constant(Value::Number(std::f64::consts::PI)));
        }
        if name.eq_ignore_ascii_case("FOO") {
            return Some(ResolvedName::Constant(Value::Number(42.0)));
        }
        if name.eq_ignore_ascii_case("BAR") {
            return Some(ResolvedName::Constant(Value::Text("bar".to_string())));
        }
        None
    }
}

fuzz_target!(|data: &[u8]| {
    if data.is_empty() {
        return;
    }

    let data = if data.len() > MAX_INPUT_BYTES {
        &data[..MAX_INPUT_BYTES]
    } else {
        data
    };

    let input = String::from_utf8_lossy(data);
    let formula = truncate_to_chars(&input, MAX_EVAL_FORMULA_CHARS);

    let selector = data[0];
    let selector2 = data.get(1).copied().unwrap_or(0);
    let selector3 = data.get(2).copied().unwrap_or(0);

    let resolver = FuzzResolver::default();
    let current_sheet = (selector as usize) % resolver.sheet_count();
    let (rows, cols) = resolver.sheet_dimensions(current_sheet);
    let current_cell = CellAddr {
        row: u32::from(selector2) % rows.max(1),
        col: u32::from(selector3) % cols.max(1),
    };

    // Parse in the canonical parser, and optionally normalize relative refs to the current cell so
    // R1C1-style offset code paths get exercised during compilation/evaluation.
    let parse_opts = formula_engine::ParseOptions {
        locale: formula_engine::LocaleConfig::en_us(),
        reference_style: if selector & 1 == 0 {
            formula_engine::ReferenceStyle::A1
        } else {
            formula_engine::ReferenceStyle::R1C1
        },
        normalize_relative_to: if selector & 2 == 0 {
            None
        } else {
            Some(formula_engine::CellAddr::new(current_cell.row, current_cell.col))
        },
    };

    let ast = match formula_engine::parse_formula(formula, parse_opts) {
        Ok(ast) => ast,
        Err(_) => return,
    };

    // Compile to the evaluation IR (sheet ids, normalized references).
    let mut resolve_sheet = |name: &str| resolver.sheet_id(name);
    let mut sheet_dimensions = |sheet_id: usize| resolver.sheet_dimensions(sheet_id);
    let compiled = formula_engine::eval::compile_canonical_expr(
        &ast.expr,
        current_sheet,
        current_cell,
        &mut resolve_sheet,
        &mut sheet_dimensions,
    );

    // Deterministic recalc context.
    let mut id_bytes = [0u8; 8];
    for (dst, src) in id_bytes.iter_mut().zip(data.iter().copied()) {
        *dst = src;
    }
    let mut recalc_ctx = RecalcContext::new(u64::from_le_bytes(id_bytes));
    recalc_ctx.now_utc = chrono::Utc.timestamp_opt(0, 0).single().unwrap();

    let eval_ctx = EvalContext {
        current_sheet,
        current_cell,
    };
    let evaluator = Evaluator::new(&resolver, eval_ctx, &recalc_ctx);
    let value = evaluator.eval_formula(&compiled);

    // Guard: evaluation should never materialize a dynamic array larger than the engine limit.
    if let Value::Array(arr) = &value {
        let total = arr.rows.checked_mul(arr.cols).unwrap_or(usize::MAX);
        assert!(
            total <= MAX_MATERIALIZED_ARRAY_CELLS,
            "materialized array exceeded MAX_MATERIALIZED_ARRAY_CELLS: {total}"
        );
    }

    std::hint::black_box(value);
});
