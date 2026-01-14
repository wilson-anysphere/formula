#![no_main]

use libfuzzer_sys::fuzz_target;

/// Keep the harness itself bounded.
///
/// The parser enforces Excel's 8,192-character display limit internally, but fuzzing should also
/// avoid passing arbitrarily-large inputs into lossy UTF-8 conversion / tokenization.
const EXCEL_MAX_FORMULA_CHARS: usize = 8_192;
const MAX_FUZZ_FORMULA_CHARS: usize = EXCEL_MAX_FORMULA_CHARS + 256;
const MAX_INPUT_BYTES: usize = MAX_FUZZ_FORMULA_CHARS * 4; // max UTF-8 bytes per char

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

fuzz_target!(|data: &[u8]| {
    if data.is_empty() {
        return;
    }

    // Avoid extremely large allocations even before we reach the parser's Excel limits.
    let data = if data.len() > MAX_INPUT_BYTES {
        &data[..MAX_INPUT_BYTES]
    } else {
        data
    };

    // Accept arbitrary bytes as input; treat invalid UTF-8 lossy.
    let input = String::from_utf8_lossy(data);
    let formula = truncate_to_chars(&input, MAX_FUZZ_FORMULA_CHARS);

    // Vary parse options to explore A1/R1C1 and locale separator code paths.
    let selector = data[0];
    let selector2 = data.get(1).copied().unwrap_or(0);

    let locale_cfg = if selector & 0b10 == 0 {
        formula_engine::LocaleConfig::en_us()
    } else {
        formula_engine::LocaleConfig::de_de()
    };
    let reference_style = if selector & 0b1 == 0 {
        formula_engine::ReferenceStyle::A1
    } else {
        formula_engine::ReferenceStyle::R1C1
    };
    let normalize_relative_to = if selector & 0b100 == 0 {
        None
    } else {
        Some(formula_engine::CellAddr::new(
            u32::from(selector2) % 128,
            u32::from(selector2.rotate_left(3)) % 128,
        ))
    };

    let parse_opts = formula_engine::ParseOptions {
        locale: locale_cfg.clone(),
        reference_style,
        normalize_relative_to,
    };

    let parsed = formula_engine::parse_formula(formula, parse_opts.clone());
    if let Ok(ast) = parsed {
        // Exercise round-trip serialization (canonicalizer-ish) and re-parse.
        let serialize_opts = formula_engine::SerializeOptions {
            locale: locale_cfg.clone(),
            reference_style,
            include_xlfn_prefix: true,
            origin: normalize_relative_to,
            omit_equals: false,
        };
        if let Ok(serialized) = ast.to_string(serialize_opts) {
            let _ = formula_engine::parse_formula(&serialized, parse_opts);
        } else {
            let _ = ast.to_string(formula_engine::SerializeOptions::default());
        }
    }

    // Exercise the evaluation parser lowering path too (uses canonical parser internally).
    let _ = formula_engine::eval::Parser::parse(formula);

    // Exercise locale translation (canonicalize/localize).
    let locale = if selector & 0b1000 == 0 {
        &formula_engine::locale::EN_US
    } else {
        &formula_engine::locale::DE_DE
    };
    if let Ok(canon) =
        formula_engine::locale::canonicalize_formula_with_style(formula, locale, reference_style)
    {
        let _ = formula_engine::locale::localize_formula_with_style(&canon, locale, reference_style);
    }
});

