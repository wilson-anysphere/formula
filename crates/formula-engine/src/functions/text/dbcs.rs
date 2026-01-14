//! Legacy DBCS / byte-count text functions.
//!
//! Excel exposes `*B` variants of several text functions (LENB, LEFTB, MIDB, RIGHTB,
//! FINDB, SEARCHB, REPLACEB). In DBCS locales (e.g. Japanese), these functions
//! operate on *byte counts* instead of character counts, and the definition of a
//! "byte" depends on the active workbook locale / code page.
//!
//! The engine models a workbook-level "text codepage" (default: 1252 / en-US). Under
//! single-byte codepages, the `*B` functions behave identically to their non-`B`
//! equivalents.
//!
//! `ASC` / `DBCS` perform half-width / full-width conversions in Japanese locales.
//! We implement these conversions when the active workbook text codepage is a
//! DBCS codepage:
//! - 932 (Shift_JIS / Japanese): fullwidth/halfwidth ASCII + symbols + katakana
//!   conversions.
//! - 936/949/950 (Chinese/Korean): fullwidth/halfwidth ASCII + symbols
//!   conversions.
//!
//! In non-DBCS codepages, they behave as identity transforms.
//!
//! `PHONETIC` depends on per-cell phonetic guide metadata (furigana).
//! When phonetic metadata is present for a referenced cell, `PHONETIC(reference)`
//! returns that stored string. When phonetic metadata is absent (the common
//! case), Excel falls back to the referenced cell’s displayed text, so the
//! engine returns the referenced value coerced to text using the current
//! locale-aware formatting rules.
//!
//! Note: Excel's DBCS semantics contain many locale-specific edge cases. This module implements
//! the core behaviors needed by typical Japanese/Chinese/Korean workbooks, but may need to be
//! extended as additional Excel oracle cases are added.

use crate::eval::CompiledExpr;
use crate::eval::MAX_MATERIALIZED_ARRAY_CELLS;
use crate::functions::array_lift;
use crate::functions::{call_function, ArgValue, FunctionContext, Reference};
use crate::value::{Array, ErrorKind, Value};
use encoding_rs::{
    Encoding, BIG5, EUC_KR, GBK, SHIFT_JIS, UTF_8, WINDOWS_1250, WINDOWS_1251, WINDOWS_1252,
    WINDOWS_1253, WINDOWS_1254, WINDOWS_1255, WINDOWS_1256, WINDOWS_1257, WINDOWS_1258,
    WINDOWS_874,
};

pub(crate) fn findb_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let codepage = ctx.text_codepage();
    if !matches!(codepage, 932 | 936 | 949 | 950) {
        // Single-byte locales: byte positions match character positions.
        return call_function(ctx, "FIND", args);
    }

    let needle = array_lift::eval_arg(ctx, &args[0]);
    let haystack = array_lift::eval_arg(ctx, &args[1]);
    let start = if args.len() >= 3 {
        array_lift::eval_arg(ctx, &args[2])
    } else {
        Value::Number(1.0)
    };

    array_lift::lift3(needle, haystack, start, |needle, haystack, start| {
        let needle = needle.coerce_to_string_with_ctx(ctx)?;
        let haystack = haystack.coerce_to_string_with_ctx(ctx)?;
        let start = start.coerce_to_i64_with_ctx(ctx)?;
        Ok(findb_impl(codepage, &needle, &haystack, start, false))
    })
}

pub(crate) fn searchb_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let codepage = ctx.text_codepage();
    if !matches!(codepage, 932 | 936 | 949 | 950) {
        // Single-byte locales: byte positions match character positions.
        return call_function(ctx, "SEARCH", args);
    }

    let needle = array_lift::eval_arg(ctx, &args[0]);
    let haystack = array_lift::eval_arg(ctx, &args[1]);
    let start = if args.len() >= 3 {
        array_lift::eval_arg(ctx, &args[2])
    } else {
        Value::Number(1.0)
    };

    array_lift::lift3(needle, haystack, start, |needle, haystack, start| {
        let needle = needle.coerce_to_string_with_ctx(ctx)?;
        let haystack = haystack.coerce_to_string_with_ctx(ctx)?;
        let start = start.coerce_to_i64_with_ctx(ctx)?;
        Ok(findb_impl(codepage, &needle, &haystack, start, true))
    })
}

pub(crate) fn replaceb_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let codepage = ctx.text_codepage();
    if !matches!(codepage, 932 | 936 | 949 | 950) {
        // Single-byte locales: byte positions match character positions.
        return call_function(ctx, "REPLACE", args);
    }

    let old_text = array_lift::eval_arg(ctx, &args[0]);
    let start_num = array_lift::eval_arg(ctx, &args[1]);
    let num_bytes = array_lift::eval_arg(ctx, &args[2]);
    let new_text = array_lift::eval_arg(ctx, &args[3]);

    array_lift::lift4(
        old_text,
        start_num,
        num_bytes,
        new_text,
        |old, start, num, new| {
            let old = old.coerce_to_string_with_ctx(ctx)?;
            let start = start.coerce_to_i64_with_ctx(ctx)?;
            let num = num.coerce_to_i64_with_ctx(ctx)?;
            let new = new.coerce_to_string_with_ctx(ctx)?;
            if start < 1 || num < 0 {
                return Err(ErrorKind::Value);
            }
            let start0 = (start - 1) as usize;
            let num = usize::try_from(num).unwrap_or(usize::MAX);
            Ok(Value::Text(replaceb_bytes(
                codepage, &old, start0, num, &new,
            )))
        },
    )
}

pub(crate) fn leftb_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let codepage = ctx.text_codepage();
    if !matches!(codepage, 932 | 936 | 949 | 950) {
        // Single-byte locales: byte counts match character counts.
        return call_function(ctx, "LEFT", args);
    }

    let text = array_lift::eval_arg(ctx, &args[0]);
    let n = if args.len() == 2 {
        array_lift::eval_arg(ctx, &args[1])
    } else {
        Value::Number(1.0)
    };
    array_lift::lift2(text, n, |text, n| {
        let text = text.coerce_to_string_with_ctx(ctx)?;
        let n = n.coerce_to_i64_with_ctx(ctx)?;
        if n < 0 {
            return Err(ErrorKind::Value);
        }
        let n = usize::try_from(n).unwrap_or(usize::MAX);
        Ok(Value::Text(slice_bytes_dbcs(codepage, &text, 0, n)))
    })
}

pub(crate) fn rightb_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let codepage = ctx.text_codepage();
    if !matches!(codepage, 932 | 936 | 949 | 950) {
        // Single-byte locales: byte counts match character counts.
        return call_function(ctx, "RIGHT", args);
    }

    let text = array_lift::eval_arg(ctx, &args[0]);
    let n = if args.len() == 2 {
        array_lift::eval_arg(ctx, &args[1])
    } else {
        Value::Number(1.0)
    };
    array_lift::lift2(text, n, |text, n| {
        let text = text.coerce_to_string_with_ctx(ctx)?;
        let n = n.coerce_to_i64_with_ctx(ctx)?;
        if n < 0 {
            return Err(ErrorKind::Value);
        }
        let n = usize::try_from(n).unwrap_or(usize::MAX);
        let total = encoded_byte_prefixes(codepage, &text)
            .last()
            .copied()
            .unwrap_or(0);
        let start0 = total.saturating_sub(n);
        Ok(Value::Text(slice_bytes_dbcs(codepage, &text, start0, n)))
    })
}

pub(crate) fn midb_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let codepage = ctx.text_codepage();
    if !matches!(codepage, 932 | 936 | 949 | 950) {
        // Single-byte locales: byte counts match character counts.
        return call_function(ctx, "MID", args);
    }

    let text = array_lift::eval_arg(ctx, &args[0]);
    let start = array_lift::eval_arg(ctx, &args[1]);
    let len = array_lift::eval_arg(ctx, &args[2]);
    array_lift::lift3(text, start, len, |text, start, len| {
        let text = text.coerce_to_string_with_ctx(ctx)?;
        let start = start.coerce_to_i64_with_ctx(ctx)?;
        let len = len.coerce_to_i64_with_ctx(ctx)?;
        if start < 1 || len < 0 {
            return Err(ErrorKind::Value);
        }
        let start0 = (start - 1) as usize;
        let len = usize::try_from(len).unwrap_or(usize::MAX);
        Ok(Value::Text(slice_bytes_dbcs(codepage, &text, start0, len)))
    })
}

pub(crate) fn lenb_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    // Excel's `LENB` only diverges from `LEN` in DBCS locales. For single-byte locales/codepages
    // (including the engine default: 1252 / en-US), `LENB` behaves like `LEN`.
    if !is_dbcs_codepage(ctx.text_codepage()) {
        return call_function(ctx, "LEN", args);
    }

    let codepage = ctx.text_codepage();
    let text = array_lift::eval_arg(ctx, &args[0]);
    array_lift::lift1(text, |text| {
        let s = text.coerce_to_string_with_ctx(ctx)?;
        Ok(Value::Number(encode_bytes_len(codepage, &s) as f64))
    })
}

fn is_dbcs_codepage(codepage: u16) -> bool {
    matches!(codepage, 932 | 936 | 949 | 950)
}

pub(crate) fn asc_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let text = array_lift::eval_arg(ctx, &args[0]);
    let codepage = ctx.text_codepage();
    let dbcs_codepage = matches!(codepage, 932 | 936 | 949 | 950);
    let cp932 = codepage == 932;
    array_lift::lift1(text, |text| {
        let s = text.coerce_to_string_with_ctx(ctx)?;
        if !dbcs_codepage {
            return Ok(Value::Text(s));
        }
        if cp932 {
            Ok(Value::Text(asc_cp932(&s)))
        } else {
            Ok(Value::Text(asc_dbcs_basic(&s)))
        }
    })
}

pub(crate) fn dbcs_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let text = array_lift::eval_arg(ctx, &args[0]);
    let codepage = ctx.text_codepage();
    let dbcs_codepage = matches!(codepage, 932 | 936 | 949 | 950);
    let cp932 = codepage == 932;
    array_lift::lift1(text, |text| {
        let s = text.coerce_to_string_with_ctx(ctx)?;
        if !dbcs_codepage {
            return Ok(Value::Text(s));
        }
        if cp932 {
            Ok(Value::Text(dbcs_cp932(&s)))
        } else {
            Ok(Value::Text(dbcs_dbcs_basic(&s)))
        }
    })
}

pub(crate) fn phonetic_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    match ctx.eval_arg(&args[0]) {
        ArgValue::Reference(reference) => phonetic_from_reference(ctx, reference),
        // References passed through higher-order constructs like LAMBDA/LET can appear as scalar
        // `Value::Reference` values. Treat them like normal reference arguments.
        ArgValue::Scalar(Value::Reference(reference)) => phonetic_from_reference(ctx, reference),
        ArgValue::Scalar(Value::ReferenceUnion(_)) => Value::Error(ErrorKind::Value),
        // Excel accepts non-reference inputs. From Microsoft's documentation:
        // https://support.microsoft.com/en-us/office/phonetic-function-9a329dac-0c0f-42f8-9a55-639086988554
        //
        // > Reference: Required. Text string or a reference to a single cell or a range of cells
        // > that contain a furigana text string.
        //
        // When no per-cell phonetic metadata exists (e.g. literals like `"abc"` or computed
        // scalars), PHONETIC behaves like a text coercion.
        ArgValue::Scalar(value) => array_lift::lift1(value, |v| {
            Ok(Value::Text(v.coerce_to_string_with_ctx(ctx)?))
        }),
        ArgValue::ReferenceUnion(_) => Value::Error(ErrorKind::Value),
    }
}

fn phonetic_from_reference(ctx: &dyn FunctionContext, reference: Reference) -> Value {
    let reference = reference.normalized();
    ctx.record_reference(&reference);

    if reference.is_single_cell() {
        let cell_value = ctx.get_cell_value(&reference.sheet_id, reference.start);
        if let Value::Error(e) = &cell_value {
            return Value::Error(*e);
        }
        if let Some(phonetic) = ctx.get_cell_phonetic(&reference.sheet_id, reference.start) {
            return Value::Text(phonetic.to_string());
        }
        return match cell_value.coerce_to_string_with_ctx(ctx) {
            Ok(s) => Value::Text(s),
            Err(e) => Value::Error(e),
        };
    }

    // Preserve the existing array/broadcast behavior for multi-cell references.
    let rows = (reference.end.row - reference.start.row + 1) as usize;
    let cols = (reference.end.col - reference.start.col + 1) as usize;
    let total = match rows.checked_mul(cols) {
        Some(v) => v,
        None => return Value::Error(ErrorKind::Spill),
    };
    if total > MAX_MATERIALIZED_ARRAY_CELLS {
        return Value::Error(ErrorKind::Spill);
    }
    let mut out = Vec::new();
    if out.try_reserve_exact(total).is_err() {
        return Value::Error(ErrorKind::Num);
    }
    for addr in reference.iter_cells() {
        let cell_value = ctx.get_cell_value(&reference.sheet_id, addr);
        // Error values are preserved per element (matching `array_lift` behavior).
        if let Value::Error(e) = cell_value {
            out.push(Value::Error(e));
            continue;
        }
        if let Some(phonetic) = ctx.get_cell_phonetic(&reference.sheet_id, addr) {
            out.push(Value::Text(phonetic.to_string()));
            continue;
        }
        out.push(match cell_value.coerce_to_string_with_ctx(ctx) {
            Ok(s) => Value::Text(s),
            Err(e) => Value::Error(e),
        });
    }
    Value::Array(Array::new(rows, cols, out))
}

const FULLWIDTH_SPACE: char = '\u{3000}';
const HALFWIDTH_DAKUTEN: char = '\u{FF9E}';
const HALFWIDTH_HANDAKUTEN: char = '\u{FF9F}';
const FULLWIDTH_CENT: char = '\u{FFE0}';
const FULLWIDTH_POUND: char = '\u{FFE1}';
const FULLWIDTH_NOT: char = '\u{FFE2}';
const FULLWIDTH_MACRON: char = '\u{FFE3}';
const FULLWIDTH_BROKEN_BAR: char = '\u{FFE4}';
const FULLWIDTH_YEN: char = '\u{FFE5}';
const FULLWIDTH_WON: char = '\u{FFE6}';
const COMBINING_DAKUTEN: char = '\u{3099}';
const COMBINING_HANDAKUTEN: char = '\u{309A}';

fn asc_cp932(input: &str) -> String {
    let mut out = String::with_capacity(input.len());
    let mut iter = input.chars().peekable();

    while let Some(ch) = iter.next() {
        if ch == FULLWIDTH_SPACE {
            out.push(' ');
            continue;
        }

        // Fullwidth ASCII variants: U+FF01..U+FF5E -> U+0021..U+007E.
        if ('\u{FF01}'..='\u{FF5E}').contains(&ch) {
            let ascii = char::from_u32((ch as u32).saturating_sub(0xFEE0))
                .expect("FF01..FF5E - 0xFEE0 must remain in Unicode scalar range");
            out.push(ascii);
            continue;
        }

        if let Some(mapped) = fullwidth_symbol_to_halfwidth(ch) {
            out.push(mapped);
            continue;
        }

        // Handle decomposed rare voiced katakana like `ヸ`/`ヹ` (U+30F0/U+30F1 + U+3099).
        //
        // These correspond to the precomposed `ヸ`/`ヹ` characters, which Excel's `ASC` converts to
        // the halfwidth base + dakuten mark (`ｲﾞ`/`ｴﾞ`). If the input text is normalized into a
        // decomposed form, preserve Excel's behavior.
        if matches!(ch, 'ヰ' | 'ヱ') {
            if let Some(&next) = iter.peek() {
                if next == COMBINING_DAKUTEN {
                    out.push_str(match ch {
                        'ヰ' => "\u{FF72}", // ｲ
                        'ヱ' => "\u{FF74}", // ｴ
                        _ => unreachable!("matches! ensures only ヰ/ヱ"),
                    });
                    out.push(HALFWIDTH_DAKUTEN);
                    iter.next(); // consume combining mark
                    continue;
                }
            }
        }

        if let Some(mapped) = fullwidth_katakana_to_halfwidth(ch) {
            // Handle decomposed voiced/semi-voiced katakana like `ガ` (U+30AB + U+3099).
            // Excel's ASC emits the halfwidth base + dakuten/handakuten marks (`ｶﾞ`).
            if let Some(&next) = iter.peek() {
                if next == COMBINING_DAKUTEN {
                    out.push_str(mapped);
                    out.push(HALFWIDTH_DAKUTEN);
                    iter.next(); // consume combining mark
                    continue;
                }
                if next == COMBINING_HANDAKUTEN {
                    out.push_str(mapped);
                    out.push(HALFWIDTH_HANDAKUTEN);
                    iter.next(); // consume combining mark
                    continue;
                }
            }
            out.push_str(mapped);
            continue;
        }

        out.push(ch);
    }

    out
}

fn asc_dbcs_basic(input: &str) -> String {
    let mut out = String::with_capacity(input.len());

    for ch in input.chars() {
        if ch == FULLWIDTH_SPACE {
            out.push(' ');
            continue;
        }

        // Fullwidth ASCII variants: U+FF01..U+FF5E -> U+0021..U+007E.
        if ('\u{FF01}'..='\u{FF5E}').contains(&ch) {
            let ascii = char::from_u32((ch as u32).saturating_sub(0xFEE0))
                .expect("FF01..FF5E - 0xFEE0 must remain in Unicode scalar range");
            out.push(ascii);
            continue;
        }

        if let Some(mapped) = fullwidth_symbol_to_halfwidth(ch) {
            out.push(mapped);
            continue;
        }

        out.push(ch);
    }

    out
}

fn dbcs_cp932(input: &str) -> String {
    // Output can grow (e.g. ASCII -> fullwidth), but input length is a reasonable lower bound.
    let mut out = String::with_capacity(input.len());
    let mut iter = input.chars().peekable();

    while let Some(ch) = iter.next() {
        if ch == ' ' {
            out.push(FULLWIDTH_SPACE);
            continue;
        }

        // ASCII (printable): U+0021..U+007E -> U+FF01..U+FF5E.
        if ('!'..='~').contains(&ch) {
            let fw = char::from_u32(ch as u32 + 0xFEE0)
                .expect("ASCII 0x21..0x7E + 0xFEE0 must be valid Unicode scalar");
            out.push(fw);
            continue;
        }

        if let Some(mapped) = halfwidth_symbol_to_fullwidth(ch) {
            out.push(mapped);
            continue;
        }

        // Halfwidth katakana + punctuation live in U+FF61..U+FF9F (including dakuten marks).
        if ('\u{FF61}'..='\u{FF9F}').contains(&ch) {
            if let Some(&next) = iter.peek() {
                if next == HALFWIDTH_DAKUTEN || next == HALFWIDTH_HANDAKUTEN {
                    if let Some(composed) = compose_halfwidth_katakana(ch, next) {
                        out.push(composed);
                        iter.next(); // consume mark
                        continue;
                    }
                }
            }

            if let Some(mapped) = halfwidth_katakana_to_fullwidth(ch) {
                out.push(mapped);
            } else {
                // Should not happen, but keep behavior deterministic.
                out.push(ch);
            }
            continue;
        }

        out.push(ch);
    }

    out
}

fn dbcs_dbcs_basic(input: &str) -> String {
    // Output can grow (e.g. ASCII -> fullwidth), but input length is a reasonable lower bound.
    let mut out = String::with_capacity(input.len());

    for ch in input.chars() {
        if ch == ' ' {
            out.push(FULLWIDTH_SPACE);
            continue;
        }

        // ASCII -> fullwidth ASCII.
        if ('!'..='~').contains(&ch) {
            let full = char::from_u32((ch as u32).saturating_add(0xFEE0))
                .expect("ASCII + 0xFEE0 must remain in Unicode scalar range");
            out.push(full);
            continue;
        }

        if let Some(mapped) = halfwidth_symbol_to_fullwidth(ch) {
            out.push(mapped);
            continue;
        }

        out.push(ch);
    }

    out
}

fn fullwidth_symbol_to_halfwidth(ch: char) -> Option<char> {
    Some(match ch {
        FULLWIDTH_CENT => '¢',
        FULLWIDTH_POUND => '£',
        FULLWIDTH_NOT => '¬',
        FULLWIDTH_MACRON => '¯',
        FULLWIDTH_BROKEN_BAR => '¦',
        FULLWIDTH_YEN => '¥',
        FULLWIDTH_WON => '₩',
        _ => return None,
    })
}

fn halfwidth_symbol_to_fullwidth(ch: char) -> Option<char> {
    Some(match ch {
        '¢' => FULLWIDTH_CENT,
        '£' => FULLWIDTH_POUND,
        '¬' => FULLWIDTH_NOT,
        '¯' => FULLWIDTH_MACRON,
        '¦' => FULLWIDTH_BROKEN_BAR,
        '¥' => FULLWIDTH_YEN,
        '₩' => FULLWIDTH_WON,
        _ => return None,
    })
}

fn fullwidth_katakana_to_halfwidth(ch: char) -> Option<&'static str> {
    Some(match ch {
        // Punctuation.
        '。' => "\u{FF61}", // ｡
        '「' => "\u{FF62}", // ｢
        '」' => "\u{FF63}", // ｣
        '、' => "\u{FF64}", // ､
        '・' => "\u{FF65}", // ･

        // Base katakana + small kana.
        'ヲ' => "\u{FF66}", // ｦ
        'ァ' => "\u{FF67}", // ｧ
        'ィ' => "\u{FF68}", // ｨ
        'ゥ' => "\u{FF69}", // ｩ
        'ェ' => "\u{FF6A}", // ｪ
        'ォ' => "\u{FF6B}", // ｫ
        'ャ' => "\u{FF6C}", // ｬ
        'ュ' => "\u{FF6D}", // ｭ
        'ョ' => "\u{FF6E}", // ｮ
        'ッ' => "\u{FF6F}", // ｯ
        'ー' => "\u{FF70}", // ｰ
        'ア' => "\u{FF71}", // ｱ
        'イ' => "\u{FF72}", // ｲ
        'ウ' => "\u{FF73}", // ｳ
        'エ' => "\u{FF74}", // ｴ
        'オ' => "\u{FF75}", // ｵ
        'カ' => "\u{FF76}", // ｶ
        'キ' => "\u{FF77}", // ｷ
        'ク' => "\u{FF78}", // ｸ
        'ケ' => "\u{FF79}", // ｹ
        'コ' => "\u{FF7A}", // ｺ
        'サ' => "\u{FF7B}", // ｻ
        'シ' => "\u{FF7C}", // ｼ
        'ス' => "\u{FF7D}", // ｽ
        'セ' => "\u{FF7E}", // ｾ
        'ソ' => "\u{FF7F}", // ｿ
        'タ' => "\u{FF80}", // ﾀ
        'チ' => "\u{FF81}", // ﾁ
        'ツ' => "\u{FF82}", // ﾂ
        'テ' => "\u{FF83}", // ﾃ
        'ト' => "\u{FF84}", // ﾄ
        'ナ' => "\u{FF85}", // ﾅ
        'ニ' => "\u{FF86}", // ﾆ
        'ヌ' => "\u{FF87}", // ﾇ
        'ネ' => "\u{FF88}", // ﾈ
        'ノ' => "\u{FF89}", // ﾉ
        'ハ' => "\u{FF8A}", // ﾊ
        'ヒ' => "\u{FF8B}", // ﾋ
        'フ' => "\u{FF8C}", // ﾌ
        'ヘ' => "\u{FF8D}", // ﾍ
        'ホ' => "\u{FF8E}", // ﾎ
        'マ' => "\u{FF8F}", // ﾏ
        'ミ' => "\u{FF90}", // ﾐ
        'ム' => "\u{FF91}", // ﾑ
        'メ' => "\u{FF92}", // ﾒ
        'モ' => "\u{FF93}", // ﾓ
        'ヤ' => "\u{FF94}", // ﾔ
        'ユ' => "\u{FF95}", // ﾕ
        'ヨ' => "\u{FF96}", // ﾖ
        'ラ' => "\u{FF97}", // ﾗ
        'リ' => "\u{FF98}", // ﾘ
        'ル' => "\u{FF99}", // ﾙ
        'レ' => "\u{FF9A}", // ﾚ
        'ロ' => "\u{FF9B}", // ﾛ
        'ワ' => "\u{FF9C}", // ﾜ
        'ン' => "\u{FF9D}", // ﾝ

        // Dakuten / handakuten composed forms.
        'ガ' => "\u{FF76}\u{FF9E}",
        'ギ' => "\u{FF77}\u{FF9E}",
        'グ' => "\u{FF78}\u{FF9E}",
        'ゲ' => "\u{FF79}\u{FF9E}",
        'ゴ' => "\u{FF7A}\u{FF9E}",
        'ザ' => "\u{FF7B}\u{FF9E}",
        'ジ' => "\u{FF7C}\u{FF9E}",
        'ズ' => "\u{FF7D}\u{FF9E}",
        'ゼ' => "\u{FF7E}\u{FF9E}",
        'ゾ' => "\u{FF7F}\u{FF9E}",
        'ダ' => "\u{FF80}\u{FF9E}",
        'ヂ' => "\u{FF81}\u{FF9E}",
        'ヅ' => "\u{FF82}\u{FF9E}",
        'デ' => "\u{FF83}\u{FF9E}",
        'ド' => "\u{FF84}\u{FF9E}",
        'バ' => "\u{FF8A}\u{FF9E}",
        'ビ' => "\u{FF8B}\u{FF9E}",
        'ブ' => "\u{FF8C}\u{FF9E}",
        'ベ' => "\u{FF8D}\u{FF9E}",
        'ボ' => "\u{FF8E}\u{FF9E}",
        'パ' => "\u{FF8A}\u{FF9F}",
        'ピ' => "\u{FF8B}\u{FF9F}",
        'プ' => "\u{FF8C}\u{FF9F}",
        'ペ' => "\u{FF8D}\u{FF9F}",
        'ポ' => "\u{FF8E}\u{FF9F}",
        'ヴ' => "\u{FF73}\u{FF9E}",

        // Less common voiced katakana supported by Unicode.
        'ヷ' => "\u{FF9C}\u{FF9E}", // ﾜﾞ
        'ヸ' => "\u{FF72}\u{FF9E}", // ｲﾞ
        'ヹ' => "\u{FF74}\u{FF9E}", // ｴﾞ
        'ヺ' => "\u{FF66}\u{FF9E}", // ｦﾞ

        // Spacing marks.
        '゛' => "\u{FF9E}",
        '゜' => "\u{FF9F}",

        _ => return None,
    })
}

fn encoding_for_codepage(codepage: u16) -> Option<&'static Encoding> {
    Some(match codepage as u32 {
        874 => WINDOWS_874,
        932 => SHIFT_JIS,
        936 => GBK,
        949 => EUC_KR,
        950 => BIG5,
        1250 => WINDOWS_1250,
        1251 => WINDOWS_1251,
        1252 => WINDOWS_1252,
        1253 => WINDOWS_1253,
        1254 => WINDOWS_1254,
        1255 => WINDOWS_1255,
        1256 => WINDOWS_1256,
        1257 => WINDOWS_1257,
        1258 => WINDOWS_1258,
        65001 => UTF_8,
        _ => return None,
    })
}

fn halfwidth_katakana_to_fullwidth(ch: char) -> Option<char> {
    Some(match ch {
        '\u{FF61}' => '。', // ｡
        '\u{FF62}' => '「', // ｢
        '\u{FF63}' => '」', // ｣
        '\u{FF64}' => '、', // ､
        '\u{FF65}' => '・', // ･
        '\u{FF66}' => 'ヲ', // ｦ
        '\u{FF67}' => 'ァ', // ｧ
        '\u{FF68}' => 'ィ', // ｨ
        '\u{FF69}' => 'ゥ', // ｩ
        '\u{FF6A}' => 'ェ', // ｪ
        '\u{FF6B}' => 'ォ', // ｫ
        '\u{FF6C}' => 'ャ', // ｬ
        '\u{FF6D}' => 'ュ', // ｭ
        '\u{FF6E}' => 'ョ', // ｮ
        '\u{FF6F}' => 'ッ', // ｯ
        '\u{FF70}' => 'ー', // ｰ
        '\u{FF71}' => 'ア', // ｱ
        '\u{FF72}' => 'イ', // ｲ
        '\u{FF73}' => 'ウ', // ｳ
        '\u{FF74}' => 'エ', // ｴ
        '\u{FF75}' => 'オ', // ｵ
        '\u{FF76}' => 'カ', // ｶ
        '\u{FF77}' => 'キ', // ｷ
        '\u{FF78}' => 'ク', // ｸ
        '\u{FF79}' => 'ケ', // ｹ
        '\u{FF7A}' => 'コ', // ｺ
        '\u{FF7B}' => 'サ', // ｻ
        '\u{FF7C}' => 'シ', // ｼ
        '\u{FF7D}' => 'ス', // ｽ
        '\u{FF7E}' => 'セ', // ｾ
        '\u{FF7F}' => 'ソ', // ｿ
        '\u{FF80}' => 'タ', // ﾀ
        '\u{FF81}' => 'チ', // ﾁ
        '\u{FF82}' => 'ツ', // ﾂ
        '\u{FF83}' => 'テ', // ﾃ
        '\u{FF84}' => 'ト', // ﾄ
        '\u{FF85}' => 'ナ', // ﾅ
        '\u{FF86}' => 'ニ', // ﾆ
        '\u{FF87}' => 'ヌ', // ﾇ
        '\u{FF88}' => 'ネ', // ﾈ
        '\u{FF89}' => 'ノ', // ﾉ
        '\u{FF8A}' => 'ハ', // ﾊ
        '\u{FF8B}' => 'ヒ', // ﾋ
        '\u{FF8C}' => 'フ', // ﾌ
        '\u{FF8D}' => 'ヘ', // ﾍ
        '\u{FF8E}' => 'ホ', // ﾎ
        '\u{FF8F}' => 'マ', // ﾏ
        '\u{FF90}' => 'ミ', // ﾐ
        '\u{FF91}' => 'ム', // ﾑ
        '\u{FF92}' => 'メ', // ﾒ
        '\u{FF93}' => 'モ', // ﾓ
        '\u{FF94}' => 'ヤ', // ﾔ
        '\u{FF95}' => 'ユ', // ﾕ
        '\u{FF96}' => 'ヨ', // ﾖ
        '\u{FF97}' => 'ラ', // ﾗ
        '\u{FF98}' => 'リ', // ﾘ
        '\u{FF99}' => 'ル', // ﾙ
        '\u{FF9A}' => 'レ', // ﾚ
        '\u{FF9B}' => 'ロ', // ﾛ
        '\u{FF9C}' => 'ワ', // ﾜ
        '\u{FF9D}' => 'ン', // ﾝ

        // Marks (when not used for composition).
        HALFWIDTH_DAKUTEN => '゛',
        HALFWIDTH_HANDAKUTEN => '゜',

        _ => return None,
    })
}

fn compose_halfwidth_katakana(base: char, mark: char) -> Option<char> {
    Some(match (base, mark) {
        // Dakuten.
        ('\u{FF76}', HALFWIDTH_DAKUTEN) => 'ガ', // ｶﾞ
        ('\u{FF77}', HALFWIDTH_DAKUTEN) => 'ギ',
        ('\u{FF78}', HALFWIDTH_DAKUTEN) => 'グ',
        ('\u{FF79}', HALFWIDTH_DAKUTEN) => 'ゲ',
        ('\u{FF7A}', HALFWIDTH_DAKUTEN) => 'ゴ',
        ('\u{FF7B}', HALFWIDTH_DAKUTEN) => 'ザ',
        ('\u{FF7C}', HALFWIDTH_DAKUTEN) => 'ジ',
        ('\u{FF7D}', HALFWIDTH_DAKUTEN) => 'ズ',
        ('\u{FF7E}', HALFWIDTH_DAKUTEN) => 'ゼ',
        ('\u{FF7F}', HALFWIDTH_DAKUTEN) => 'ゾ',
        ('\u{FF80}', HALFWIDTH_DAKUTEN) => 'ダ',
        ('\u{FF81}', HALFWIDTH_DAKUTEN) => 'ヂ',
        ('\u{FF82}', HALFWIDTH_DAKUTEN) => 'ヅ',
        ('\u{FF83}', HALFWIDTH_DAKUTEN) => 'デ',
        ('\u{FF84}', HALFWIDTH_DAKUTEN) => 'ド',
        ('\u{FF8A}', HALFWIDTH_DAKUTEN) => 'バ',
        ('\u{FF8B}', HALFWIDTH_DAKUTEN) => 'ビ',
        ('\u{FF8C}', HALFWIDTH_DAKUTEN) => 'ブ',
        ('\u{FF8D}', HALFWIDTH_DAKUTEN) => 'ベ',
        ('\u{FF8E}', HALFWIDTH_DAKUTEN) => 'ボ',
        ('\u{FF73}', HALFWIDTH_DAKUTEN) => 'ヴ', // ｳﾞ
        ('\u{FF9C}', HALFWIDTH_DAKUTEN) => 'ヷ', // ﾜﾞ
        ('\u{FF72}', HALFWIDTH_DAKUTEN) => 'ヸ', // ｲﾞ
        ('\u{FF74}', HALFWIDTH_DAKUTEN) => 'ヹ', // ｴﾞ
        ('\u{FF66}', HALFWIDTH_DAKUTEN) => 'ヺ', // ｦﾞ

        // Handakuten.
        ('\u{FF8A}', HALFWIDTH_HANDAKUTEN) => 'パ',
        ('\u{FF8B}', HALFWIDTH_HANDAKUTEN) => 'ピ',
        ('\u{FF8C}', HALFWIDTH_HANDAKUTEN) => 'プ',
        ('\u{FF8D}', HALFWIDTH_HANDAKUTEN) => 'ペ',
        ('\u{FF8E}', HALFWIDTH_HANDAKUTEN) => 'ポ',

        _ => return None,
    })
}

fn encode_bytes_len(codepage: u16, text: &str) -> usize {
    // Excel semantics: `*B` byte-count functions only differ from their non-`B` equivalents in
    // DBCS locales. For single-byte codepages, byte count matches character count.
    //
    // Note: even in single-byte locales, strings may contain characters that are not representable
    // in the legacy codepage. Excel still treats these as single-byte for `LENB` in non-DBCS
    // environments, so we use `chars().count()` rather than attempting to encode.
    match codepage as u32 {
        932 | 936 | 949 | 950 => {}
        _ => return text.chars().count(),
    }

    let Some(encoding) = encoding_for_codepage(codepage) else {
        // Best-effort fallback: treat byte count as character count.
        return text.chars().count();
    };
    // `encoding_rs::Encoding::encode` emits HTML numeric character references for unmappable
    // characters, which is Web-correct but not what we want for Excel byte-count semantics. We
    // instead count the number of bytes that would be produced by the encoding while treating any
    // unmappable code points as a single replacement byte (matching the behavior of common Windows
    // codepages).
    let mut encoder = encoding.new_encoder();
    let mut remaining = text;
    let mut total = 0usize;
    let mut scratch: Vec<u8> = vec![0u8; 64];

    while !remaining.is_empty() {
        let (result, read, written) =
            encoder.encode_from_utf8_without_replacement(remaining, &mut scratch, true);
        total = total.saturating_add(written);
        remaining = remaining.get(read..).unwrap_or("");

        match result {
            encoding_rs::EncoderResult::InputEmpty => {}
            encoding_rs::EncoderResult::OutputFull => {
                // Ensure progress even in the unlikely event that the scratch buffer is too small
                // to encode a single code point.
                if read == 0 && written == 0 {
                    let new_len = scratch.len().saturating_mul(2).max(1);
                    scratch.resize(new_len, 0);
                }
            }
            encoding_rs::EncoderResult::Unmappable(_) => {
                // Treat unmappable code points as a single replacement byte (e.g. '?').
                total = total.saturating_add(1);
                if read == 0 {
                    // Defensive: ensure forward progress if the encoder reports an unmappable
                    // without consuming input.
                    if let Some(ch) = remaining.chars().next() {
                        remaining = &remaining[ch.len_utf8()..];
                    } else {
                        remaining = "";
                    }
                }
            }
        }
    }

    // Flush any pending encoder state (not expected for the encodings we use, but keep the length
    // calculation conservative).
    loop {
        let (result, _read, written) =
            encoder.encode_from_utf8_without_replacement("", &mut scratch, true);
        total = total.saturating_add(written);
        match result {
            encoding_rs::EncoderResult::InputEmpty => break,
            encoding_rs::EncoderResult::OutputFull => {
                let new_len = scratch.len().saturating_mul(2).max(1);
                scratch.resize(new_len, 0);
            }
            encoding_rs::EncoderResult::Unmappable(_) => {
                total = total.saturating_add(1);
            }
        }
    }

    total
}

fn encoded_byte_prefixes(codepage: u16, text: &str) -> Vec<usize> {
    let Some(encoding) = encoding_for_codepage(codepage) else {
        // Fallback: treat each character as a single byte.
        let mut out = Vec::with_capacity(text.chars().count().saturating_add(1));
        out.push(0);
        for (idx, _ch) in text.chars().enumerate() {
            out.push(idx.saturating_add(1));
        }
        return out;
    };

    let mut out = Vec::with_capacity(text.chars().count().saturating_add(1));
    out.push(0);

    let mut total = 0usize;
    for ch in text.chars() {
        total = total.saturating_add(encoded_byte_len_for_char(encoding, ch));
        out.push(total);
    }
    out
}

fn encoded_byte_len_for_char(encoding: &'static Encoding, ch: char) -> usize {
    // DBCS encodings used by Excel are stateless; encode a single codepoint and count bytes.
    let mut encoder = encoding.new_encoder();
    let mut utf8_buf = [0u8; 4];
    let input = ch.encode_utf8(&mut utf8_buf);
    let mut scratch = [0u8; 8];
    loop {
        let (result, _read, written) =
            encoder.encode_from_utf8_without_replacement(input, &mut scratch, true);
        match result {
            encoding_rs::EncoderResult::InputEmpty => return written,
            encoding_rs::EncoderResult::Unmappable(_) => return 1,
            encoding_rs::EncoderResult::OutputFull => {
                // Should not happen for DBCS codepages (<=2 bytes/codepoint), but be defensive.
                let mut bigger = vec![0u8; scratch.len().saturating_mul(2).max(16)];
                let (result, _read, written) =
                    encoder.encode_from_utf8_without_replacement(input, &mut bigger, true);
                match result {
                    encoding_rs::EncoderResult::InputEmpty => return written,
                    encoding_rs::EncoderResult::Unmappable(_) => return 1,
                    encoding_rs::EncoderResult::OutputFull => {
                        // Give up; treat as single-byte.
                        return 1;
                    }
                }
            }
        }
    }
}

fn slice_bytes_dbcs(codepage: u16, text: &str, start0: usize, len: usize) -> String {
    if len == 0 {
        return String::new();
    }

    let prefixes = encoded_byte_prefixes(codepage, text);
    let total = prefixes.last().copied().unwrap_or(0);

    let start0 = start0.min(total);
    let end0 = start0.saturating_add(len).min(total);

    // Align start to the next character boundary (ceil) and end to the previous boundary (floor)
    // so we never return partial DBCS code units.
    let start_char = prefixes
        .iter()
        .position(|&b| b >= start0)
        .unwrap_or(prefixes.len().saturating_sub(1));
    let end_char_excl = prefixes.iter().rposition(|&b| b <= end0).unwrap_or(0);

    if end_char_excl <= start_char {
        return String::new();
    }

    text.chars()
        .skip(start_char)
        .take(end_char_excl - start_char)
        .collect()
}

fn replaceb_bytes(
    codepage: u16,
    old_text: &str,
    start0: usize,
    len: usize,
    new_text: &str,
) -> String {
    let prefixes = encoded_byte_prefixes(codepage, old_text);
    let total = prefixes.last().copied().unwrap_or(0);

    let start0 = start0.min(total);
    let end0 = start0.saturating_add(len).min(total);

    let start_char = prefixes
        .iter()
        .position(|&b| b >= start0)
        .unwrap_or(prefixes.len().saturating_sub(1));
    let mut end_char_excl = prefixes.iter().rposition(|&b| b <= end0).unwrap_or(0);
    if end_char_excl < start_char {
        end_char_excl = start_char;
    }

    let start_byte = char_pos_to_byte(old_text, start_char);
    let end_byte = char_pos_to_byte(old_text, end_char_excl);

    let mut out = String::with_capacity(
        old_text.len() - end_byte.saturating_sub(start_byte) + new_text.len(),
    );
    out.push_str(&old_text[..start_byte]);
    out.push_str(new_text);
    out.push_str(&old_text[end_byte..]);
    out
}

fn char_pos_to_byte(s: &str, char_pos: usize) -> usize {
    if char_pos == 0 {
        return 0;
    }
    s.char_indices()
        .nth(char_pos)
        .map(|(idx, _)| idx)
        .unwrap_or_else(|| s.len())
}

fn findb_impl(
    codepage: u16,
    needle: &str,
    haystack: &str,
    start: i64,
    case_insensitive: bool,
) -> Value {
    if start < 1 {
        return Value::Error(ErrorKind::Value);
    }

    let mut hay_chars: Vec<char> = Vec::new();
    let mut byte_prefixes: Vec<usize> = Vec::new();
    byte_prefixes.push(0);
    let mut total_bytes = 0usize;

    let encoding = encoding_for_codepage(codepage);
    for ch in haystack.chars() {
        hay_chars.push(ch);
        let len = match encoding {
            Some(enc) => encoded_byte_len_for_char(enc, ch),
            None => 1,
        };
        total_bytes = total_bytes.saturating_add(len);
        byte_prefixes.push(total_bytes);
    }

    let needle_chars: Vec<char> = needle.chars().collect();

    let Ok(start0) = usize::try_from(start.saturating_sub(1)) else {
        return Value::Error(ErrorKind::Value);
    };
    if start0 > total_bytes {
        return Value::Error(ErrorKind::Value);
    }

    // Convert byte-based start offset to a character index aligned to the next boundary.
    let start_idx = byte_prefixes
        .iter()
        .position(|&b| b >= start0)
        .unwrap_or(hay_chars.len());
    if start_idx > hay_chars.len() {
        return Value::Error(ErrorKind::Value);
    }

    if needle_chars.is_empty() {
        return Value::Number(start as f64);
    }

    if case_insensitive {
        // Excel SEARCH is case-insensitive using Unicode-aware uppercasing (e.g. ß -> SS).
        // Fold both pattern and haystack into a comparable char stream.
        let needle_folded: Vec<char> = if needle.is_ascii() {
            needle.chars().map(|c| c.to_ascii_uppercase()).collect()
        } else {
            needle.chars().flat_map(|c| c.to_uppercase()).collect()
        };
        let needle_tokens = parse_search_pattern(&needle_folded);

        let mut hay_folded = Vec::new();
        let mut folded_starts = Vec::with_capacity(hay_chars.len());
        for ch in &hay_chars {
            folded_starts.push(hay_folded.len());
            if ch.is_ascii() {
                hay_folded.push(ch.to_ascii_uppercase());
            } else {
                hay_folded.extend(ch.to_uppercase());
            }
        }

        for orig_idx in start_idx..hay_chars.len() {
            let folded_idx = folded_starts[orig_idx];
            if matches_pattern(&needle_tokens, &hay_folded, folded_idx) {
                let byte_pos = byte_prefixes
                    .get(orig_idx)
                    .copied()
                    .unwrap_or(0)
                    .saturating_add(1);
                return Value::Number(byte_pos as f64);
            }
        }
        Value::Error(ErrorKind::Value)
    } else {
        let needle_tokens = vec![PatternToken::LiteralSeq(needle_chars)];
        for i in start_idx..hay_chars.len() {
            if matches_pattern(&needle_tokens, &hay_chars, i) {
                let byte_pos = byte_prefixes.get(i).copied().unwrap_or(0).saturating_add(1);
                return Value::Number(byte_pos as f64);
            }
        }
        Value::Error(ErrorKind::Value)
    }
}

#[derive(Debug, Clone)]
enum PatternToken {
    LiteralSeq(Vec<char>),
    AnyOne,
    AnyMany,
}

fn parse_search_pattern(pattern: &[char]) -> Vec<PatternToken> {
    let mut tokens = Vec::new();
    let mut literal = Vec::new();
    let mut idx = 0;
    while idx < pattern.len() {
        let ch = pattern[idx];
        if ch == '~' {
            idx += 1;
            if idx < pattern.len() {
                literal.push(pattern[idx]);
                idx += 1;
            } else {
                literal.push('~');
            }
            continue;
        }
        match ch {
            '*' => {
                if !literal.is_empty() {
                    tokens.push(PatternToken::LiteralSeq(std::mem::take(&mut literal)));
                }
                tokens.push(PatternToken::AnyMany);
                idx += 1;
            }
            '?' => {
                if !literal.is_empty() {
                    tokens.push(PatternToken::LiteralSeq(std::mem::take(&mut literal)));
                }
                tokens.push(PatternToken::AnyOne);
                idx += 1;
            }
            _ => {
                literal.push(ch);
                idx += 1;
            }
        }
    }
    if !literal.is_empty() {
        tokens.push(PatternToken::LiteralSeq(literal));
    }
    tokens
}

fn matches_pattern(tokens: &[PatternToken], hay: &[char], start: usize) -> bool {
    let mut memo = vec![vec![None; hay.len() + 1]; tokens.len() + 1];
    match_rec(tokens, hay, start, 0, &mut memo)
}

fn match_rec(
    tokens: &[PatternToken],
    hay: &[char],
    hay_idx: usize,
    tok_idx: usize,
    memo: &mut [Vec<Option<bool>>],
) -> bool {
    if let Some(cached) = memo[tok_idx][hay_idx] {
        return cached;
    }
    let result = if tok_idx == tokens.len() {
        true
    } else {
        match &tokens[tok_idx] {
            PatternToken::LiteralSeq(seq) => {
                if hay_idx + seq.len() > hay.len() {
                    false
                } else if hay[hay_idx..hay_idx + seq.len()] == *seq {
                    match_rec(tokens, hay, hay_idx + seq.len(), tok_idx + 1, memo)
                } else {
                    false
                }
            }
            PatternToken::AnyOne => {
                if hay_idx >= hay.len() {
                    false
                } else {
                    match_rec(tokens, hay, hay_idx + 1, tok_idx + 1, memo)
                }
            }
            PatternToken::AnyMany => {
                if match_rec(tokens, hay, hay_idx, tok_idx + 1, memo) {
                    true
                } else if hay_idx < hay.len() {
                    match_rec(tokens, hay, hay_idx + 1, tok_idx, memo)
                } else {
                    false
                }
            }
        }
    };
    memo[tok_idx][hay_idx] = Some(result);
    result
}
