//! Legacy DBCS / byte-count text functions.
//!
//! Excel exposes `*B` variants of several text functions (LENB, LEFTB, MIDB, RIGHTB,
//! FINDB, SEARCHB, REPLACEB). In DBCS locales (e.g. Japanese), these functions
//! operate on *byte counts* instead of character counts, and the definition of a
//! "byte" depends on the active workbook locale / code page.
//!
//! The formula engine currently assumes an en-US workbook locale and Unicode
//! strings. Under that single-byte locale, the `*B` functions behave identically
//! to their non-`B` equivalents.
//!
//! `ASC` / `DBCS` perform half-width / full-width conversions in Japanese locales.
//! We implement these conversions only when the active workbook text codepage is
//! 932 (Shift_JIS / Japanese). In other locales/codepages, they behave as
//! identity transforms.
//!
//! `PHONETIC` depends on per-cell phonetic guide metadata (furigana).
//! When phonetic metadata is present for a referenced cell, `PHONETIC(reference)`
//! returns that stored string. When phonetic metadata is absent (the common
//! case), Excel falls back to the referenced cell’s displayed text, so the
//! engine returns the referenced value coerced to text using the current
//! locale-aware formatting rules.
//!
//! Once workbook locale + codepage + phonetic metadata are modeled, this module
//! can be extended to implement real Excel semantics for DBCS workbooks.

use crate::eval::CompiledExpr;
use crate::eval::MAX_MATERIALIZED_ARRAY_CELLS;
use crate::functions::array_lift;
use crate::functions::{call_function, ArgValue, FunctionContext, Reference};
use crate::value::{Array, ErrorKind, Value};

pub(crate) fn findb_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    // en-US: byte counts match character counts.
    call_function(ctx, "FIND", args)
}

pub(crate) fn searchb_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    // en-US: byte counts match character counts.
    call_function(ctx, "SEARCH", args)
}

pub(crate) fn replaceb_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    // en-US: byte counts match character counts.
    call_function(ctx, "REPLACE", args)
}

pub(crate) fn leftb_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    // en-US: byte counts match character counts.
    call_function(ctx, "LEFT", args)
}

pub(crate) fn rightb_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    // en-US: byte counts match character counts.
    call_function(ctx, "RIGHT", args)
}

pub(crate) fn midb_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    // en-US: byte counts match character counts.
    call_function(ctx, "MID", args)
}

pub(crate) fn lenb_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    // en-US: byte counts match character counts.
    call_function(ctx, "LEN", args)
}

pub(crate) fn asc_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let text = array_lift::eval_arg(ctx, &args[0]);
    let cp932 = ctx.text_codepage() == 932;
    array_lift::lift1(text, |text| {
        let s = text.coerce_to_string_with_ctx(ctx)?;
        if !cp932 {
            return Ok(Value::Text(s));
        }
        Ok(Value::Text(asc_cp932(&s)))
    })
}

pub(crate) fn dbcs_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let text = array_lift::eval_arg(ctx, &args[0]);
    let cp932 = ctx.text_codepage() == 932;
    array_lift::lift1(text, |text| {
        let s = text.coerce_to_string_with_ctx(ctx)?;
        if !cp932 {
            return Ok(Value::Text(s));
        }
        Ok(Value::Text(dbcs_cp932(&s)))
    })
}

pub(crate) fn phonetic_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    match ctx.eval_arg(&args[0]) {
        ArgValue::Reference(reference) => phonetic_from_reference(ctx, reference),
        // TODO: Verify Excel's behavior for scalar/non-reference arguments (e.g. `PHONETIC("abc")`).
        // Historically, the engine treated PHONETIC as a string-coercion placeholder; preserve that
        // behavior until we have an Excel oracle case for scalar arguments.
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

fn asc_cp932(input: &str) -> String {
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

        if let Some(mapped) = fullwidth_katakana_to_halfwidth(ch) {
            out.push_str(mapped);
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
        'ヺ' => "\u{FF66}\u{FF9E}", // ｦﾞ

        // Spacing marks.
        '゛' => "\u{FF9E}",
        '゜' => "\u{FF9F}",

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
