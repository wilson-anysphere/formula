use crate::eval::FormulaParseError;
use crate::parser::{lex, Token, TokenKind};
use crate::value::casefold;
use crate::{ErrorKind, LocaleConfig, ParseOptions, ReferenceStyle};

use super::FormulaLocale;

/// Convert a locale-specific formula into the canonical form we persist/evaluate.
///
/// Canonical form uses:
/// - English function names (e.g. `SUM`)
/// - `,` as list/argument separator (and union operator)
/// - `.` as decimal separator
/// - en-US array separators (`,` columns, `;` rows)
///
/// The input may include an optional leading `=`, which is preserved in the output.
pub fn canonicalize_formula(
    formula: &str,
    locale: &FormulaLocale,
) -> Result<String, FormulaParseError> {
    translate_formula_with_style(formula, locale, Direction::ToCanonical, ReferenceStyle::A1)
}

/// Convert a locale-specific formula into the canonical form we persist/evaluate, using the
/// provided reference style for tokenization (A1 vs R1C1).
///
/// This is useful for UI workflows that allow users to edit formulas in R1C1 mode while still
/// supporting localized function names and separators.
pub fn canonicalize_formula_with_style(
    formula: &str,
    locale: &FormulaLocale,
    reference_style: ReferenceStyle,
) -> Result<String, FormulaParseError> {
    translate_formula_with_style(formula, locale, Direction::ToCanonical, reference_style)
}

/// Convert a canonical (English) formula into its locale-specific display form.
///
/// The input may include an optional leading `=`, which is preserved in the output.
pub fn localize_formula(
    formula: &str,
    locale: &FormulaLocale,
) -> Result<String, FormulaParseError> {
    translate_formula_with_style(formula, locale, Direction::ToLocalized, ReferenceStyle::A1)
}

/// Convert a canonical (English) formula into its locale-specific display form, using the provided
/// reference style for tokenization (A1 vs R1C1).
pub fn localize_formula_with_style(
    formula: &str,
    locale: &FormulaLocale,
    reference_style: ReferenceStyle,
) -> Result<String, FormulaParseError> {
    translate_formula_with_style(formula, locale, Direction::ToLocalized, reference_style)
}

#[derive(Debug, Clone, Copy)]
enum Direction {
    ToCanonical,
    ToLocalized,
}

fn bool_literal(value: bool) -> &'static str {
    if value {
        "TRUE"
    } else {
        "FALSE"
    }
}

fn translate_formula_with_style(
    formula: &str,
    locale: &FormulaLocale,
    dir: Direction,
    reference_style: ReferenceStyle,
) -> Result<String, FormulaParseError> {
    // Match the previous implementation: accept leading whitespace and keep an optional leading `=`.
    let trimmed = formula.trim_start();
    let (has_equals, expr_src) = if let Some(rest) = trimmed.strip_prefix('=') {
        (true, rest)
    } else {
        (false, trimmed)
    };

    let canonical_config = LocaleConfig::en_us();
    let (src_config, dst_config) = match dir {
        Direction::ToCanonical => (&locale.config, &canonical_config),
        Direction::ToLocalized => (&canonical_config, &locale.config),
    };

    let parse_opts = ParseOptions {
        locale: src_config.clone(),
        reference_style,
        normalize_relative_to: None,
    };
    let tokens = lex(expr_src, &parse_opts).map_err(map_lex_error)?;

    let mut out = String::with_capacity(trimmed.len());
    if has_equals {
        out.push('=');
    }

    let mut bracket_depth: usize = 0;
    let mut idx = 0usize;
    while idx < tokens.len() {
        // Special-case localized function names that contain dots (e.g. `CONTAR.SI(...)` in es-ES).
        //
        // The lexer tokenizes `.` as [`TokenKind::Dot`] (used for field access and other syntax),
        // so without this step a localized function like `CONTAR.SI(` would appear as the token
        // sequence `Ident("CONTAR")`, `Dot`, `Ident("SI")`, `LParen`, and the function-name
        // translation logic would incorrectly treat `SI` as a field-access selector.
        if bracket_depth == 0 && matches!(dir, Direction::ToCanonical) {
            if let Some((translated, next_idx)) =
                try_translate_dotted_function_call(&tokens, idx, locale)
            {
                out.push_str(&translated);
                idx = next_idx;
                continue;
            }
        }

        let tok = &tokens[idx];
        match &tok.kind {
            TokenKind::Eof => break,
            TokenKind::LBracket => {
                bracket_depth += 1;
                out.push_str(token_slice(expr_src, tok)?);
                idx += 1;
            }
            TokenKind::RBracket => {
                // Excel escapes `]` inside structured references as `]]`. At the outermost bracket
                // depth, treat double `]]` as a literal `]` and keep the bracket depth unchanged.
                if bracket_depth == 1
                    && matches!(
                        tokens.get(idx + 1).map(|t| &t.kind),
                        Some(TokenKind::RBracket)
                    )
                {
                    out.push_str(token_slice(expr_src, tok)?);
                    if let Some(next) = tokens.get(idx + 1) {
                        out.push_str(token_slice(expr_src, next)?);
                    }
                    idx += 2;
                    continue;
                }

                bracket_depth = bracket_depth.saturating_sub(1);
                out.push_str(token_slice(expr_src, tok)?);
                idx += 1;
            }
            _ if bracket_depth > 0 => {
                // Do not translate anything inside `[...]` bracket groups.
                //
                // This includes:
                // - external workbook/sheet prefixes like `[Book.xlsx]Sheet1!A1`
                // - structured references like `Table1[[#Headers],[Qty]]`
                // - field access selectors like `A1.["Field Name"]`
                //
                // For structured references specifically, Excel keeps both the inner separators and
                // the reserved item keywords (`[#Headers]`, `[#Data]`, `[#Totals]`, `[#All]`,
                // `[#This Row]`) canonical across locales, so we intentionally avoid rewriting
                // anything inside the bracketed segments.
                //
                // To verify the `Formula` ↔ `FormulaLocal` behavior against a real Excel install,
                // see `tools/excel-oracle/extract-structured-reference-keywords.ps1`.
                out.push_str(token_slice(expr_src, tok)?);
                idx += 1;
            }
            TokenKind::Whitespace(raw) | TokenKind::Intersect(raw) => {
                out.push_str(raw);
                idx += 1;
            }
            TokenKind::String(_) | TokenKind::QuotedIdent(_) => {
                out.push_str(token_slice(expr_src, tok)?);
                idx += 1;
            }
            TokenKind::Boolean(value) => {
                // Preserve boolean keywords used as field-access selectors (e.g. `A1.TRUE`) by
                // skipping localization/canonicalization when preceded by `.`.
                if matches!(prev_non_trivia_kind(&tokens, idx), Some(TokenKind::Dot)) {
                    out.push_str(token_slice(expr_src, tok)?);
                } else {
                    match dir {
                        Direction::ToCanonical => out.push_str(bool_literal(*value)),
                        Direction::ToLocalized => {
                            out.push_str(locale.localized_boolean_literal(*value))
                        }
                    }
                }
                idx += 1;
            }
            TokenKind::Error(raw) => {
                match dir {
                    Direction::ToCanonical => {
                        // 1) Apply locale-specific mapping (localized -> canonical) if present.
                        // 2) Normalize canonical spelling/casing for known errors
                        //    (e.g. `#N/A!` -> `#N/A`, `#value!` -> `#VALUE!`).
                        // 3) Otherwise preserve the original token text.
                        let canonical = locale.canonical_error_literal(raw).unwrap_or(raw.as_str());
                        if let Some(kind) = ErrorKind::from_code(canonical) {
                            out.push_str(kind.as_code());
                        } else {
                            out.push_str(canonical);
                        }
                    }
                    Direction::ToLocalized => {
                        // Some legacy/corrupt canonical formulas may contain non-canonical spellings
                        // like `#N/A!` or mixed casing. Normalize first so locale lookup works.
                        let canonical = ErrorKind::from_code(raw)
                            .map(|kind| kind.as_code())
                            .unwrap_or(raw.as_str());
                        if let Some(loc) = locale.localized_error_literal(canonical) {
                            out.push_str(loc);
                        } else {
                            out.push_str(canonical);
                        };
                    }
                }
                idx += 1;
            }
            TokenKind::Number(raw) => {
                match dir {
                    Direction::ToCanonical => out.push_str(&translate_number(
                        raw,
                        src_config.decimal_separator,
                        dst_config.decimal_separator,
                    )),
                    Direction::ToLocalized => out.push_str(&localize_number(
                        raw,
                        src_config.decimal_separator,
                        dst_config,
                    )),
                }
                idx += 1;
            }
            TokenKind::Ident(raw)
                if is_function_ident(&tokens, idx) && !is_field_access_selector(&tokens, idx) =>
            {
                match dir {
                    Direction::ToCanonical => out.push_str(&locale.canonical_function_name(raw)),
                    Direction::ToLocalized => out.push_str(&locale.localized_function_name(raw)),
                }
                idx += 1;
            }
            TokenKind::Ident(raw) if matches!(dir, Direction::ToCanonical) => {
                // Boolean keywords are locale-specific (e.g. `WAHR`/`FALSCH` for German), but those
                // tokens can also appear as identifiers (e.g. table names, sheet prefixes). Only
                // translate them when they are used as standalone scalar literals.
                if !is_sheet_prefix_ident(&tokens, idx)
                    && !is_table_name_ident(&tokens, idx)
                    && !is_field_access_selector(&tokens, idx)
                {
                    if let Some(value) = locale.canonical_boolean_literal(raw) {
                        out.push_str(bool_literal(value));
                    } else {
                        out.push_str(token_slice(expr_src, tok)?);
                    }
                } else {
                    out.push_str(token_slice(expr_src, tok)?);
                }
                idx += 1;
            }
            TokenKind::ArgSep | TokenKind::Union => {
                out.push(dst_config.arg_separator);
                idx += 1;
            }
            TokenKind::ArrayRowSep => {
                out.push(dst_config.array_row_separator);
                idx += 1;
            }
            TokenKind::ArrayColSep => {
                out.push(dst_config.array_col_separator);
                idx += 1;
            }
            _ => {
                out.push_str(token_slice(expr_src, tok)?);
                idx += 1;
            }
        }
    }

    Ok(out)
}

fn is_function_ident(tokens: &[Token], idx: usize) -> bool {
    if !matches!(tokens.get(idx).map(|t| &t.kind), Some(TokenKind::Ident(_))) {
        return false;
    }

    let mut j = idx + 1;
    while matches!(
        tokens.get(j).map(|t| &t.kind),
        Some(TokenKind::Whitespace(_))
    ) {
        j += 1;
    }

    matches!(tokens.get(j).map(|t| &t.kind), Some(TokenKind::LParen))
}

fn try_translate_dotted_function_call(
    tokens: &[Token],
    idx: usize,
    locale: &FormulaLocale,
) -> Option<(String, usize)> {
    let TokenKind::Ident(first) = tokens.get(idx).map(|t| &t.kind)? else {
        return None;
    };

    // A dotted function name must start at the beginning of the identifier sequence (not as a
    // field-access selector like `A1.TRUE`).
    if matches!(prev_non_trivia_kind(tokens, idx), Some(TokenKind::Dot)) {
        return None;
    }

    let mut j = idx;
    let mut combined = first.clone();
    while matches!(tokens.get(j + 1).map(|t| &t.kind), Some(TokenKind::Dot)) {
        let Some(TokenKind::Ident(next)) = tokens.get(j + 2).map(|t| &t.kind) else {
            break;
        };
        combined.push('.');
        combined.push_str(next);
        j += 2;
    }

    // No dots => not a dotted function name.
    if j == idx {
        return None;
    }

    // Must be a function call (allow whitespace before `(`).
    let mut k = j + 1;
    while matches!(
        tokens.get(k).map(|t| &t.kind),
        Some(TokenKind::Whitespace(_))
    ) {
        k += 1;
    }
    if !matches!(tokens.get(k).map(|t| &t.kind), Some(TokenKind::LParen)) {
        return None;
    }

    let translated = locale.canonical_function_name(&combined);

    // Only treat this as a function name when it is an explicit locale translation. Otherwise,
    // preserve the original token stream so expressions like `SomeRecord.Field(...)` are not
    // rewritten.
    if translated == casefold_function_name_for_compare(&combined) {
        return None;
    }

    // Advance to the token immediately after the last identifier; any whitespace before the `(`
    // should be handled by the main loop (to preserve formatting).
    Some((translated, j + 1))
}

fn is_field_access_selector(tokens: &[Token], idx: usize) -> bool {
    if !matches!(tokens.get(idx).map(|t| &t.kind), Some(TokenKind::Ident(_))) {
        return false;
    }

    let mut j = idx;
    while j > 0 {
        j -= 1;
        match tokens.get(j).map(|t| &t.kind) {
            Some(TokenKind::Whitespace(_)) => continue,
            Some(TokenKind::Dot) => return true,
            _ => return false,
        }
    }

    false
}

fn casefold_function_name_for_compare(name: &str) -> String {
    let (has_prefix, base) = split_xlfn_prefix(name);
    let mut out = String::new();
    if has_prefix {
        out.push_str("_xlfn.");
    }
    out.push_str(&casefold(base));
    out
}

fn split_xlfn_prefix(name: &str) -> (bool, &str) {
    const PREFIX: &str = "_xlfn.";
    let Some(prefix) = name.get(..PREFIX.len()) else {
        return (false, name);
    };
    if prefix.eq_ignore_ascii_case(PREFIX) {
        (true, &name[PREFIX.len()..])
    } else {
        (false, name)
    }
}

fn next_non_trivia_kind<'a>(tokens: &'a [Token], idx: usize) -> Option<&'a TokenKind> {
    let mut j = idx + 1;
    while matches!(
        tokens.get(j).map(|t| &t.kind),
        Some(TokenKind::Whitespace(_))
    ) {
        j += 1;
    }
    tokens.get(j).map(|t| &t.kind)
}

fn prev_non_trivia_kind<'a>(tokens: &'a [Token], idx: usize) -> Option<&'a TokenKind> {
    if idx == 0 {
        return None;
    }

    let mut j = idx;
    while j > 0 {
        j -= 1;
        match &tokens[j].kind {
            TokenKind::Whitespace(_) => continue,
            other => return Some(other),
        }
    }
    None
}

fn is_sheet_prefix_ident(tokens: &[Token], idx: usize) -> bool {
    match next_non_trivia_kind(tokens, idx) {
        Some(TokenKind::Bang) => true,
        // 3D sheet span: `Sheet1:Sheet3!A1`
        Some(TokenKind::Colon) => {
            // idx -> start sheet ident
            // colon -> end sheet name
            // bang -> reference separator
            let mut colon_idx = idx + 1;
            while matches!(
                tokens.get(colon_idx).map(|t| &t.kind),
                Some(TokenKind::Whitespace(_))
            ) {
                colon_idx += 1;
            }
            if !matches!(
                tokens.get(colon_idx).map(|t| &t.kind),
                Some(TokenKind::Colon)
            ) {
                return false;
            }

            let mut end_sheet_idx = colon_idx + 1;
            while matches!(
                tokens.get(end_sheet_idx).map(|t| &t.kind),
                Some(TokenKind::Whitespace(_))
            ) {
                end_sheet_idx += 1;
            }
            match tokens.get(end_sheet_idx).map(|t| &t.kind) {
                Some(TokenKind::Ident(_)) | Some(TokenKind::QuotedIdent(_)) => {}
                _ => return false,
            }

            matches!(
                next_non_trivia_kind(tokens, end_sheet_idx),
                Some(TokenKind::Bang)
            )
        }
        _ => false,
    }
}

fn is_table_name_ident(tokens: &[Token], idx: usize) -> bool {
    matches!(next_non_trivia_kind(tokens, idx), Some(TokenKind::LBracket))
}

fn translate_number(raw: &str, decimal_in: char, decimal_out: char) -> String {
    if decimal_in == decimal_out {
        return raw.to_string();
    }
    raw.chars()
        .map(|ch| if ch == decimal_in { decimal_out } else { ch })
        .collect()
}

fn localize_number(raw: &str, decimal_in: char, dst: &LocaleConfig) -> String {
    let translated = translate_number(raw, decimal_in, dst.decimal_separator);
    let Some(group_sep) = dst.thousands_separator else {
        return translated;
    };

    // Avoid inserting grouping separators that would make the formula ambiguous or conflict with
    // formula syntax.
    //
    // Note: en-US sets `thousands_separator: None` because `,` conflicts with the arg separator.
    // This check is defensive for any future locale configs.
    if group_sep == dst.decimal_separator || group_sep == dst.arg_separator || group_sep == ' ' {
        return translated;
    }

    // Split mantissa/exponent (`E` is always uppercase after lexing).
    let (mantissa, exponent) = if let Some(idx) = translated.find('E') {
        (&translated[..idx], &translated[idx..])
    } else {
        (translated.as_str(), "")
    };

    // Split integer/fractional portions.
    let (int_part, frac_part) = if let Some(idx) = mantissa.find(dst.decimal_separator) {
        (&mantissa[..idx], &mantissa[idx..])
    } else {
        (mantissa, "")
    };

    // Avoid allocating if there's nothing to do.
    if int_part.len() <= 3 || !int_part.chars().all(|c| c.is_ascii_digit()) {
        return translated;
    }

    let grouped_int = insert_thousands_separators(int_part, group_sep);
    let mut out = String::with_capacity(grouped_int.len() + frac_part.len() + exponent.len());
    out.push_str(&grouped_int);
    out.push_str(frac_part);
    out.push_str(exponent);
    out
}

fn insert_thousands_separators(int_part: &str, sep: char) -> String {
    debug_assert!(int_part.chars().all(|c| c.is_ascii_digit()));

    let len = int_part.len();
    if len <= 3 {
        return int_part.to_string();
    }

    // Group from the left: the first group is 1-3 digits, followed by groups of 3 digits.
    let first = match len % 3 {
        0 => 3,
        n => n,
    };

    let sep_count = (len - 1) / 3;
    let mut out = String::with_capacity(len + sep_count);
    out.push_str(&int_part[..first]);
    let mut idx = first;
    while idx < len {
        out.push(sep);
        out.push_str(&int_part[idx..idx + 3]);
        idx += 3;
    }
    out
}

fn token_slice<'a>(src: &'a str, tok: &Token) -> Result<&'a str, FormulaParseError> {
    src.get(tok.span.start..tok.span.end)
        .ok_or_else(|| FormulaParseError::UnexpectedToken("invalid token span".to_string()))
}

fn map_lex_error(err: crate::ParseError) -> FormulaParseError {
    // The legacy translation API uses `FormulaParseError`; keep the mapping coarse.
    if err.message.to_ascii_lowercase().contains("unterminated") {
        FormulaParseError::UnexpectedEof
    } else {
        FormulaParseError::UnexpectedToken(err.message)
    }
}
