use crate::eval::FormulaParseError;
use crate::parser::{lex, Token, TokenKind};
use crate::{LocaleConfig, ParseOptions, ReferenceStyle};

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
                // Do not translate anything inside workbook/structured reference brackets.
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
                match dir {
                    Direction::ToCanonical => out.push_str(bool_literal(*value)),
                    Direction::ToLocalized => {
                        out.push_str(locale.localized_boolean_literal(*value))
                    }
                }
                idx += 1;
            }
            TokenKind::Error(raw) => {
                match dir {
                    Direction::ToCanonical => {
                        if let Some(canon) = locale.canonical_error_literal(raw) {
                            out.push_str(canon);
                        } else {
                            out.push_str(token_slice(expr_src, tok)?);
                        }
                    }
                    Direction::ToLocalized => {
                        if let Some(loc) = locale.localized_error_literal(raw) {
                            out.push_str(loc);
                        } else {
                            out.push_str(token_slice(expr_src, tok)?);
                        }
                    }
                }
                idx += 1;
            }
            TokenKind::Number(raw) => {
                out.push_str(&translate_number(
                    raw,
                    src_config.decimal_separator,
                    dst_config.decimal_separator,
                ));
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
            if !matches!(tokens.get(colon_idx).map(|t| &t.kind), Some(TokenKind::Colon)) {
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

            matches!(next_non_trivia_kind(tokens, end_sheet_idx), Some(TokenKind::Bang))
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
