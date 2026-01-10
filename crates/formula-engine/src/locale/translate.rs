use crate::eval::FormulaParseError;
use crate::parser::{lex, Token, TokenKind};
use crate::LocaleConfig;

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
pub fn canonicalize_formula(formula: &str, locale: &FormulaLocale) -> Result<String, FormulaParseError> {
    translate_formula(formula, locale, Direction::ToCanonical)
}

/// Convert a canonical (English) formula into its locale-specific display form.
///
/// The input may include an optional leading `=`, which is preserved in the output.
pub fn localize_formula(formula: &str, locale: &FormulaLocale) -> Result<String, FormulaParseError> {
    translate_formula(formula, locale, Direction::ToLocalized)
}

#[derive(Debug, Clone, Copy)]
enum Direction {
    ToCanonical,
    ToLocalized,
}

fn translate_formula(
    formula: &str,
    locale: &FormulaLocale,
    dir: Direction,
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

    let tokens = lex(expr_src, src_config).map_err(map_lex_error)?;

    let mut out = String::with_capacity(trimmed.len());
    if has_equals {
        out.push('=');
    }

    for (idx, tok) in tokens.iter().enumerate() {
        match &tok.kind {
            TokenKind::Eof => break,
            TokenKind::Whitespace(raw) | TokenKind::Intersect(raw) => out.push_str(raw),
            TokenKind::String(_) | TokenKind::QuotedIdent(_) => {
                out.push_str(token_slice(expr_src, tok)?);
            }
            TokenKind::Number(raw) => {
                out.push_str(&translate_number(raw, src_config.decimal_separator, dst_config.decimal_separator));
            }
            TokenKind::Ident(raw) if is_function_ident(&tokens, idx) => match dir {
                Direction::ToCanonical => out.push_str(&locale.canonical_function_name(raw)),
                Direction::ToLocalized => out.push_str(&locale.localized_function_name(raw)),
            },
            TokenKind::ArgSep | TokenKind::Union => {
                out.push(dst_config.arg_separator);
            }
            TokenKind::ArrayRowSep => {
                out.push(dst_config.array_row_separator);
            }
            TokenKind::ArrayColSep => {
                out.push(dst_config.array_col_separator);
            }
            _ => {
                out.push_str(token_slice(expr_src, tok)?);
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
    while matches!(tokens.get(j).map(|t| &t.kind), Some(TokenKind::Whitespace(_))) {
        j += 1;
    }

    matches!(tokens.get(j).map(|t| &t.kind), Some(TokenKind::LParen))
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

