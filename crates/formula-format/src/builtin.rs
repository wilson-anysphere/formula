use std::borrow::Cow;

use crate::Locale;

/// Lookup table for Excel's built-in number format codes.
///
/// The OOXML spec defines format IDs 0-49 as built-ins. Many additional IDs are
/// reserved by Excel, but 0-49 cover the vast majority of files.
///
/// ## Locale-variant IDs
///
/// Excel's built-in table is locale-dependent. In practice, this most visibly
/// affects:
/// - currency/accounting formats (currency symbol + sometimes placement)
///   - ids **5–8** (Currency) and **42/44** (Accounting w/ currency)
/// - short date formats (month/day ordering differs by locale)
///   - ids **14/22** and reserved locale ids **27–31**
///
/// [`builtin_format_code`] returns the canonical *en-US* variants.
/// [`builtin_format_code_with_locale`] can be used when the caller knows the
/// workbook locale and wants a closer Excel match.
///
/// References:
/// - ECMA-376 Part 1, 18.8.30 `numFmts`
/// - Excel "Format Cells" built-in formats
const BUILTIN_FORMATS_EN_US: [&str; 50] = [
    // 0-4: General/number.
    "General",  // 0
    "0",        // 1
    "0.00",     // 2
    "#,##0",    // 3
    "#,##0.00", // 4
    // 5-8: Currency.
    "$#,##0_);($#,##0)",             // 5
    "$#,##0_);[Red]($#,##0)",        // 6
    "$#,##0.00_);($#,##0.00)",       // 7
    "$#,##0.00_);[Red]($#,##0.00)",  // 8
    // 9-11: Percent / scientific.
    "0%",        // 9
    "0.00%",     // 10
    "0.00E+00",  // 11
    // 12-13: Fractions.
    "# ?/?",     // 12
    "# ??/??",   // 13
    // 14-22: Date/time.
    "m/d/yyyy",        // 14
    "d-mmm-yy",        // 15
    "d-mmm",           // 16
    "mmm-yy",          // 17
    "h:mm AM/PM",      // 18
    "h:mm:ss AM/PM",   // 19
    "h:mm",            // 20
    "h:mm:ss",         // 21
    "m/d/yyyy h:mm",   // 22
    // 23-26: Accounting-style negatives (no currency symbol).
    "#,##0_);(#,##0)",            // 23
    "#,##0_);[Red](#,##0)",       // 24
    "#,##0.00_);(#,##0.00)",      // 25
    "#,##0.00_);[Red](#,##0.00)", // 26
    // 27-36: Locale-reserved date/time ids (en-US duplicates).
    "m/d/yyyy",  // 27
    "m/d/yyyy",  // 28
    "m/d/yyyy",  // 29
    "m/d/yyyy",  // 30
    "m/d/yyyy",  // 31
    "h:mm:ss",   // 32
    "h:mm:ss",   // 33
    "h:mm:ss",   // 34
    "h:mm:ss",   // 35
    "h:mm:ss",   // 36
    // 37-40: Accounting-style negatives (duplicate of 23-26 in en-US).
    "#,##0_);(#,##0)",            // 37
    "#,##0_);[Red](#,##0)",       // 38
    "#,##0.00_);(#,##0.00)",      // 39
    "#,##0.00_);[Red](#,##0.00)", // 40
    // 41-44: Accounting formats (with alignment underscores/fill).
    r#"_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)"#,        // 41
    r#"_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)"#,     // 42
    r#"_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)"#, // 43
    r#"_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)"#, // 44
    // 45-47: Time/duration.
    "mm:ss",      // 45
    "[h]:mm:ss",  // 46
    "mm:ss.0",    // 47
    // 48-49: Scientific + text.
    "##0.0E+0",   // 48
    "@",          // 49
];

pub fn builtin_format_code(id: u16) -> Option<&'static str> {
    BUILTIN_FORMATS_EN_US.get(id as usize).copied()
}

/// Reverse lookup: return the built-in format ID for an exact (canonical) format
/// code.
///
/// This uses the *en-US* built-in table and returns the first matching ID. Some
/// format codes are duplicated across multiple built-in IDs (e.g. 23 and 37),
/// so callers should treat the returned ID as a canonical representative rather
/// than a guaranteed round-trip value.
pub fn builtin_format_id(code: &str) -> Option<u16> {
    BUILTIN_FORMATS_EN_US
        .iter()
        .position(|c| *c == code)
        .and_then(|idx| u16::try_from(idx).ok())
}

/// Locale-aware resolver for Excel's built-in number format codes.
///
/// This returns the best-effort format code Excel would use for the given
/// locale.
///
/// Notes:
/// - The format engine treats `,` (grouping) and `.` (decimal) as invariant
///   tokens in the format code; the caller's [`crate::FormatOptions::locale`]
///   controls how those separators are rendered. As a result, this function
///   does **not** rewrite numeric separators inside the returned format code.
/// - Date separators in patterns that use `/` and time separators in patterns
///   that use `:` are rendered according to [`Locale`]. For common locale
///   differences, this function primarily adjusts month/day ordering.
pub fn builtin_format_code_with_locale(id: u16, locale: Locale) -> Option<Cow<'static, str>> {
    let base = builtin_format_code(id)?;

    // Fast path for the canonical en-US table.
    if locale == Locale::en_us() {
        return Some(Cow::Borrowed(base));
    }

    // Built-in short date formats are locale-variant in Excel: most locales are
    // day-first, while en-US is month-first.
    let day_first_dates = locale != Locale::en_us();

    match id {
        // Short date formats: flip m/d ordering for day-first locales.
        14 | 27 | 28 | 29 | 30 | 31 if day_first_dates => Some(Cow::Borrowed("dd/mm/yyyy")),
        22 if day_first_dates => Some(Cow::Borrowed("dd/mm/yyyy h:mm")),
        // Currency/accounting formats: substitute currency symbol. For now we
        // only special-case the locales represented by [`Locale`] constructors.
        5..=8 | 42 | 44 => {
            let currency = match locale {
                l if l == Locale::de_de()
                    || l == Locale::fr_fr()
                    || l == Locale::it_it()
                    || l == Locale::es_es() =>
                {
                    "€"
                }
                _ => "$",
            };

            if currency == "$" {
                Some(Cow::Borrowed(base))
            } else {
                Some(Cow::Owned(base.replace('$', currency)))
            }
        }
        _ => Some(Cow::Borrowed(base)),
    }
}
