//! BIFF function table (Ftab) for Excel built-in functions.
//!
//! Excel encodes built-in functions in tokenized BIFF formulas (Rgce) using a 16-bit
//! function identifier (`iftab`) in `PtgFunc` / `PtgFuncVar` tokens. That identifier
//! indexes into a fixed function table commonly referred to as **Ftab**.
//!
//! ## Provenance
//!
//! The `FTAB` table below (ids `0..=484`) is derived from the Microsoft Office binary
//! file format documentation (BIFF Function Codes) and cross-checked against the
//! open-source `calamine` crate (v0.32.0, `src/utils.rs`).
//!
//! Newer Excel functions that appear in formulas with an `_xlfn.` prefix (for forward
//! compatibility) are often encoded in BIFF as **user-defined functions** using the
//! special `iftab` value 255 plus a separate name token. For those,
//! [`function_id_from_name`] returns 255 when the name is `_xlfn.`-prefixed (even if
//! it is not present in `FTAB`), and for a curated allowlist of unprefixed names used
//! by `formula-engine`.

use std::collections::HashMap;
use std::sync::OnceLock;

/// BIFF `iftab` value used for user-defined / add-in / future functions.
pub const FTAB_USER_DEFINED: u16 = 255;

#[cfg(feature = "encode")]
fn is_formula_engine_function(name: &str) -> bool {
    formula_engine::functions::lookup_function(name).is_some()
}

#[cfg(not(feature = "encode"))]
fn is_formula_engine_function(_name: &str) -> bool {
    false
}

/// Function name table indexed by BIFF `iftab`. Empty strings denote reserved ids.
///
/// NOTE: The table currently matches the BIFF12 function table as published in the
/// Microsoft binary format docs up through Excel 2007-era additions.
pub const FTAB: [&str; 485] = [
    "COUNT",
    "IF",
    "ISNA",
    "ISERROR",
    "SUM",
    "AVERAGE",
    "MIN",
    "MAX",
    "ROW",
    "COLUMN",
    "NA",
    "NPV",
    "STDEV",
    "DOLLAR",
    "FIXED",
    "SIN",
    "COS",
    "TAN",
    "ATAN",
    "PI",
    "SQRT",
    "EXP",
    "LN",
    "LOG10",
    "ABS",
    "INT",
    "SIGN",
    "ROUND",
    "LOOKUP",
    "INDEX",
    "REPT",
    "MID",
    "LEN",
    "VALUE",
    "TRUE",
    "FALSE",
    "AND",
    "OR",
    "NOT",
    "MOD",
    "DCOUNT",
    "DSUM",
    "DAVERAGE",
    "DMIN",
    "DMAX",
    "DSTDEV",
    "VAR",
    "DVAR",
    "TEXT",
    "LINEST",
    "TREND",
    "LOGEST",
    "GROWTH",
    "GOTO",
    "HALT",
    "RETURN",
    "PV",
    "FV",
    "NPER",
    "PMT",
    "RATE",
    "MIRR",
    "IRR",
    "RAND",
    "MATCH",
    "DATE",
    "TIME",
    "DAY",
    "MONTH",
    "YEAR",
    "WEEKDAY",
    "HOUR",
    "MINUTE",
    "SECOND",
    "NOW",
    "AREAS",
    "ROWS",
    "COLUMNS",
    "OFFSET",
    "ABSREF",
    "RELREF",
    "ARGUMENT",
    "SEARCH",
    "TRANSPOSE",
    "ERROR",
    "STEP",
    "TYPE",
    "ECHO",
    "SET.NAME",
    "CALLER",
    "DEREF",
    "WINDOWS",
    "SERIES",
    "DOCUMENTS",
    "ACTIVE.CELL",
    "SELECTION",
    "RESULT",
    "ATAN2",
    "ASIN",
    "ACOS",
    "CHOOSE",
    "HLOOKUP",
    "VLOOKUP",
    "LINKS",
    "INPUT",
    "ISREF",
    "GET.FORMULA",
    "GET.NAME",
    "SET.VALUE",
    "LOG",
    "EXEC",
    "CHAR",
    "LOWER",
    "UPPER",
    "PROPER",
    "LEFT",
    "RIGHT",
    "EXACT",
    "TRIM",
    "REPLACE",
    "SUBSTITUTE",
    "CODE",
    "NAMES",
    "DIRECTORY",
    "FIND",
    "CELL",
    "ISERR",
    "ISTEXT",
    "ISNUMBER",
    "ISBLANK",
    "T",
    "N",
    "FOPEN",
    "FCLOSE",
    "FSIZE",
    "FREADLN",
    "FREAD",
    "FWRITELN",
    "FWRITE",
    "FPOS",
    "DATEVALUE",
    "TIMEVALUE",
    "SLN",
    "SYD",
    "DDB",
    "GET.DEF",
    "REFTEXT",
    "TEXTREF",
    "INDIRECT",
    "REGISTER",
    "CALL",
    "ADD.BAR",
    "ADD.MENU",
    "ADD.COMMAND",
    "ENABLE.COMMAND",
    "CHECK.COMMAND",
    "RENAME.COMMAND",
    "SHOW.BAR",
    "DELETE.MENU",
    "DELETE.COMMAND",
    "GET.CHART.ITEM",
    "DIALOG.BOX",
    "CLEAN",
    "MDETERM",
    "MINVERSE",
    "MMULT",
    "FILES",
    "IPMT",
    "PPMT",
    "COUNTA",
    "CANCEL.KEY",
    "FOR",
    "WHILE",
    "BREAK",
    "NEXT",
    "INITIATE",
    "REQUEST",
    "POKE",
    "EXECUTE",
    "TERMINATE",
    "RESTART",
    "HELP",
    "GET.BAR",
    "PRODUCT",
    "FACT",
    "GET.CELL",
    "GET.WORKSPACE",
    "GET.WINDOW",
    "GET.DOCUMENT",
    "DPRODUCT",
    "ISNONTEXT",
    "GET.NOTE",
    "NOTE",
    "STDEVP",
    "VARP",
    "DSTDEVP",
    "DVARP",
    "TRUNC",
    "ISLOGICAL",
    "DCOUNTA",
    "DELETE.BAR",
    "UNREGISTER",
    "",
    "",
    "USDOLLAR",
    "FINDB",
    "SEARCHB",
    "REPLACEB",
    "LEFTB",
    "RIGHTB",
    "MIDB",
    "LENB",
    "ROUNDUP",
    "ROUNDDOWN",
    "ASC",
    "DBCS",
    "RANK",
    "",
    "",
    "ADDRESS",
    "DAYS360",
    "TODAY",
    "VDB",
    "ELSE",
    "ELSE.IF",
    "END.IF",
    "FOR.CELL",
    "MEDIAN",
    "SUMPRODUCT",
    "SINH",
    "COSH",
    "TANH",
    "ASINH",
    "ACOSH",
    "ATANH",
    "DGET",
    "CREATE.OBJECT",
    "VOLATILE",
    "LAST.ERROR",
    "CUSTOM.UNDO",
    "CUSTOM.REPEAT",
    "FORMULA.CONVERT",
    "GET.LINK.INFO",
    "TEXT.BOX",
    "INFO",
    "GROUP",
    "GET.OBJECT",
    "DB",
    "PAUSE",
    "",
    "",
    "RESUME",
    "FREQUENCY",
    "ADD.TOOLBAR",
    "DELETE.TOOLBAR",
    "USER",
    "RESET.TOOLBAR",
    "EVALUATE",
    "GET.TOOLBAR",
    "GET.TOOL",
    "SPELLING.CHECK",
    "ERROR.TYPE",
    "APP.TITLE",
    "WINDOW.TITLE",
    "SAVE.TOOLBAR",
    "ENABLE.TOOL",
    "PRESS.TOOL",
    "REGISTER.ID",
    "GET.WORKBOOK",
    "AVEDEV",
    "BETADIST",
    "GAMMALN",
    "BETAINV",
    "BINOMDIST",
    "CHIDIST",
    "CHIINV",
    "COMBIN",
    "CONFIDENCE",
    "CRITBINOM",
    "EVEN",
    "EXPONDIST",
    "FDIST",
    "FINV",
    "FISHER",
    "FISHERINV",
    "FLOOR",
    "GAMMADIST",
    "GAMMAINV",
    "CEILING",
    "HYPGEOMDIST",
    "LOGNORMDIST",
    "LOGINV",
    "NEGBINOMDIST",
    "NORMDIST",
    "NORMSDIST",
    "NORMINV",
    "NORMSINV",
    "STANDARDIZE",
    "ODD",
    "PERMUT",
    "POISSON",
    "TDIST",
    "WEIBULL",
    "SUMXMY2",
    "SUMX2MY2",
    "SUMX2PY2",
    "CHITEST",
    "CORREL",
    "COVAR",
    "FORECAST",
    "FTEST",
    "INTERCEPT",
    "PEARSON",
    "RSQ",
    "STEYX",
    "SLOPE",
    "TTEST",
    "PROB",
    "DEVSQ",
    "GEOMEAN",
    "HARMEAN",
    "SUMSQ",
    "KURT",
    "SKEW",
    "ZTEST",
    "LARGE",
    "SMALL",
    "QUARTILE",
    "PERCENTILE",
    "PERCENTRANK",
    "MODE",
    "TRIMMEAN",
    "TINV",
    "",
    "MOVIE.COMMAND",
    "GET.MOVIE",
    "CONCATENATE",
    "POWER",
    "PIVOT.ADD.DATA",
    "GET.PIVOT.TABLE",
    "GET.PIVOT.FIELD",
    "GET.PIVOT.ITEM",
    "RADIANS",
    "DEGREES",
    "SUBTOTAL",
    "SUMIF",
    "COUNTIF",
    "COUNTBLANK",
    "SCENARIO.GET",
    "OPTIONS.LISTS.GET",
    "ISPMT",
    "DATEDIF",
    "DATESTRING",
    "NUMBERSTRING",
    "ROMAN",
    "OPEN.DIALOG",
    "SAVE.DIALOG",
    "VIEW.GET",
    "GETPIVOTDATA",
    "HYPERLINK",
    "PHONETIC",
    "AVERAGEA",
    "MAXA",
    "MINA",
    "STDEVPA",
    "VARPA",
    "STDEVA",
    "VARA",
    "BAHTTEXT",
    "THAIDAYOFWEEK",
    "THAIDIGIT",
    "THAIMONTHOFYEAR",
    "THAINUMSOUND",
    "THAINUMSTRING",
    "THAISTRINGLENGTH",
    "ISTHAIDIGIT",
    "ROUNDBAHTDOWN",
    "ROUNDBAHTUP",
    "THAIYEAR",
    "RTD",
    "CUBEVALUE",
    "CUBEMEMBER",
    "CUBEMEMBERPROPERTY",
    "CUBERANKEDMEMBER",
    "HEX2BIN",
    "HEX2DEC",
    "HEX2OCT",
    "DEC2BIN",
    "DEC2HEX",
    "DEC2OCT",
    "OCT2BIN",
    "OCT2HEX",
    "OCT2DEC",
    "BIN2DEC",
    "BIN2OCT",
    "BIN2HEX",
    "IMSUB",
    "IMDIV",
    "IMPOWER",
    "IMABS",
    "IMSQRT",
    "IMLN",
    "IMLOG2",
    "IMLOG10",
    "IMSIN",
    "IMCOS",
    "IMEXP",
    "IMARGUMENT",
    "IMCONJUGATE",
    "IMAGINARY",
    "IMREAL",
    "COMPLEX",
    "IMSUM",
    "IMPRODUCT",
    "SERIESSUM",
    "FACTDOUBLE",
    "SQRTPI",
    "QUOTIENT",
    "DELTA",
    "GESTEP",
    "ISEVEN",
    "ISODD",
    "MROUND",
    "ERF",
    "ERFC",
    "BESSELJ",
    "BESSELK",
    "BESSELY",
    "BESSELI",
    "XIRR",
    "XNPV",
    "PRICEMAT",
    "YIELDMAT",
    "INTRATE",
    "RECEIVED",
    "DISC",
    "PRICEDISC",
    "YIELDDISC",
    "TBILLEQ",
    "TBILLPRICE",
    "TBILLYIELD",
    "PRICE",
    "YIELD",
    "DOLLARDE",
    "DOLLARFR",
    "NOMINAL",
    "EFFECT",
    "CUMPRINC",
    "CUMIPMT",
    "EDATE",
    "EOMONTH",
    "YEARFRAC",
    "COUPDAYBS",
    "COUPDAYS",
    "COUPDAYSNC",
    "COUPNCD",
    "COUPNUM",
    "COUPPCD",
    "DURATION",
    "MDURATION",
    "ODDLPRICE",
    "ODDLYIELD",
    "ODDFPRICE",
    "ODDFYIELD",
    "RANDBETWEEN",
    "WEEKNUM",
    "AMORDEGRC",
    "AMORLINC",
    "CONVERT",
    "ACCRINT",
    "ACCRINTM",
    "WORKDAY",
    "NETWORKDAYS",
    "GCD",
    "MULTINOMIAL",
    "LCM",
    "FVSCHEDULE",
    "CUBEKPIMEMBER",
    "CUBESET",
    "CUBESETCOUNT",
    "IFERROR",
    "COUNTIFS",
    "SUMIFS",
    "AVERAGEIF",
    "AVERAGEIFS",
];

static NAME_TO_ID: OnceLock<HashMap<&'static str, u16>> = OnceLock::new();

fn name_to_id() -> &'static HashMap<&'static str, u16> {
    NAME_TO_ID.get_or_init(|| {
        let mut map = HashMap::new();
        let _ = map.try_reserve(FTAB.len());
        for (id, name) in FTAB.iter().enumerate() {
            if name.is_empty() {
                continue;
            }
            let Ok(id_u16) = u16::try_from(id) else {
                continue;
            };
            let prev = map.insert(*name, id_u16);
            debug_assert!(prev.is_none(), "duplicate FTAB entry for {name}");
        }
        map
    })
}

/// Return the canonical Excel function name for a BIFF `iftab` id.
pub fn function_name_from_id(id: u16) -> Option<&'static str> {
    FTAB.get(id as usize)
        .copied()
        .filter(|name| !name.is_empty())
}

/// Return the BIFF `iftab` id for a given function name.
///
/// Lookup rules:
/// - Case-insensitive (ASCII)
/// - Accepts the `_xlfn.` prefix used in files for forward-compatible functions, as well as
///   the `_xlws.` / `_xludf.` namespaces that appear under `_xlfn.` in some workbook files.
/// - Returns [`FTAB_USER_DEFINED`] (255) for unknown `_xlfn.` names, as well as
///   known future-function names not present in `FTAB`.
pub fn function_id_from_name(name: &str) -> Option<u16> {
    let name = name.trim();
    if name.is_empty() {
        return None;
    }

    // Fast path: avoid heap allocation for typical (short) function names by uppercasing into a
    // stack buffer. Uppercasing ASCII bytes preserves UTF-8 validity.
    let mut buf = [0u8; 64];
    if name.len() <= buf.len() {
        for (dst, src) in buf[..name.len()].iter_mut().zip(name.as_bytes()) {
            *dst = src.to_ascii_uppercase();
        }
        let upper = match std::str::from_utf8(&buf[..name.len()]) {
            Ok(s) => s,
            Err(_) => {
                // `name` is valid UTF-8 and ASCII uppercasing preserves all non-ASCII bytes.
                debug_assert!(false, "ASCII uppercasing should preserve UTF-8 for {name:?}");
                return None;
            }
        };
        return function_id_from_uppercase_name(upper);
    }

    let upper = name.to_ascii_uppercase();
    function_id_from_uppercase_name(&upper)
}

/// Return the BIFF `iftab` id for an already-uppercase function name.
///
/// This avoids allocating a temporary `String` when the caller already has an uppercase name
/// (e.g. during formula compilation).
///
/// Contract:
/// - `name` should be ASCII-uppercase (and typically trimmed).
/// - Lookup semantics match [`function_id_from_name`].
pub fn function_id_from_uppercase_name(name: &str) -> Option<u16> {
    let name = name.trim();
    if name.is_empty() {
        return None;
    }

    let (had_xlfn_prefix, normalized) = match name.strip_prefix("_XLFN.") {
        Some(stripped) => (true, stripped),
        None => (false, name),
    };

    // Namespace-qualified functions (`_xlws.*`, `_xludf.*`) are encoded as future/UDF calls in
    // BIFF (iftab=255). Excel commonly stores these as `_xlfn._xlws.*` / `_xlfn._xludf.*`, but
    // accept the unwrapped form for round-trip safety.
    if normalized.starts_with("_XLWS.") || normalized.starts_with("_XLUDF.") {
        return Some(FTAB_USER_DEFINED);
    }

    if let Some(id) = name_to_id().get(normalized).copied() {
        return Some(id);
    }

    if had_xlfn_prefix
        || FUTURE_UDF_FUNCTIONS.contains(&normalized)
        || is_formula_engine_function(normalized)
    {
        return Some(FTAB_USER_DEFINED);
    }

    None
}

// Functions not present in the published FTAB but implemented by formula-engine.
//
// These are typically stored by Excel as `_xlfn.` functions and encoded in BIFF as
// user-defined function calls (`iftab = 255`) with an accompanying name token.
//
// Keep this list sorted (ASCII) for maintainability.
const FUTURE_UDF_FUNCTIONS: &[&str] = &[
    "ACOT",
    "ACOTH",
    "AGGREGATE",
    "ARABIC",
    "BASE",
    "BETA.DIST",
    "BETA.INV",
    "BINOM.DIST",
    "BINOM.DIST.RANGE",
    "BINOM.INV",
    "BITAND",
    "BITLSHIFT",
    "BITOR",
    "BITRSHIFT",
    "BITXOR",
    "BYCOL",
    "BYROW",
    "CEILING.MATH",
    "CEILING.PRECISE",
    "CHISQ.DIST",
    "CHISQ.DIST.RT",
    "CHISQ.INV",
    "CHISQ.INV.RT",
    "CHISQ.TEST",
    "CHOOSECOLS",
    "CHOOSEROWS",
    "COMBINA",
    "CONCAT",
    "CONFIDENCE.NORM",
    "CONFIDENCE.T",
    "COT",
    "COTH",
    "COVARIANCE.P",
    "COVARIANCE.S",
    "CSC",
    "CSCH",
    "DAYS",
    "DECIMAL",
    "DROP",
    "EXPAND",
    "EXPON.DIST",
    "F.DIST",
    "F.DIST.RT",
    "F.INV",
    "F.INV.RT",
    "F.TEST",
    "FILTER",
    "FLOOR.MATH",
    "FLOOR.PRECISE",
    "FORECAST.ETS",
    "FORECAST.ETS.CONFINT",
    "FORECAST.ETS.SEASONALITY",
    "FORECAST.ETS.STAT",
    "FORECAST.LINEAR",
    "FORMULATEXT",
    "GAMMA",
    "GAMMA.DIST",
    "GAMMA.INV",
    "GAMMALN.PRECISE",
    "GAUSS",
    "HSTACK",
    "HYPGEOM.DIST",
    "IFNA",
    "IFS",
    "IMAGE",
    "ISFORMULA",
    "ISO.CEILING",
    "ISO.WEEKNUM",
    "ISOMITTED",
    "ISOWEEKNUM",
    "LAMBDA",
    "LET",
    "LOGNORM.DIST",
    "LOGNORM.INV",
    "MAKEARRAY",
    "MAP",
    "MAXIFS",
    "MINIFS",
    "MODE.MULT",
    "MODE.SNGL",
    "MUNIT",
    "NEGBINOM.DIST",
    "NETWORKDAYS.INTL",
    "NORM.DIST",
    "NORM.INV",
    "NORM.S.DIST",
    "NORM.S.INV",
    "NUMBERVALUE",
    "PDURATION",
    "PERCENTILE.EXC",
    "PERCENTILE.INC",
    "PERCENTRANK.EXC",
    "PERCENTRANK.INC",
    "PERMUTATIONA",
    "PHI",
    "POISSON.DIST",
    "QUARTILE.EXC",
    "QUARTILE.INC",
    "RANDARRAY",
    "RANK.AVG",
    "RANK.EQ",
    "REDUCE",
    "RRI",
    "SCAN",
    "SEC",
    "SECH",
    "SEQUENCE",
    "SHEET",
    "SHEETS",
    "SKEW.P",
    "SORT",
    "SORTBY",
    "STDEV.P",
    "STDEV.S",
    "SWITCH",
    "T.DIST",
    "T.DIST.2T",
    "T.DIST.RT",
    "T.INV",
    "T.INV.2T",
    "T.TEST",
    "TAKE",
    "TEXTAFTER",
    "TEXTBEFORE",
    "TEXTJOIN",
    "TEXTSPLIT",
    "TOCOL",
    "TOROW",
    "UNICHAR",
    "UNICODE",
    "UNIQUE",
    "VALUETOTEXT",
    "VAR.P",
    "VAR.S",
    "VSTACK",
    "WEIBULL.DIST",
    "WORKDAY.INTL",
    "WRAPCOLS",
    "WRAPROWS",
    "XLOOKUP",
    "XMATCH",
    "XOR",
    "Z.TEST",
];

#[cfg(test)]
mod tests {
    use std::collections::HashSet;

    use super::{function_id_from_name, FTAB, FTAB_USER_DEFINED, FUTURE_UDF_FUNCTIONS};

    fn extract_const_str_list(src: &str, const_name: &str) -> Vec<String> {
        let marker = format!("const {const_name}: &[&str] = &[");
        let start = src
            .find(&marker)
            .unwrap_or_else(|| panic!("could not find start of `{const_name}` list"));
        let rest = &src[start + marker.len()..];
        let end = rest
            .find("];")
            .unwrap_or_else(|| panic!("could not find end of `{const_name}` list"));
        let body = &rest[..end];

        body.lines()
            .filter_map(|line| {
                let line = line.trim();
                let line = line.strip_prefix('"')?;
                let end = line.find('"')?;
                Some(line[..end].to_string())
            })
            .collect()
    }

    #[test]
    fn future_udf_functions_are_sorted_and_unique() {
        let mut prev: Option<&str> = None;
        let mut seen = HashSet::new();
        let _ = seen.try_reserve(FUTURE_UDF_FUNCTIONS.len());
        for &name in FUTURE_UDF_FUNCTIONS {
            if let Some(prev) = prev {
                assert!(
                    prev < name,
                    "FUTURE_UDF_FUNCTIONS must be ASCII-sorted; found out-of-order entries: {prev} then {name}"
                );
            }
            assert!(
                seen.insert(name),
                "FUTURE_UDF_FUNCTIONS must not contain duplicates; duplicate entry: {name}"
            );
            prev = Some(name);
        }
    }

    #[test]
    fn future_udf_functions_do_not_overlap_ftab() {
        let ftab_names: HashSet<&str> = FTAB.iter().copied().filter(|name| !name.is_empty()).collect();
        for &name in FUTURE_UDF_FUNCTIONS {
            assert!(
                !ftab_names.contains(name),
                "FUTURE_UDF_FUNCTIONS should not contain FTAB entries; {name} is already in FTAB"
            );
        }
    }

    #[test]
    fn future_udf_functions_map_to_user_defined_id() {
        for &name in FUTURE_UDF_FUNCTIONS {
            assert_eq!(
                function_id_from_name(name),
                Some(FTAB_USER_DEFINED),
                "{name} should map to the BIFF UDF sentinel id"
            );

            let prefixed = format!("_xlfn.{name}");
            assert_eq!(
                function_id_from_name(&prefixed),
                Some(FTAB_USER_DEFINED),
                "{prefixed} should map to the BIFF UDF sentinel id"
            );
        }
    }

    #[test]
    fn future_udf_functions_cover_formula_engine_non_ftab_functions() {
        // This list exists so `function_id_from_name(\"NAME\")` can still return the BIFF UDF sentinel
        // (255) in no-`encode` builds where we cannot consult `formula-engine` at runtime.
        let ftab_names: HashSet<&str> = FTAB.iter().copied().filter(|name| !name.is_empty()).collect();
        let future_names: HashSet<&str> = FUTURE_UDF_FUNCTIONS.iter().copied().collect();

        for spec in formula_engine::functions::iter_function_specs() {
            let upper = spec.name.to_ascii_uppercase();
            let name = upper.strip_prefix("_XLFN.").unwrap_or(&upper);
            if ftab_names.contains(name) {
                continue;
            }
            assert!(
                future_names.contains(name),
                "missing FUTURE_UDF_FUNCTIONS entry for formula-engine function {name}"
            );
        }
    }

    #[test]
    fn function_id_from_uppercase_name_matches_standard_lookup() {
        assert_eq!(
            super::function_id_from_uppercase_name("SUM"),
            function_id_from_name("sum")
        );
        assert_eq!(
            super::function_id_from_uppercase_name("_XLFN.CONCAT"),
            function_id_from_name("_xlfn.concat")
        );
        assert_eq!(super::function_id_from_uppercase_name(""), None);
    }

    #[test]
    fn future_udf_functions_match_ooxml_xlfn_required_list() {
        let xlsx_src = include_str!(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../formula-xlsx/src/formula_text.rs"
        ));
        let xlsx = extract_const_str_list(xlsx_src, "XL_FN_REQUIRED_FUNCTIONS");
        let biff = FUTURE_UDF_FUNCTIONS.iter().map(|s| (*s).to_string()).collect::<Vec<_>>();
        assert_eq!(
            xlsx, biff,
            "The OOXML `_xlfn.` required-function list (formula-xlsx) must match the BIFF future/UDF list (formula-biff)"
        );
    }
}
