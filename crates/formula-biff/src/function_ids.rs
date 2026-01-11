#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FunctionSpec {
    pub id: u16,
    pub name: &'static str,
    pub min_args: u8,
    pub max_args: u8,
}

// NOTE: Function IDs are BIFF built-in function indices (the `iftab` values used
// by `PtgFunc`/`PtgFuncVar`). These are shared across BIFF8/BIFF12 for
// "classic" Excel functions.
//
// The encoder/decoder currently only needs a small subset for early editing
// workflows (matching the functions implemented by `formula-engine`).
pub(crate) const FUNCTION_SPECS: &[FunctionSpec] = &[
    // Statistics / Math
    FunctionSpec {
        id: 0x0000,
        name: "COUNT",
        min_args: 1,
        max_args: 255,
    },
    FunctionSpec {
        id: 0x0004,
        name: "SUM",
        min_args: 1,
        max_args: 255,
    },
    FunctionSpec {
        id: 0x0005,
        name: "AVERAGE",
        min_args: 1,
        max_args: 255,
    },
    FunctionSpec {
        id: 0x0006,
        name: "MIN",
        min_args: 1,
        max_args: 255,
    },
    FunctionSpec {
        id: 0x0007,
        name: "MAX",
        min_args: 1,
        max_args: 255,
    },
    FunctionSpec {
        id: 0x0018,
        name: "ABS",
        min_args: 1,
        max_args: 1,
    },
    FunctionSpec {
        id: 0x0019,
        name: "INT",
        min_args: 1,
        max_args: 1,
    },
    FunctionSpec {
        id: 0x001B,
        name: "ROUND",
        min_args: 2,
        max_args: 2,
    },
    FunctionSpec {
        id: 0x001D,
        name: "INDEX",
        min_args: 2,
        max_args: 4,
    },
    FunctionSpec {
        id: 0x0027,
        name: "MOD",
        min_args: 2,
        max_args: 2,
    },
    // Text
    FunctionSpec {
        id: 0x001F,
        name: "MID",
        min_args: 3,
        max_args: 3,
    },
    FunctionSpec {
        id: 0x0020,
        name: "LEN",
        min_args: 1,
        max_args: 1,
    },
    FunctionSpec {
        id: 0x0073,
        name: "LEFT",
        min_args: 1,
        max_args: 2,
    },
    FunctionSpec {
        id: 0x0074,
        name: "RIGHT",
        min_args: 1,
        max_args: 2,
    },
    FunctionSpec {
        id: 0x0076,
        name: "TRIM",
        min_args: 1,
        max_args: 1,
    },
    FunctionSpec {
        id: 0x0077,
        name: "UPPER",
        min_args: 1,
        max_args: 1,
    },
    FunctionSpec {
        id: 0x0078,
        name: "LOWER",
        min_args: 1,
        max_args: 1,
    },
    FunctionSpec {
        id: 0x007C,
        name: "FIND",
        min_args: 2,
        max_args: 3,
    },
    FunctionSpec {
        id: 0x0052,
        name: "SEARCH",
        min_args: 2,
        max_args: 3,
    },
    FunctionSpec {
        id: 0x0150,
        name: "CONCATENATE",
        min_args: 1,
        max_args: 255,
    },
    // Logical
    FunctionSpec {
        id: 0x0001,
        name: "IF",
        min_args: 2,
        max_args: 3,
    },
    FunctionSpec {
        id: 0x0003,
        name: "ISERROR",
        min_args: 1,
        max_args: 1,
    },
    FunctionSpec {
        id: 0x0024,
        name: "AND",
        min_args: 1,
        max_args: 255,
    },
    FunctionSpec {
        id: 0x0025,
        name: "OR",
        min_args: 1,
        max_args: 255,
    },
    FunctionSpec {
        id: 0x0026,
        name: "NOT",
        min_args: 1,
        max_args: 1,
    },
    // Lookup
    FunctionSpec {
        id: 0x0040,
        name: "MATCH",
        min_args: 2,
        max_args: 3,
    },
    FunctionSpec {
        id: 0x0065,
        name: "HLOOKUP",
        min_args: 3,
        max_args: 4,
    },
    FunctionSpec {
        id: 0x0066,
        name: "VLOOKUP",
        min_args: 3,
        max_args: 4,
    },
    // Date/time
    FunctionSpec {
        id: 0x0041,
        name: "DATE",
        min_args: 3,
        max_args: 3,
    },
    FunctionSpec {
        id: 0x0043,
        name: "DAY",
        min_args: 1,
        max_args: 1,
    },
    FunctionSpec {
        id: 0x0044,
        name: "MONTH",
        min_args: 1,
        max_args: 1,
    },
    FunctionSpec {
        id: 0x0045,
        name: "YEAR",
        min_args: 1,
        max_args: 1,
    },
    FunctionSpec {
        id: 0x004A,
        name: "NOW",
        min_args: 0,
        max_args: 0,
    },
    FunctionSpec {
        id: 0x00DD,
        name: "TODAY",
        min_args: 0,
        max_args: 0,
    },
    // Newer functions (Excel 2007+). These IDs are stable in BIFF12 but may not
    // be present in BIFF8.
    FunctionSpec {
        id: 0x0159,
        name: "IFERROR",
        min_args: 2,
        max_args: 2,
    },
    FunctionSpec {
        id: 0x0165,
        name: "IFNA",
        min_args: 2,
        max_args: 2,
    },
    FunctionSpec {
        id: 0x00A9,
        name: "COUNTA",
        min_args: 1,
        max_args: 255,
    },
    FunctionSpec {
        id: 0x015B,
        name: "COUNTBLANK",
        min_args: 1,
        max_args: 1,
    },
    FunctionSpec {
        id: 0x00D4,
        name: "ROUNDUP",
        min_args: 2,
        max_args: 2,
    },
    FunctionSpec {
        id: 0x00D5,
        name: "ROUNDDOWN",
        min_args: 2,
        max_args: 2,
    },
    // Excel 365 functions. These are typically stored with `_xlfn.` prefix in
    // OOXML; BIFF12 uses function IDs directly.
    FunctionSpec {
        id: 0x01C6,
        name: "CONCAT",
        min_args: 1,
        max_args: 255,
    },
    FunctionSpec {
        id: 0x01C7,
        name: "XMATCH",
        min_args: 2,
        max_args: 4,
    },
    FunctionSpec {
        id: 0x01C8,
        name: "XLOOKUP",
        min_args: 3,
        max_args: 6,
    },
];

pub fn function_name_to_id(name: &str) -> Option<u16> {
    let upper = name.trim().to_ascii_uppercase();
    FUNCTION_SPECS
        .iter()
        .find(|spec| spec.name == upper)
        .map(|spec| spec.id)
}

pub fn function_id_to_name(id: u16) -> Option<&'static str> {
    FUNCTION_SPECS
        .iter()
        .find(|spec| spec.id == id)
        .map(|spec| spec.name)
}

pub fn function_spec_from_id(id: u16) -> Option<FunctionSpec> {
    FUNCTION_SPECS.iter().find(|spec| spec.id == id).copied()
}

#[cfg(feature = "encode")]
pub(crate) fn function_spec_from_name(name: &str) -> Option<FunctionSpec> {
    let upper = name.trim().to_ascii_uppercase();
    FUNCTION_SPECS
        .iter()
        .find(|spec| spec.name == upper)
        .copied()
}
