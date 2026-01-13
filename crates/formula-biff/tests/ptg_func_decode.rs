use formula_biff::{decode_rgce, function_id_to_name, function_spec_from_id};
use pretty_assertions::assert_eq;

#[cfg(feature = "encode")]
use formula_biff::encode_rgce;

fn ptg_int(n: u16) -> [u8; 3] {
    let [lo, hi] = n.to_le_bytes();
    [0x1E, lo, hi] // PtgInt
}

fn ptg_func(iftab: u16) -> [u8; 3] {
    let [lo, hi] = iftab.to_le_bytes();
    [0x21, lo, hi] // PtgFunc
}

#[test]
fn decodes_ptgfunc_fixed_arity_functions() {
    // 0-arg
    assert_eq!(decode_rgce(&ptg_func(74)).expect("decode NOW"), "NOW()");
    assert_eq!(decode_rgce(&ptg_func(221)).expect("decode TODAY"), "TODAY()");
    assert_eq!(decode_rgce(&ptg_func(19)).expect("decode PI"), "PI()");
    assert_eq!(decode_rgce(&ptg_func(89)).expect("decode CALLER"), "CALLER()");

    // 1-arg
    let mut abs = Vec::new();
    abs.extend_from_slice(&ptg_int(1));
    abs.extend_from_slice(&ptg_func(24)); // ABS
    assert_eq!(decode_rgce(&abs).expect("decode ABS"), "ABS(1)");

    let mut intf = Vec::new();
    intf.extend_from_slice(&ptg_int(1));
    intf.extend_from_slice(&ptg_func(25)); // INT
    assert_eq!(decode_rgce(&intf).expect("decode INT"), "INT(1)");

    let mut sin = Vec::new();
    sin.extend_from_slice(&ptg_int(1));
    sin.extend_from_slice(&ptg_func(15)); // SIN
    assert_eq!(decode_rgce(&sin).expect("decode SIN"), "SIN(1)");

    let mut cos = Vec::new();
    cos.extend_from_slice(&ptg_int(1));
    cos.extend_from_slice(&ptg_func(16)); // COS
    assert_eq!(decode_rgce(&cos).expect("decode COS"), "COS(1)");

    // 2-arg
    let mut round = Vec::new();
    round.extend_from_slice(&ptg_int(1));
    round.extend_from_slice(&ptg_int(2));
    round.extend_from_slice(&ptg_func(27)); // ROUND
    assert_eq!(decode_rgce(&round).expect("decode ROUND"), "ROUND(1,2)");

    let mut get_cell = Vec::new();
    get_cell.extend_from_slice(&ptg_int(1));
    get_cell.extend_from_slice(&ptg_int(2));
    get_cell.extend_from_slice(&ptg_func(185)); // GET.CELL
    assert_eq!(
        decode_rgce(&get_cell).expect("decode GET.CELL"),
        "GET.CELL(1,2)"
    );

    // 4-arg
    let mut series = Vec::new();
    series.extend_from_slice(&ptg_int(1));
    series.extend_from_slice(&ptg_int(2));
    series.extend_from_slice(&ptg_int(3));
    series.extend_from_slice(&ptg_int(4));
    series.extend_from_slice(&ptg_func(92)); // SERIES
    assert_eq!(
        decode_rgce(&series).expect("decode SERIES"),
        "SERIES(1,2,3,4)"
    );

    // 1-arg XLM helpers commonly found in defined names / chart-related formulas.
    let mut evaluate = Vec::new();
    evaluate.extend_from_slice(&ptg_int(1));
    evaluate.extend_from_slice(&ptg_func(257)); // EVALUATE
    assert_eq!(
        decode_rgce(&evaluate).expect("decode EVALUATE"),
        "EVALUATE(1)"
    );

    let mut get_workbook = Vec::new();
    get_workbook.extend_from_slice(&ptg_int(1));
    get_workbook.extend_from_slice(&ptg_func(268)); // GET.WORKBOOK
    assert_eq!(
        decode_rgce(&get_workbook).expect("decode GET.WORKBOOK"),
        "GET.WORKBOOK(1)"
    );

    let mut get_workspace = Vec::new();
    get_workspace.extend_from_slice(&ptg_int(1));
    get_workspace.extend_from_slice(&ptg_func(186)); // GET.WORKSPACE
    assert_eq!(
        decode_rgce(&get_workspace).expect("decode GET.WORKSPACE"),
        "GET.WORKSPACE(1)"
    );

    let mut get_window = Vec::new();
    get_window.extend_from_slice(&ptg_int(1));
    get_window.extend_from_slice(&ptg_func(187)); // GET.WINDOW
    assert_eq!(
        decode_rgce(&get_window).expect("decode GET.WINDOW"),
        "GET.WINDOW(1)"
    );

    let mut get_document = Vec::new();
    get_document.extend_from_slice(&ptg_int(1));
    get_document.extend_from_slice(&ptg_func(188)); // GET.DOCUMENT
    assert_eq!(
        decode_rgce(&get_document).expect("decode GET.DOCUMENT"),
        "GET.DOCUMENT(1)"
    );
}

#[cfg(feature = "encode")]
#[test]
fn encode_roundtrips_for_new_ptgfunc_functions() {
    // Ensure `function_spec_from_name` consults expanded metadata and that the encoder
    // picks `PtgFunc` for fixed-arity built-ins.
    for formula in [
        "PI()",
        "SIN(1)",
        "COS(1)",
        "ROUND(1,2)",
        "TODAY()",
        "NOW()",
        "CALLER()",
        "EVALUATE(1)",
        "GET.CELL(1,2)",
        "GET.WORKBOOK(1)",
        "GET.WORKSPACE(1)",
        "GET.WINDOW(1)",
        "GET.DOCUMENT(1)",
        "SERIES(1,2,3,4)",
    ] {
        let rgce = encode_rgce(formula).expect("encode");
        let decoded = decode_rgce(&rgce).expect("decode");
        assert_eq!(decoded, formula);
    }
}

#[test]
fn function_spec_coverage_for_ftab_ids() {
    // If an FTAB entry has a non-empty name, `function_spec_from_id` should either:
    // - return a spec with a valid argument range, or
    // - be explicitly excluded because it is an XLM macro / command function that we
    //   do not currently support in the BIFF encoder / `PtgFunc` decoder.
    //
    // The excluded IDs correspond to functions that appear in macro sheets (Excel 4.0)
    // and UI/command helpers. BIFF formulas in normal worksheets do not use them.
    const UNSUPPORTED_IDS: &[u16] = &[
        53, 54, 55, 79, 80, 81, 84, 85, 87, 88, 90, 91, 93, 94, 95, 96, 103, 104, 106,
        107, 108, 110, 122, 123, 132, 133, 134, 135, 136, 137, 138, 139, 145, 146, 147, 149,
        150, 151, 152, 153, 154, 155, 156, 157, 158, 159, 160, 161, 166, 170, 171, 172, 173,
        174, 175, 176, 177, 178, 179, 180, 181, 182, 191, 192, 200, 201,
        223, 224, 225, 226, 236, 237, 238, 239, 240, 241, 242, 243, 245, 246, 248, 251, 253,
        254, 255, 256, 258, 259, 260, 262, 263, 264, 265, 266, 267, 334, 335, 338,
        339, 340, 341, 348, 349, 352, 353, 355, 356, 357,
    ];

    for id in 0u16..=484u16 {
        let Some(name) = function_id_to_name(id) else {
            continue;
        };

        match function_spec_from_id(id) {
            Some(spec) => {
                assert_eq!(spec.id, id);
                assert_eq!(spec.name, name);
                assert!(spec.min_args <= spec.max_args, "{id} {name} invalid arg range");
            }
            None => {
                assert!(
                    UNSUPPORTED_IDS.contains(&id),
                    "missing FunctionSpec for FTAB id={id} name={name}"
                );
            }
        }
    }
}
