use formula_xlsb::rgce::{
    decode_rgce_with_context, encode_rgce_with_context, CellCoord, EncodeError,
};
use formula_xlsb::XlsbWorkbook;
use pretty_assertions::assert_eq;

#[test]
fn addin_supbook_distinguishes_functions_from_other_extern_names() {
    let path = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/udf.xlsb");
    let wb = XlsbWorkbook::open(path).expect("open xlsb");
    let ctx = wb.workbook_context();

    // `PtgNameX` references to add-in functions should render as the bare function name.
    let func_ref = vec![0x39, 0x00, 0x00, 0x01, 0x00]; // PtgNameX(ixti=0, nameIndex=1)
    let decoded = decode_rgce_with_context(&func_ref, ctx).expect("decode add-in function ref");
    assert_eq!(decoded, "MyAddinFunc");

    // Non-function add-in extern names must be qualified so they don't collide with workbook names.
    let name_ref = vec![0x39, 0x00, 0x00, 0x02, 0x00]; // PtgNameX(ixti=0, nameIndex=2)
    let decoded = decode_rgce_with_context(&name_ref, ctx).expect("decode add-in name ref");
    assert_eq!(decoded, "'[AddIn]MyAddinConst'");

    // Encoding should still treat true add-in functions as UDF calls via NameX.
    let encoded =
        encode_rgce_with_context("=MyAddinFunc(1,2)", ctx, CellCoord::new(0, 0)).expect("encode");
    assert_eq!(
        encoded.rgce,
        vec![
            0x1E, 0x01, 0x00, // 1
            0x1E, 0x02, 0x00, // 2
            0x39, 0x00, 0x00, 0x01, 0x00, // PtgNameX(ixti=0, nameIndex=1)
            0x22, 0x03, 0xFF, 0x00, // PtgFuncVar(argc=3, iftab=0x00FF)
        ]
    );
    let decoded = decode_rgce_with_context(&encoded.rgce, ctx).expect("decode");
    assert_eq!(decoded, "MyAddinFunc(1,2)");

    // Non-function add-in extern names must *not* be treated as UDF-callable functions.
    let err = encode_rgce_with_context("=MyAddinConst(1)", ctx, CellCoord::new(0, 0))
        .expect_err("should reject non-function NameX as a function");
    assert_eq!(
        err,
        EncodeError::UnknownFunction {
            name: "MYADDINCONST".to_string()
        }
    );
}

