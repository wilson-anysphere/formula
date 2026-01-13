use formula_biff::{decode_rgce, DecodeRgceError};
use pretty_assertions::assert_eq;

fn normalize(formula: &str) -> String {
    let ast = formula_engine::parse_formula(formula, formula_engine::ParseOptions::default())
        .expect("parse formula");
    ast.to_string(formula_engine::SerializeOptions {
        omit_equals: true,
        ..Default::default()
    })
    .expect("serialize formula")
}

/// Build a BIFF12 structured reference token (`PtgList`) encoded as `PtgExtend` + `etpg=0x19`.
///
/// Payload layout (MS-XLSB 2.5.198.51):
/// `[table_id: u32][flags: u16][col_first: u16][col_last: u16][reserved: u16]`.
fn ptg_list(table_id: u32, flags: u16, col_first: u16, col_last: u16, ptg: u8) -> Vec<u8> {
    let mut out = Vec::new();
    out.push(ptg);
    out.push(0x19); // etpg=0x19 (PtgList)
    out.extend_from_slice(&table_id.to_le_bytes());
    out.extend_from_slice(&flags.to_le_bytes());
    out.extend_from_slice(&col_first.to_le_bytes());
    out.extend_from_slice(&col_last.to_le_bytes());
    out.extend_from_slice(&0u16.to_le_bytes()); // reserved
    out
}

fn ptg_int(n: u16) -> [u8; 3] {
    let [lo, hi] = n.to_le_bytes();
    [0x1E, lo, hi] // PtgInt
}

fn ptg_funcvar(argc: u8, iftab: u16) -> [u8; 4] {
    let [lo, hi] = iftab.to_le_bytes();
    [0x22, argc, lo, hi] // PtgFuncVar
}

fn ptg_paren() -> [u8; 1] {
    [0x15] // PtgParen
}

fn ptg_percent() -> [u8; 1] {
    [0x14] // PtgPercent
}

fn ptg_spill_range() -> [u8; 1] {
    [0x2F] // PtgSpillRange
}

#[test]
fn decodes_structured_ref_union_inside_function_arg() {
    // Build SUM((Table1[Column2],Table1[Column4])) to ensure union-containing args are
    // parenthesized (union uses ',' which is also the function arg separator).
    let mut rgce = ptg_list(1, 0x0000, 2, 2, 0x18);
    rgce.extend_from_slice(&ptg_list(1, 0x0000, 4, 4, 0x18));
    rgce.push(0x10); // PtgUnion
    rgce.extend_from_slice(&ptg_funcvar(1, 4)); // SUM
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(
        normalize(&text),
        normalize("SUM((Table1[Column2],Table1[Column4]))")
    );
}

#[test]
fn decodes_structured_ref_ignores_reserved_field() {
    let mut rgce = ptg_list(1, 0x0000, 2, 2, 0x18);
    // Reserved u16 is the final 2 bytes of the fixed 12-byte payload.
    let len = rgce.len();
    rgce[len - 2] = 0x34;
    rgce[len - 1] = 0x12;
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "Table1[Column2]");
    assert_eq!(normalize(&text), normalize("Table1[Column2]"));
}

#[test]
fn decodes_structured_ref_uses_stable_placeholder_names_for_unknown_ids() {
    let rgce = ptg_list(42, 0x0000, 7, 7, 0x18);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "Table42[Column7]");
    assert_eq!(normalize(&text), normalize("Table42[Column7]"));
}

#[test]
fn decodes_structured_ref_with_postfix_ops() {
    // Exercise postfix operators on a structured reference token.
    let mut rgce = ptg_list(1, 0x0000, 2, 2, 0x18);
    rgce.extend_from_slice(&ptg_percent());
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(normalize(&text), normalize("Table1[Column2]%"));

    let mut rgce = ptg_list(1, 0x0000, 2, 2, 0x18);
    rgce.extend_from_slice(&ptg_spill_range());
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(normalize(&text), normalize("Table1[Column2]#"));
}

#[test]
fn decodes_structured_ref_with_paren() {
    let mut rgce = ptg_list(1, 0x0000, 2, 2, 0x18);
    rgce.extend_from_slice(&ptg_paren());
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(normalize(&text), normalize("(Table1[Column2])"));
}

#[test]
fn decodes_value_class_structured_ref_with_postfix_ops() {
    // Value-class structured refs may start with '@'; ensure postfix operators still work.
    let mut rgce = ptg_list(1, 0x0000, 2, 2, 0x38);
    rgce.extend_from_slice(&ptg_spill_range());
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(normalize(&text), normalize("@Table1[Column2]#"));
}

#[test]
fn decodes_structured_ref_table_column() {
    let rgce = ptg_list(1, 0x0000, 2, 2, 0x18);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "Table1[Column2]");
    assert_eq!(normalize(&text), normalize("Table1[Column2]"));
}

#[test]
fn decodes_structured_ref_this_row() {
    // Best-effort: 0x0010 is treated as "#This Row" and rendered using the `[@Col]` shorthand.
    let rgce = ptg_list(1, 0x0010, 2, 2, 0x18);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "[@Column2]");
    assert_eq!(normalize(&text), normalize("[@Column2]"));
}

#[test]
fn decodes_structured_ref_this_row_all_columns() {
    let rgce = ptg_list(1, 0x0010, 0, 0, 0x18);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "[@]");
    assert_eq!(normalize(&text), normalize("[@]"));
}

#[test]
fn decodes_structured_ref_headers_column() {
    let rgce = ptg_list(1, 0x0002, 2, 2, 0x18);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "Table1[[#Headers],[Column2]]");
    assert_eq!(
        normalize(&text),
        normalize("Table1[[#Headers],[Column2]]")
    );
}

#[test]
fn decodes_structured_ref_totals_column() {
    let rgce = ptg_list(1, 0x0008, 2, 2, 0x18);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "Table1[[#Totals],[Column2]]");
    assert_eq!(normalize(&text), normalize("Table1[[#Totals],[Column2]]"));
}

#[test]
fn decodes_structured_ref_item_only_headers() {
    let rgce = ptg_list(1, 0x0002, 0, 0, 0x18);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "Table1[#Headers]");
    assert_eq!(normalize(&text), normalize("Table1[#Headers]"));
}

#[test]
fn decodes_structured_ref_item_only_totals() {
    let rgce = ptg_list(1, 0x0008, 0, 0, 0x18);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "Table1[#Totals]");
    assert_eq!(normalize(&text), normalize("Table1[#Totals]"));
}

#[test]
fn decodes_structured_ref_all_column() {
    let rgce = ptg_list(1, 0x0001, 2, 2, 0x18);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "Table1[[#All],[Column2]]");
    assert_eq!(normalize(&text), normalize("Table1[[#All],[Column2]]"));
}

#[test]
fn decodes_structured_ref_headers_column_range() {
    let rgce = ptg_list(1, 0x0002, 2, 4, 0x18);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "Table1[[#Headers],[Column2]:[Column4]]");
    assert_eq!(
        normalize(&text),
        normalize("Table1[[#Headers],[Column2]:[Column4]]")
    );
}

#[test]
fn decodes_structured_ref_item_only_all() {
    let rgce = ptg_list(1, 0x0001, 0, 0, 0x18);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "Table1[#All]");
    assert_eq!(normalize(&text), normalize("Table1[#All]"));
}

#[test]
fn decodes_structured_ref_item_only_data() {
    let rgce = ptg_list(1, 0x0004, 0, 0, 0x18);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "Table1[#Data]");
    assert_eq!(normalize(&text), normalize("Table1[#Data]"));
}

#[test]
fn decodes_structured_ref_data_flag_with_column_uses_simple_form() {
    // `#Data` is the default row selector for a column reference, so Excel canonical text omits it.
    let rgce = ptg_list(1, 0x0004, 2, 2, 0x18);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "Table1[Column2]");
    assert_eq!(normalize(&text), normalize("Table1[Column2]"));
}

#[test]
fn decodes_structured_ref_data_flag_with_column_range_uses_simple_form() {
    // `#Data` is the default row selector for a column range reference, so Excel canonical text
    // omits it.
    let rgce = ptg_list(1, 0x0004, 2, 4, 0x18);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "Table1[[Column2]:[Column4]]");
    assert_eq!(
        normalize(&text),
        normalize("Table1[[Column2]:[Column4]]")
    );
}

#[test]
fn decodes_structured_ref_default_data_when_no_item_and_no_columns() {
    // If both the row/item flags and the column selectors are absent, Excel treats the reference
    // as the table's data body.
    let rgce = ptg_list(1, 0x0000, 0, 0, 0x18);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "Table1[#Data]");
    assert_eq!(normalize(&text), normalize("Table1[#Data]"));
}

#[test]
fn decodes_structured_ref_column_range() {
    let rgce = ptg_list(1, 0x0000, 2, 4, 0x18);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "Table1[[Column2]:[Column4]]");
    assert_eq!(
        normalize(&text),
        normalize("Table1[[Column2]:[Column4]]")
    );
}

#[test]
fn decodes_structured_ref_this_row_range() {
    let rgce = ptg_list(1, 0x0010, 2, 4, 0x18);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "[@[Column2]:[Column4]]");
    assert_eq!(normalize(&text), normalize("[@[Column2]:[Column4]]"));
}

#[test]
fn decodes_structured_ref_value_class_column_range_adds_implicit_intersection() {
    let rgce = ptg_list(1, 0x0000, 2, 4, 0x38);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "@Table1[[Column2]:[Column4]]");
    assert_eq!(normalize(&text), normalize("@Table1[[Column2]:[Column4]]"));
}

#[test]
fn decodes_structured_ref_value_class_emits_explicit_implicit_intersection() {
    let rgce = ptg_list(1, 0x0000, 2, 2, 0x38);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "@Table1[Column2]");
    assert_eq!(normalize(&text), normalize("@Table1[Column2]"));
}

#[test]
fn decodes_structured_ref_in_binary_expr() {
    let mut rgce = ptg_list(1, 0x0000, 2, 2, 0x18);
    rgce.extend_from_slice(&ptg_int(1));
    rgce.push(0x03); // PtgAdd
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(normalize(&text), normalize("Table1[Column2]+1"));
}

#[test]
fn decodes_value_class_structured_ref_in_binary_expr() {
    let mut rgce = ptg_list(1, 0x0000, 2, 2, 0x38);
    rgce.extend_from_slice(&ptg_int(1));
    rgce.push(0x03); // PtgAdd
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(normalize(&text), normalize("@Table1[Column2]+1"));
}

#[test]
fn decodes_structured_ref_used_as_function_arg() {
    let mut rgce = ptg_list(1, 0x0000, 2, 2, 0x18);
    // SUM is id=4 in Excel's function table.
    rgce.extend_from_slice(&ptg_funcvar(1, 4));
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(normalize(&text), normalize("SUM(Table1[Column2])"));
}

#[test]
fn decodes_structured_ref_value_class_headers_single_does_not_add_implicit_intersection() {
    // Headers+single column resolves to a single cell, so value-class should not force '@'.
    let rgce = ptg_list(1, 0x0002, 2, 2, 0x38);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "Table1[[#Headers],[Column2]]");
    assert_eq!(normalize(&text), normalize("Table1[[#Headers],[Column2]]"));
}

#[test]
fn decodes_structured_ref_value_class_totals_single_does_not_add_implicit_intersection() {
    // Totals+single column resolves to a single cell, so value-class should not force '@'.
    let rgce = ptg_list(1, 0x0008, 2, 2, 0x38);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "Table1[[#Totals],[Column2]]");
    assert_eq!(normalize(&text), normalize("Table1[[#Totals],[Column2]]"));
}

#[test]
fn decodes_structured_ref_value_class_this_row_single_does_not_add_implicit_intersection() {
    // This-row+single column resolves to a single cell.
    let rgce = ptg_list(1, 0x0010, 2, 2, 0x38);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "[@Column2]");
    assert_eq!(normalize(&text), normalize("[@Column2]"));
}

#[test]
fn decodes_structured_ref_value_class_this_row_all_columns_adds_implicit_intersection() {
    // This-row+all columns resolves to a row range; value-class should preserve legacy implicit
    // intersection by prefixing '@'.
    let rgce = ptg_list(1, 0x0010, 0, 0, 0x38);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "@[@]");
    assert_eq!(normalize(&text), normalize("@[@]"));
}

#[test]
fn decodes_structured_ref_value_class_this_row_range_adds_implicit_intersection() {
    let rgce = ptg_list(1, 0x0010, 2, 4, 0x38);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "@[@[Column2]:[Column4]]");
    assert_eq!(normalize(&text), normalize("@[@[Column2]:[Column4]]"));
}

#[test]
fn decodes_structured_ref_value_class_headers_all_columns_adds_implicit_intersection() {
    let rgce = ptg_list(1, 0x0002, 0, 0, 0x38);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "@Table1[#Headers]");
    assert_eq!(normalize(&text), normalize("@Table1[#Headers]"));
}

#[test]
fn decodes_structured_ref_value_class_headers_column_range_adds_implicit_intersection() {
    let rgce = ptg_list(1, 0x0002, 2, 4, 0x38);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "@Table1[[#Headers],[Column2]:[Column4]]");
    assert_eq!(
        normalize(&text),
        normalize("@Table1[[#Headers],[Column2]:[Column4]]")
    );
}

#[test]
fn decodes_structured_ref_value_class_totals_all_columns_adds_implicit_intersection() {
    let rgce = ptg_list(1, 0x0008, 0, 0, 0x38);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "@Table1[#Totals]");
    assert_eq!(normalize(&text), normalize("@Table1[#Totals]"));
}

#[test]
fn decodes_structured_ref_value_class_all_adds_implicit_intersection() {
    let rgce = ptg_list(1, 0x0001, 0, 0, 0x38);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "@Table1[#All]");
    assert_eq!(normalize(&text), normalize("@Table1[#All]"));
}

#[test]
fn decodes_structured_ref_value_class_all_single_column_adds_implicit_intersection() {
    let rgce = ptg_list(1, 0x0001, 2, 2, 0x38);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "@Table1[[#All],[Column2]]");
    assert_eq!(normalize(&text), normalize("@Table1[[#All],[Column2]]"));
}

#[test]
fn decodes_structured_ref_value_class_data_adds_implicit_intersection() {
    let rgce = ptg_list(1, 0x0004, 0, 0, 0x38);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "@Table1[#Data]");
    assert_eq!(normalize(&text), normalize("@Table1[#Data]"));
}

#[test]
fn decodes_structured_ref_unknown_flags_are_ignored() {
    // Unknown flag bits should not hard-fail decode.
    let rgce = ptg_list(1, 0x8000, 2, 2, 0x18);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "Table1[Column2]");
    assert_eq!(normalize(&text), normalize("Table1[Column2]"));
}

#[test]
fn decodes_structured_ref_unknown_flags_preserve_known_item_bits() {
    let rgce = ptg_list(1, 0x8000 | 0x0002, 2, 2, 0x18);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "Table1[[#Headers],[Column2]]");
    assert_eq!(normalize(&text), normalize("Table1[[#Headers],[Column2]]"));
}

#[test]
fn decodes_structured_ref_unknown_flags_preserve_known_this_row_bit() {
    let rgce = ptg_list(1, 0x8000 | 0x0010, 2, 2, 0x18);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "[@Column2]");
    assert_eq!(normalize(&text), normalize("[@Column2]"));
}

#[test]
fn decodes_structured_ref_multiple_item_flags_prefers_headers() {
    // Excel's flags are not strictly mutually exclusive; ensure we stay best-effort by choosing a
    // stable priority order (Headers > Totals > All > Data), matching formula-xlsb's decoder.
    let rgce = ptg_list(1, 0x0002 | 0x0004, 2, 2, 0x18);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "Table1[[#Headers],[Column2]]");
    assert_eq!(normalize(&text), normalize("Table1[[#Headers],[Column2]]"));
}

#[test]
fn decodes_structured_ref_multiple_item_flags_prefers_totals_over_data() {
    let rgce = ptg_list(1, 0x0008 | 0x0004, 2, 2, 0x18);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "Table1[[#Totals],[Column2]]");
    assert_eq!(normalize(&text), normalize("Table1[[#Totals],[Column2]]"));
}

#[test]
fn decodes_structured_ref_multiple_item_flags_prefers_this_row() {
    // Ensure the "#This Row" flag wins over other item flags.
    let rgce = ptg_list(1, 0x0010 | 0x0002, 2, 2, 0x18);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "[@Column2]");
    assert_eq!(normalize(&text), normalize("[@Column2]"));
}

#[test]
fn decodes_structured_ref_multiple_item_flags_prefers_this_row_all_columns() {
    let rgce = ptg_list(1, 0x0010 | 0x0001, 0, 0, 0x18);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "[@]");
    assert_eq!(normalize(&text), normalize("[@]"));
}

#[test]
fn decodes_structured_ref_extend_a_is_supported() {
    // PtgExtendA (0x58) uses the same payload as ref/value class variants.
    let rgce = ptg_list(1, 0x0000, 2, 2, 0x58);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "Table1[Column2]");
    assert_eq!(normalize(&text), normalize("Table1[Column2]"));
}

#[test]
fn rejects_legacy_placeholder_ptglist_token() {
    // Older versions of `formula-biff` used a non-standard, string-based "PtgList" token with id
    // 0x30. Ensure we reject it so callers don't accidentally rely on it.
    let rgce = vec![
        0x30, // legacy placeholder token id
        0x00, 0x00, // flags
        0x00, 0x00, // table name (utf16 len=0)
        0x00, 0x00, // col1 (utf16 len=0)
        0x00, 0x00, // col2 (utf16 len=0)
    ];
    let err = decode_rgce(&rgce).unwrap_err();
    assert!(matches!(
        err,
        DecodeRgceError::UnsupportedToken {
            offset: 0,
            ptg: 0x30
        }
    ));
}

#[test]
fn ptg_extend_missing_etpg_is_unexpected_eof() {
    // PtgExtend without the following etpg subtype byte.
    let err = decode_rgce(&[0x18]).unwrap_err();
    assert!(matches!(
        err,
        DecodeRgceError::UnexpectedEof {
            offset: 0,
            ptg: 0x18,
            needed: 1,
            remaining: 0
        }
    ));
}

#[test]
fn ptg_extend_unknown_etpg_is_unsupported() {
    let err = decode_rgce(&[0x18, 0xFF]).unwrap_err();
    assert!(matches!(
        err,
        DecodeRgceError::UnsupportedToken {
            offset: 0,
            ptg: 0x18
        }
    ));
}

#[test]
fn ptg_extend_list_truncated_payload_is_unexpected_eof() {
    // PtgExtend + PtgList etpg, but not enough bytes for the fixed 12-byte payload.
    let rgce = vec![0x18, 0x19, 0x00, 0x00, 0x00, 0x00, 0x00];
    let err = decode_rgce(&rgce).unwrap_err();
    assert!(matches!(
        err,
        DecodeRgceError::UnexpectedEof {
            offset: 0,
            ptg: 0x18,
            needed: 12,
            remaining: 5
        }
    ));
}
