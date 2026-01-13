use formula_biff::decode_rgce;
use pretty_assertions::assert_eq;

fn ptg_int(n: u16) -> [u8; 3] {
    let [lo, hi] = n.to_le_bytes();
    [0x1E, lo, hi] // PtgInt
}

fn ptg_namex(ixti: u16, name_index: u16) -> [u8; 5] {
    let [ixti_lo, ixti_hi] = ixti.to_le_bytes();
    let [idx_lo, idx_hi] = name_index.to_le_bytes();
    [0x39, ixti_lo, ixti_hi, idx_lo, idx_hi] // PtgNameX
}

fn ptg_funcvar_udf(argc: u8) -> [u8; 4] {
    // PtgFuncVar(argc, iftab=0x00FF)
    [0x22, argc, 0xFF, 0x00]
}

#[test]
fn decodes_udf_call_via_namex_and_sentinel_funcvar() {
    // Excel add-in / UDF call pattern:
    //   args..., PtgNameX(func), PtgFuncVar(argc+1, 0x00FF)
    let mut rgce = Vec::new();
    rgce.extend_from_slice(&ptg_int(1));
    rgce.extend_from_slice(&ptg_int(2));
    rgce.extend_from_slice(&ptg_namex(1, 2));
    rgce.extend_from_slice(&ptg_funcvar_udf(3));

    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "ExternName1:2(1,2)");
}

#[test]
fn decodes_3d_refs_with_placeholder_sheet_prefix() {
    // PtgRef3d: [ptg][ixti: u16][row: u32][col: u16]
    let mut ref3d = vec![0x3A];
    ref3d.extend_from_slice(&7u16.to_le_bytes());
    ref3d.extend_from_slice(&0u32.to_le_bytes()); // A1 row=0
    ref3d.extend_from_slice(&0xC000u16.to_le_bytes()); // A, relative row/col
    assert_eq!(decode_rgce(&ref3d).expect("decode"), "'Sheet7'!A1");

    // PtgArea3d: [ptg][ixti: u16][rowFirst: u32][rowLast: u32][colFirst: u16][colLast: u16]
    let mut area3d = vec![0x3B];
    area3d.extend_from_slice(&3u16.to_le_bytes());
    area3d.extend_from_slice(&0u32.to_le_bytes()); // A1
    area3d.extend_from_slice(&1u32.to_le_bytes()); // B2 rowLast=1
    area3d.extend_from_slice(&0xC000u16.to_le_bytes()); // colFirst=A
    area3d.extend_from_slice(&0xC001u16.to_le_bytes()); // colLast=B
    assert_eq!(decode_rgce(&area3d).expect("decode"), "'Sheet3'!A1:B2");
}

#[test]
fn skips_memfunc_and_attrchoose_payloads() {
    // Ensure non-printing tokens that carry payload bytes (PtgMemFunc + PtgAttr(tAttrChoose))
    // are consumed so subsequent tokens are decoded correctly.
    let mut rgce = Vec::new();
    rgce.extend_from_slice(&ptg_int(1));

    // PtgMemFunc: [ptg=0x29][cce: u16][cce bytes...]
    rgce.push(0x29);
    rgce.extend_from_slice(&3u16.to_le_bytes());
    rgce.extend_from_slice(&[0x30, 0x30, 0x30]); // would be invalid ptgs if not skipped

    // PtgAttr(tAttrChoose): [ptg=0x19][grbit=0x04][wAttr: u16][jump_table...]
    rgce.push(0x19);
    rgce.push(0x04); // tAttrChoose
    rgce.extend_from_slice(&2u16.to_le_bytes()); // wAttr=2 -> 4 jump-table bytes
    rgce.extend_from_slice(&[0x30, 0x30, 0x30, 0x30]); // would desync if not skipped

    rgce.extend_from_slice(&ptg_int(2));
    rgce.push(0x03); // PtgAdd

    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "1+2");
}

