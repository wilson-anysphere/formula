use formula_biff::decode_rgce;
use formula_engine::{Expr, UnaryOp};
use pretty_assertions::assert_eq;

fn ptg_name(name_id: u32, ptg: u8) -> Vec<u8> {
    let mut out = Vec::new();
    out.push(ptg);
    out.extend_from_slice(&name_id.to_le_bytes());
    out.extend_from_slice(&0u16.to_le_bytes()); // reserved
    out
}

fn ptg_namex(ixti: u16, name_index: u16, ptg: u8) -> Vec<u8> {
    let mut out = Vec::new();
    out.push(ptg);
    out.extend_from_slice(&ixti.to_le_bytes());
    out.extend_from_slice(&name_index.to_le_bytes());
    out
}

fn ptg_int(n: u16) -> [u8; 3] {
    let [lo, hi] = n.to_le_bytes();
    [0x1E, lo, hi] // PtgInt
}

fn ptg_funcvar_udf(argc: u8) -> [u8; 4] {
    // PtgFuncVar(argc, iftab=0x00FF)
    [0x22, argc, 0xFF, 0x00]
}

fn parse(formula: &str) -> formula_engine::Ast {
    formula_engine::parse_formula(formula, formula_engine::ParseOptions::default())
        .expect("parse formula")
}

fn assert_parseable(formula: &str) {
    parse(formula);
}

#[test]
fn decodes_ptg_name_to_safe_placeholder() {
    let rgce = ptg_name(123, 0x23);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "Name_123");
    assert_parseable(&text);
    let ast = parse(&text);
    assert!(
        matches!(&ast.expr, Expr::NameRef(n) if n.name == "Name_123"),
        "expected NameRef(Name_123), got {ast:?}"
    );
}

#[test]
fn decodes_ptg_namex_to_safe_placeholder() {
    let rgce = ptg_namex(0, 1, 0x39);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "ExternName_IXTI0_N1");
    assert_parseable(&text);
    let ast = parse(&text);
    assert!(
        matches!(&ast.expr, Expr::NameRef(n) if n.name == "ExternName_IXTI0_N1"),
        "expected NameRef(ExternName_IXTI0_N1), got {ast:?}"
    );
}

#[test]
fn decodes_value_class_ptg_name_with_implicit_intersection_prefix() {
    let rgce = ptg_name(123, 0x43);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "@Name_123");
    assert_parseable(&text);
    let ast = parse(&text);
    assert!(
        matches!(
            &ast.expr,
            Expr::Unary(u)
                if u.op == UnaryOp::ImplicitIntersection
                    && matches!(&*u.expr, Expr::NameRef(n) if n.name == "Name_123")
        ),
        "expected @NameRef(Name_123), got {ast:?}"
    );
}

#[test]
fn decodes_value_class_ptg_namex_with_implicit_intersection_prefix() {
    let rgce = ptg_namex(0, 1, 0x59);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "@ExternName_IXTI0_N1");
    assert_parseable(&text);
    let ast = parse(&text);
    assert!(
        matches!(
            &ast.expr,
            Expr::Unary(u)
                if u.op == UnaryOp::ImplicitIntersection
                    && matches!(&*u.expr, Expr::NameRef(n) if n.name == "ExternName_IXTI0_N1")
        ),
        "expected @NameRef(ExternName_IXTI0_N1), got {ast:?}"
    );
}

#[test]
fn decodes_array_class_ptg_name_to_safe_placeholder() {
    let rgce = ptg_name(123, 0x63);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "Name_123");
    assert_parseable(&text);
}

#[test]
fn decodes_array_class_ptg_namex_to_safe_placeholder() {
    let rgce = ptg_namex(0, 1, 0x79);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "ExternName_IXTI0_N1");
    assert_parseable(&text);
}

#[test]
fn decodes_ptg_name_as_udf_function_name_via_sentinel_funcvar() {
    // Name-based UDF call pattern:
    //   args..., PtgName(func), PtgFuncVar(argc+1, 0x00FF)
    let mut rgce = Vec::new();
    rgce.extend_from_slice(&ptg_int(1));
    rgce.extend_from_slice(&ptg_int(2));
    rgce.extend_from_slice(&ptg_name(123, 0x23));
    rgce.extend_from_slice(&ptg_funcvar_udf(3));

    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "Name_123(1,2)");
    assert_parseable(&text);
    let ast = parse(&text);
    assert!(
        matches!(&ast.expr, Expr::FunctionCall(call) if call.name.original == "Name_123"),
        "expected FunctionCall(Name_123), got {ast:?}"
    );
}

#[test]
fn decodes_value_class_ptg_name_as_udf_function_name_via_sentinel_funcvar() {
    // Same pattern as `decodes_ptg_name_as_udf_function_name_via_sentinel_funcvar`, but using the
    // value-class `PtgName` variant (0x43). The decoder should preserve the implicit intersection
    // marker.
    let mut rgce = Vec::new();
    rgce.extend_from_slice(&ptg_int(1));
    rgce.extend_from_slice(&ptg_int(2));
    rgce.extend_from_slice(&ptg_name(123, 0x43));
    rgce.extend_from_slice(&ptg_funcvar_udf(3));

    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "@Name_123(1,2)");
    assert_parseable(&text);
    let ast = parse(&text);
    assert!(
        matches!(
            &ast.expr,
            Expr::Unary(u)
                if u.op == UnaryOp::ImplicitIntersection
                    && matches!(&*u.expr, Expr::FunctionCall(call) if call.name.original == "Name_123")
        ),
        "expected @FunctionCall(Name_123), got {ast:?}"
    );
}

#[test]
fn decodes_ptg_namex_as_udf_function_name_via_sentinel_funcvar() {
    // NameX-based UDF call pattern:
    //   args..., PtgNameX(func), PtgFuncVar(argc+1, 0x00FF)
    let mut rgce = Vec::new();
    rgce.extend_from_slice(&ptg_int(1));
    rgce.extend_from_slice(&ptg_int(2));
    rgce.extend_from_slice(&ptg_namex(1, 2, 0x39));
    rgce.extend_from_slice(&ptg_funcvar_udf(3));

    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "ExternName_IXTI1_N2(1,2)");
    assert_parseable(&text);
    let ast = parse(&text);
    assert!(
        matches!(
            &ast.expr,
            Expr::FunctionCall(call) if call.name.original == "ExternName_IXTI1_N2"
        ),
        "expected FunctionCall(ExternName_IXTI1_N2), got {ast:?}"
    );
}

#[test]
fn decodes_value_class_ptg_namex_as_udf_function_name_via_sentinel_funcvar() {
    // Same pattern as `decodes_ptg_namex_as_udf_function_name_via_sentinel_funcvar`, but using the
    // value-class `PtgNameX` variant (0x59). The decoder should preserve the implicit intersection
    // marker.
    let mut rgce = Vec::new();
    rgce.extend_from_slice(&ptg_int(1));
    rgce.extend_from_slice(&ptg_int(2));
    rgce.extend_from_slice(&ptg_namex(1, 2, 0x59));
    rgce.extend_from_slice(&ptg_funcvar_udf(3));

    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "@ExternName_IXTI1_N2(1,2)");
    assert_parseable(&text);
    let ast = parse(&text);
    assert!(
        matches!(
            &ast.expr,
            Expr::Unary(u)
                if u.op == UnaryOp::ImplicitIntersection
                    && matches!(&*u.expr, Expr::FunctionCall(call) if call.name.original == "ExternName_IXTI1_N2")
        ),
        "expected @FunctionCall(ExternName_IXTI1_N2), got {ast:?}"
    );
}
