use formula_xlsb::rgce::decode_rgce;
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

#[test]
fn rgce_roundtrip_cube_and_rtd_functions() {
    for formula in [
        r#"RTD("prog","server","topic")"#,
        r#"CUBEVALUE("conn","[Measures].[Sales]")"#,
        r#"CUBEMEMBER("conn","[Dim].[All]","Caption")"#,
        r#"CUBEMEMBERPROPERTY("conn","[Dim].[All]","PROP")"#,
        r#"CUBERANKEDMEMBER("conn","[SetExpr]",1,"Caption")"#,
        r#"CUBEKPIMEMBER("conn","KPI","KPIValue")"#,
        r#"CUBESET("conn","[SetExpr]","Caption",1,"[Measures].[Sales]")"#,
        r#"CUBESETCOUNT("[SetHandle]")"#,
    ] {
        let rgce = formula_biff::encode_rgce(formula).expect("encode");
        let decoded = decode_rgce(&rgce).expect("decode");
        assert_eq!(normalize(formula), normalize(&decoded));
    }
}

