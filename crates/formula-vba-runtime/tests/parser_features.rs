use formula_vba_runtime::{
    parse_program, ArrayDim, CaseComparisonOp, CaseCondition, Expr, LoopConditionKind, Stmt, VbaType,
};

#[test]
fn parses_option_explicit_and_module_level_decls() {
    let code = r#"
Option Explicit
Public Const Greeting As String = "Hi"
Private counter As Long

Sub Test()
End Sub
"#;
    let program = parse_program(code).expect("program parses");
    assert!(program.option_explicit);
    assert_eq!(program.module_consts.len(), 1);
    assert_eq!(program.module_vars.len(), 1);

    assert_eq!(program.module_consts[0].name, "Greeting");
    assert_eq!(program.module_consts[0].ty, Some(VbaType::String));
    match &program.module_consts[0].value {
        Expr::Literal(v) => assert_eq!(v.to_string_lossy(), "Hi"),
        other => panic!("expected literal const, got {other:?}"),
    }

    assert_eq!(program.module_vars[0].name, "counter");
    assert_eq!(program.module_vars[0].ty, VbaType::Long);
}

#[test]
fn parses_typed_dim_and_arrays() {
    let code = r#"
Sub Test()
    Dim x As Integer, s As String, d As Date
    Dim a(1 To 3) As Long
End Sub
"#;

    let program = parse_program(code).expect("program parses");
    let proc = program.get("test").expect("procedure parsed");
    assert_eq!(proc.body.len(), 2);

    let Stmt::Dim(vars) = &proc.body[0] else {
        panic!("expected Dim, got {:?}", proc.body[0]);
    };
    assert_eq!(vars.len(), 3);
    assert_eq!(vars[0].name, "x");
    assert_eq!(vars[0].ty, VbaType::Integer);
    assert_eq!(vars[1].name, "s");
    assert_eq!(vars[1].ty, VbaType::String);
    assert_eq!(vars[2].name, "d");
    assert_eq!(vars[2].ty, VbaType::Date);

    let Stmt::Dim(arrs) = &proc.body[1] else {
        panic!("expected Dim, got {:?}", proc.body[1]);
    };
    assert_eq!(arrs.len(), 1);
    assert_eq!(arrs[0].name, "a");
    assert_eq!(arrs[0].ty, VbaType::Long);
    assert_eq!(arrs[0].dims.len(), 1);
    let ArrayDim { lower, upper } = &arrs[0].dims[0];
    assert!(matches!(lower, Some(Expr::Literal(_))));
    assert!(matches!(upper, Expr::Literal(_)));
}

#[test]
fn parses_for_each_do_until_while_wend_select_case_with_named_args_and_resume() {
    let code = r#"
Sub Test()
    Dim c
    Set c = New Collection
    c.Add 1: c.Add 2

    Dim v
    For Each v In c
        Debug.Print v
    Next v

    Dim i As Long
    i = 0
    Do Until i = 3
        i = i + 1
    Loop

    While i < 10
        i = i + 1
    Wend

    Select Case i
        Case 10
            i = 11
        Case 11 To 12
            i = 13
        Case Is >= 14
            i = 15
        Case Else
            i = 0
    End Select

    With Range("A1")
        .Value = "X"
        .Offset(, 1).Value = .Value & "Y"
        .AutoFill Destination:=Range("A1:A3")
    End With

    On Error GoTo ErrHandler
    Range().Value = 1
    Exit Sub
ErrHandler:
    Resume Next
End Sub
"#;

    let program = parse_program(code).expect("program parses");
    let proc = program.get("test").expect("procedure parsed");

    assert!(
        proc.body.iter().any(|s| matches!(s, Stmt::ForEach { .. })),
        "expected ForEach in body: {:#?}",
        proc.body
    );
    assert!(
        proc.body.iter().any(|s| matches!(s, Stmt::DoLoop { pre_condition: Some((LoopConditionKind::Until, _)), .. })),
        "expected Do Until in body: {:#?}",
        proc.body
    );
    assert!(
        proc.body.iter().any(|s| matches!(s, Stmt::While { .. })),
        "expected While/Wend in body: {:#?}",
        proc.body
    );

    let select = proc
        .body
        .iter()
        .find(|s| matches!(s, Stmt::SelectCase { .. }))
        .expect("select case present");
    let Stmt::SelectCase { cases, else_body, .. } = select else {
        unreachable!();
    };
    assert!(cases.iter().any(|arm| arm.conditions.iter().any(|c| matches!(c, CaseCondition::Range { .. }))));
    assert!(cases.iter().any(|arm| arm.conditions.iter().any(|c| matches!(c, CaseCondition::Is { op: CaseComparisonOp::Ge, .. }))));
    assert!(!else_body.is_empty());

    // Ensure `Offset(, 1)` parses a missing first arg.
    let with_stmt = proc
        .body
        .iter()
        .find(|s| matches!(s, Stmt::With { .. }))
        .expect("with present");
    let Stmt::With { body, .. } = with_stmt else { unreachable!() };
    fn is_offset_assignment(stmt: &Stmt) -> bool {
        let Stmt::Assign { target, .. } = stmt else {
            return false;
        };
        let Expr::Member { object, .. } = target else {
            return false;
        };
        let Expr::Call { callee, .. } = &**object else {
            return false;
        };
        let Expr::Member { member, .. } = &**callee else {
            return false;
        };
        member.eq_ignore_ascii_case("offset")
    }
    let offset_assign = body
        .iter()
        .find(|s| is_offset_assignment(s))
        .expect("Offset assignment inside with");
    let Stmt::Assign { target, .. } = offset_assign else { unreachable!() };
    let Expr::Member { object, .. } = target else { panic!("expected member assign") };
    let Expr::Call { callee, args } = &**object else { panic!("expected call on offset") };
    let Expr::Member { member, .. } = &**callee else { panic!("expected .Offset member") };
    assert_eq!(member.to_ascii_lowercase(), "offset");
    assert!(matches!(args[0].expr, Expr::Missing));
    assert!(!matches!(args[1].expr, Expr::Missing));
}
