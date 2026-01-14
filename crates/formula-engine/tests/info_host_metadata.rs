use formula_engine::{Engine, EngineInfo, ErrorKind, Value};

#[test]
fn info_system_updates_when_host_metadata_changes() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", r#"=INFO("system")"#)
        .unwrap();

    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Text("pcdos".to_string())
    );

    let mut info = EngineInfo::default();
    info.system = Some("mac".to_string());
    engine.set_engine_info(info);
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Text("mac".to_string())
    );

    engine.set_engine_info(EngineInfo::default());
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Text("pcdos".to_string())
    );
}

#[test]
fn info_os_version_release_version_roundtrip() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", r#"=INFO("osversion")"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", r#"=INFO("release")"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A3", r#"=INFO("version")"#)
        .unwrap();

    engine.recalculate_single_threaded();
    for addr in ["A1", "A2", "A3"] {
        assert_eq!(
            engine.get_cell_value("Sheet1", addr),
            Value::Error(ErrorKind::NA)
        );
    }

    let mut info = EngineInfo::default();
    info.osversion = Some("os".to_string());
    info.release = Some("release".to_string());
    info.version = Some("version".to_string());
    engine.set_engine_info(info);

    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Text("os".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A2"),
        Value::Text("release".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A3"),
        Value::Text("version".to_string())
    );
}

#[test]
fn info_mem_roundtrips() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", r#"=INFO("memavail")"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", r#"=INFO("totmem")"#)
        .unwrap();

    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(ErrorKind::NA)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A2"),
        Value::Error(ErrorKind::NA)
    );

    let mut info = EngineInfo::default();
    info.memavail = Some(123.0);
    info.totmem = Some(456.0);
    engine.set_engine_info(info);
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(123.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(456.0));
}

#[test]
fn info_origin_is_per_sheet_and_absolute() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", r#"=INFO("origin")"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet2", "A1", r#"=INFO("origin")"#)
        .unwrap();

    engine.set_sheet_origin("Sheet1", Some("$A$1")).unwrap();
    engine.set_sheet_origin("Sheet2", Some("$B$2")).unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Text("$A$1".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet2", "A1"),
        Value::Text("$B$2".to_string())
    );
}
