use desktop::commands::IpcPivotFieldRef;
use formula_engine::pivot::PivotFieldRef;

#[test]
fn ipc_pivot_field_ref_parses_dax_measure_strings() {
    let ipc: IpcPivotFieldRef = serde_json::from_str("\"[Total Sales]\"").unwrap();
    let core: PivotFieldRef = ipc.into();
    assert_eq!(
        core,
        PivotFieldRef::DataModelMeasure("Total Sales".to_string())
    );
}

#[test]
fn ipc_pivot_field_ref_parses_dax_column_strings() {
    let ipc: IpcPivotFieldRef = serde_json::from_str("\"Sales[Amount]\"").unwrap();
    let core: PivotFieldRef = ipc.into();
    assert_eq!(
        core,
        PivotFieldRef::DataModelColumn {
            table: "Sales".to_string(),
            column: "Amount".to_string()
        }
    );
}

#[test]
fn ipc_pivot_field_ref_parses_escaped_bracket_column_strings() {
    let ipc: IpcPivotFieldRef = serde_json::from_str("\"T[A]]B]\"").unwrap();
    let core: PivotFieldRef = ipc.into();
    assert_eq!(
        core,
        PivotFieldRef::DataModelColumn {
            table: "T".to_string(),
            column: "A]B".to_string()
        }
    );
}

#[test]
fn ipc_pivot_field_ref_parses_quoted_dax_column_strings() {
    let ipc: IpcPivotFieldRef = serde_json::from_str("\"'Sales Table'[Amount]\"").unwrap();
    let core: PivotFieldRef = ipc.into();
    assert_eq!(
        core,
        PivotFieldRef::DataModelColumn {
            table: "Sales Table".to_string(),
            column: "Amount".to_string()
        }
    );
}

#[test]
fn ipc_pivot_field_ref_parses_escaped_quote_in_quoted_table_name() {
    let ipc: IpcPivotFieldRef = serde_json::from_str("\"'O''Reilly'[Name]\"").unwrap();
    let core: PivotFieldRef = ipc.into();
    assert_eq!(
        core,
        PivotFieldRef::DataModelColumn {
            table: "O'Reilly".to_string(),
            column: "Name".to_string()
        }
    );
}

#[test]
fn ipc_pivot_field_ref_parses_escaped_bracket_measure_strings() {
    let ipc: IpcPivotFieldRef = serde_json::from_str("\"[A]]B]\"").unwrap();
    let core: PivotFieldRef = ipc.into();
    assert_eq!(core, PivotFieldRef::DataModelMeasure("A]B".to_string()));
}

#[test]
fn ipc_pivot_field_ref_parses_structured_dax_column_object() {
    let ipc: IpcPivotFieldRef = serde_json::from_str(r#"{"table":"Sales","column":"Amount"}"#).unwrap();
    let core: PivotFieldRef = ipc.into();
    assert_eq!(
        core,
        PivotFieldRef::DataModelColumn {
            table: "Sales".to_string(),
            column: "Amount".to_string()
        }
    );
}

#[test]
fn ipc_pivot_field_ref_preserves_brackets_in_structured_column_object() {
    // When the IPC payload is structured, we should treat the values as raw identifiers (not
    // DAX-escaped strings).
    let ipc: IpcPivotFieldRef = serde_json::from_str(r#"{"table":"T","column":"A]B"}"#).unwrap();
    let core: PivotFieldRef = ipc.into();
    assert_eq!(
        core,
        PivotFieldRef::DataModelColumn {
            table: "T".to_string(),
            column: "A]B".to_string()
        }
    );
}

#[test]
fn ipc_pivot_field_ref_parses_structured_measure_object() {
    let ipc: IpcPivotFieldRef = serde_json::from_str(r#"{"measure":"Total Sales"}"#).unwrap();
    let core: PivotFieldRef = ipc.into();
    assert_eq!(
        core,
        PivotFieldRef::DataModelMeasure("Total Sales".to_string())
    );
}

#[test]
fn ipc_pivot_field_ref_parses_structured_measure_name_object() {
    let ipc: IpcPivotFieldRef = serde_json::from_str(r#"{"name":"Total Sales"}"#).unwrap();
    let core: PivotFieldRef = ipc.into();
    assert_eq!(
        core,
        PivotFieldRef::DataModelMeasure("Total Sales".to_string())
    );
}

#[test]
fn ipc_pivot_field_ref_rejects_ambiguous_object_shapes() {
    // `{table,column}` and `{measure}` are mutually exclusive.
    let err = serde_json::from_str::<IpcPivotFieldRef>(r#"{"table":"T","measure":"X"}"#);
    assert!(err.is_err());
}

#[test]
fn ipc_pivot_field_ref_rejects_oversize_strings() {
    let max = desktop::resource_limits::MAX_PIVOT_TEXT_BYTES;
    let long = "a".repeat(max + 1);
    let json = serde_json::to_string(&long).unwrap();
    let err = serde_json::from_str::<IpcPivotFieldRef>(&json).unwrap_err();
    assert!(err.to_string().contains("string is too large"));
}

#[test]
fn ipc_pivot_field_ref_rejects_oversize_structured_table_names() {
    let max = desktop::resource_limits::MAX_PIVOT_TEXT_BYTES;
    let long = "a".repeat(max + 1);
    let json = format!(r#"{{"table":"{long}","column":"Amount"}}"#);
    let err = serde_json::from_str::<IpcPivotFieldRef>(&json).unwrap_err();
    assert!(err.to_string().contains("string is too large"));
}

#[test]
fn ipc_pivot_field_ref_leaves_non_dax_strings_as_cache_field_names() {
    let ipc: IpcPivotFieldRef = serde_json::from_str("\"Region\"").unwrap();
    let core: PivotFieldRef = ipc.into();
    assert_eq!(core, PivotFieldRef::CacheFieldName("Region".to_string()));
}
