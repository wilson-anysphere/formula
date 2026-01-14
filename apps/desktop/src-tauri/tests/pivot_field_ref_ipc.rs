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
fn ipc_pivot_field_ref_leaves_non_dax_strings_as_cache_field_names() {
    let ipc: IpcPivotFieldRef = serde_json::from_str("\"Region\"").unwrap();
    let core: PivotFieldRef = ipc.into();
    assert_eq!(core, PivotFieldRef::CacheFieldName("Region".to_string()));
}
