use formula_desktop_tauri::sql::{self, SqlDataType};
use serde_json::json;

#[tokio::test]
async fn sqlite_in_memory_query_returns_rows_and_columns() {
    let result = sql::sql_query(
        json!({ "kind": "sqlite", "inMemory": true }),
        "SELECT CAST(1 AS INTEGER) AS one, CAST('two' AS TEXT) AS two".to_string(),
        Vec::new(),
        None,
    )
    .await
    .expect("sql query should succeed");

    assert_eq!(result.columns, vec!["one".to_string(), "two".to_string()]);
    assert_eq!(result.rows, vec![vec![json!(1), json!("two")]]);

    let types = result.types.expect("expected type map");
    assert_eq!(types.get("one"), Some(&SqlDataType::Number));
    assert_eq!(types.get("two"), Some(&SqlDataType::String));
}

#[tokio::test]
async fn unsupported_connection_kind_returns_clear_error() {
    let err = sql::sql_query(
        json!({ "kind": "odbc", "connectionString": "Driver={PostgreSQL};Server=localhost;" }),
        "SELECT 1".to_string(),
        Vec::new(),
        None,
    )
    .await
    .expect_err("expected unsupported kind to error");

    assert!(
        err.to_string().contains("Unsupported SQL connection kind"),
        "unexpected error: {err}"
    );
}
