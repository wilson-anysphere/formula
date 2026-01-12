use desktop::sql::{self, SqlDataType};
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
        json!({ "kind": "mysql", "url": "mysql://localhost" }),
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

#[tokio::test]
async fn sqlserver_connections_return_clear_error() {
    let err = sql::sql_query(
        json!({ "kind": "sql", "server": "localhost", "database": "db" }),
        "SELECT 1".to_string(),
        Vec::new(),
        Some(json!({ "user": "sa", "password": "pw" })),
    )
    .await
    .expect_err("expected sqlserver kind to error");

    assert!(
        err.to_string().contains("SQL Server connections are not supported"),
        "unexpected error: {err}"
    );
}

#[tokio::test]
async fn odbc_sqlite_in_memory_query_executes() {
    let result = sql::sql_query(
        json!({ "kind": "odbc", "connectionString": "Driver=SQLite3;Database=:memory:" }),
        "SELECT CAST(1 AS INTEGER) AS one".to_string(),
        Vec::new(),
        None,
    )
    .await
    .expect("odbc sqlite query should succeed");

    assert_eq!(result.columns, vec!["one".to_string()]);
    assert_eq!(result.rows, vec![vec![json!(1)]]);
}

#[tokio::test]
async fn odbc_with_unsupported_driver_returns_clear_error() {
    let err = sql::sql_query(
        json!({ "kind": "odbc", "connectionString": "Driver={SQL Server};Server=localhost;Database=db;" }),
        "SELECT 1".to_string(),
        Vec::new(),
        None,
    )
    .await
    .expect_err("expected unsupported ODBC driver to error");

    assert!(
        err.to_string().contains("Unsupported ODBC driver"),
        "unexpected error: {err}"
    );
}

#[tokio::test]
async fn odbc_dsn_connections_return_clear_error() {
    let err = sql::sql_query(
        json!({ "kind": "odbc", "connectionString": "dsn=mydb" }),
        "SELECT 1".to_string(),
        Vec::new(),
        None,
    )
    .await
    .expect_err("expected DSN-only connection to error");

    assert!(
        err.to_string().to_ascii_lowercase().contains("dsn") && err.to_string().contains("not supported"),
        "unexpected error: {err}"
    );
}
