use desktop::sql::{self, SqlDataType};
use sqlx::Connection;
use sqlx::Executor;
use serde_json::json;
use std::path::{Path, PathBuf};

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

async fn create_sqlite_db(path: &Path) {
    let opts = sqlx::sqlite::SqliteConnectOptions::new()
        .filename(path)
        .create_if_missing(true);
    let mut conn = sqlx::SqliteConnection::connect_with(&opts)
        .await
        .expect("create sqlite db");
    // Ensure a well-formed sqlite file exists.
    conn.execute("PRAGMA user_version = 1;")
        .await
        .expect("write pragma");
}

fn desktop_allowed_roots() -> Vec<PathBuf> {
    let mut roots = Vec::new();

    if let Some(base) = directories::BaseDirs::new() {
        if let Ok(canon) = std::fs::canonicalize(base.home_dir()) {
            roots.push(canon);
        }
    }

    if let Some(user) = directories::UserDirs::new() {
        if let Some(doc) = user.document_dir() {
            if let Ok(canon) = std::fs::canonicalize(doc) {
                if !roots.iter().any(|p| p == &canon) {
                    roots.push(canon);
                }
            }
        }
    }

    roots
}

fn is_in_roots(path: &Path, roots: &[PathBuf]) -> bool {
    roots.iter().any(|root| path.starts_with(root))
}

#[tokio::test]
async fn sqlite_file_path_within_scope_is_allowed() {
    let roots = desktop_allowed_roots();
    assert!(!roots.is_empty(), "expected at least one allowed root");

    let home_root = roots[0].clone();
    let allowed_dir = tempfile::tempdir_in(&home_root).expect("tempdir in allowed root");
    let db_path = allowed_dir.path().join("allowed.sqlite");
    create_sqlite_db(&db_path).await;

    let result = sql::sql_query(
        json!({ "kind": "sqlite", "path": db_path }),
        "SELECT CAST(1 AS INTEGER) AS one".to_string(),
        Vec::new(),
        None,
    )
    .await
    .expect("sqlite query should succeed for scoped db path");

    assert_eq!(result.columns, vec!["one".to_string()]);
    assert_eq!(result.rows, vec![vec![json!(1)]]);
}

#[tokio::test]
async fn sqlite_file_path_outside_scope_is_denied() {
    let roots = desktop_allowed_roots();
    assert!(!roots.is_empty(), "expected at least one allowed root");

    let outside_dir = tempfile::tempdir().expect("tempdir");
    let outside_dir_canon = std::fs::canonicalize(outside_dir.path()).expect("canonicalize");

    if is_in_roots(&outside_dir_canon, &roots) {
        // Extremely unusual environment (e.g. HOME=/tmp). Skip rather than producing a false
        // negative.
        eprintln!(
            "skipping out-of-scope sqlite test: tempdir {} is within allowed roots {roots:?}",
            outside_dir_canon.display()
        );
        return;
    }

    let db_path = outside_dir.path().join("outside.sqlite");
    create_sqlite_db(&db_path).await;

    let err = sql::sql_query(
        json!({ "kind": "sqlite", "path": db_path }),
        "SELECT 1".to_string(),
        Vec::new(),
        None,
    )
    .await
    .expect_err("expected sqlite db path outside scope to error");

    assert!(
        err.to_string().contains("Access denied: SQLite database path is outside the allowed filesystem scope"),
        "unexpected error: {err}"
    );
}

#[cfg(unix)]
#[tokio::test]
async fn sqlite_symlink_escape_is_denied() {
    use std::os::unix::fs::symlink;

    let roots = desktop_allowed_roots();
    assert!(!roots.is_empty(), "expected at least one allowed root");

    // Create a real database file outside the allowed roots.
    let outside_dir = tempfile::tempdir().expect("tempdir");
    let outside_dir_canon = std::fs::canonicalize(outside_dir.path()).expect("canonicalize");
    if is_in_roots(&outside_dir_canon, &roots) {
        eprintln!(
            "skipping symlink escape sqlite test: tempdir {} is within allowed roots {roots:?}",
            outside_dir_canon.display()
        );
        return;
    }
    let outside_db = outside_dir.path().join("outside.sqlite");
    create_sqlite_db(&outside_db).await;

    // Create a symlink inside an allowed root pointing to the outside database.
    let allowed_dir = tempfile::tempdir_in(&roots[0]).expect("tempdir in allowed root");
    let link_path = allowed_dir.path().join("db.sqlite");
    symlink(&outside_db, &link_path).expect("create symlink");

    let err = sql::sql_query(
        json!({ "kind": "sqlite", "path": link_path }),
        "SELECT 1".to_string(),
        Vec::new(),
        None,
    )
    .await
    .expect_err("expected symlink escape to error");

    assert!(
        err.to_string().contains("Access denied: SQLite database path is outside the allowed filesystem scope"),
        "unexpected error: {err}"
    );
}

#[tokio::test]
async fn odbc_sqlite_file_path_outside_scope_is_denied() {
    let roots = desktop_allowed_roots();
    assert!(!roots.is_empty(), "expected at least one allowed root");

    let outside_dir = tempfile::tempdir().expect("tempdir");
    let outside_dir_canon = std::fs::canonicalize(outside_dir.path()).expect("canonicalize");

    if is_in_roots(&outside_dir_canon, &roots) {
        eprintln!(
            "skipping out-of-scope ODBC sqlite test: tempdir {} is within allowed roots {roots:?}",
            outside_dir_canon.display()
        );
        return;
    }

    let db_path = outside_dir.path().join("outside.sqlite");
    create_sqlite_db(&db_path).await;

    let conn_str = format!("Driver=SQLite3;Database={};", db_path.display());
    let err = sql::sql_query(
        json!({ "kind": "odbc", "connectionString": conn_str }),
        "SELECT 1".to_string(),
        Vec::new(),
        None,
    )
    .await
    .expect_err("expected odbc sqlite db path outside scope to error");

    assert!(
        err.to_string().contains("Access denied: SQLite database path is outside the allowed filesystem scope"),
        "unexpected error: {err}"
    );
}
