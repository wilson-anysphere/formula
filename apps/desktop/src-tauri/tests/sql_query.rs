use desktop::sql::{self, SqlDataType};
use desktop::ipc_limits;
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
async fn sqlite_path_memory_query_executes() {
    let result = sql::sql_query(
        json!({ "kind": "sqlite", "path": ":memory:" }),
        "SELECT CAST(1 AS INTEGER) AS one".to_string(),
        Vec::new(),
        None,
    )
    .await
    .expect("sqlite :memory: query should succeed");

    assert_eq!(result.columns, vec!["one".to_string()]);
    assert_eq!(result.rows, vec![vec![json!(1)]]);
}

#[tokio::test]
async fn sqlite_path_memory_get_schema_executes() {
    let result = sql::sql_get_schema(
        json!({ "kind": "sqlite", "path": ":memory:" }),
        "SELECT CAST(1 AS INTEGER) AS one".to_string(),
        None,
    )
    .await
    .expect("sqlite :memory: get_schema should succeed");

    assert_eq!(result.columns, vec!["one".to_string()]);
    let types = result.types.expect("expected type map");
    assert_eq!(types.get("one"), Some(&SqlDataType::Number));
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
async fn sql_query_rejects_oversized_sql_text() {
    let sql_text = "a".repeat(ipc_limits::MAX_SQL_QUERY_TEXT_BYTES + 1);
    let err = sql::sql_query(
        json!({ "kind": "sqlite", "inMemory": true }),
        sql_text,
        Vec::new(),
        None,
    )
    .await
    .expect_err("expected oversized sql text to error");

    let msg = err.to_string();
    assert!(
        msg.contains("MAX_SQL_QUERY_TEXT_BYTES"),
        "expected error to mention the limit, got: {msg}"
    );
}

#[tokio::test]
async fn sql_query_rejects_too_many_params() {
    let params = vec![json!(1); ipc_limits::MAX_SQL_QUERY_PARAMS + 1];
    let err = sql::sql_query(
        json!({ "kind": "sqlite", "inMemory": true }),
        "SELECT 1".to_string(),
        params,
        None,
    )
    .await
    .expect_err("expected too many params to error");

    let msg = err.to_string();
    assert!(
        msg.contains("MAX_SQL_QUERY_PARAMS"),
        "expected error to mention the limit, got: {msg}"
    );
}

#[tokio::test]
async fn sql_query_rejects_oversized_json_param() {
    // JSON string serialization adds 2 bytes for the surrounding quotes, so a string of
    // `MAX_SQL_QUERY_PARAM_BYTES` chars is guaranteed to exceed the limit.
    let oversized = json!("a".repeat(ipc_limits::MAX_SQL_QUERY_PARAM_BYTES));
    let err = sql::sql_query(
        json!({ "kind": "sqlite", "inMemory": true }),
        "SELECT 1".to_string(),
        vec![oversized],
        None,
    )
    .await
    .expect_err("expected oversized param to error");

    let msg = err.to_string();
    assert!(
        msg.contains("MAX_SQL_QUERY_PARAM_BYTES"),
        "expected error to mention the limit, got: {msg}"
    );
}

#[tokio::test]
async fn sql_query_rejects_oversized_connection_descriptor() {
    let connection = json!({
        "kind": "sqlite",
        "inMemory": true,
        "padding": "a".repeat(ipc_limits::MAX_SQL_QUERY_CONNECTION_BYTES)
    });
    let err = sql::sql_query(connection, "SELECT 1".to_string(), Vec::new(), None)
        .await
        .expect_err("expected oversized connection descriptor to error");

    let msg = err.to_string();
    assert!(
        msg.contains("MAX_SQL_QUERY_CONNECTION_BYTES"),
        "expected error to mention the limit, got: {msg}"
    );
}

#[tokio::test]
async fn sql_query_rejects_oversized_credentials() {
    let credentials = json!({
        "password": "a".repeat(ipc_limits::MAX_SQL_QUERY_CREDENTIALS_BYTES)
    });
    let err = sql::sql_query(
        json!({ "kind": "sqlite", "inMemory": true }),
        "SELECT 1".to_string(),
        Vec::new(),
        Some(credentials),
    )
    .await
    .expect_err("expected oversized credentials to error");

    let msg = err.to_string();
    assert!(
        msg.contains("MAX_SQL_QUERY_CREDENTIALS_BYTES"),
        "expected error to mention the limit, got: {msg}"
    );
}

#[tokio::test]
async fn sql_get_schema_rejects_oversized_sql_text() {
    let sql_text = "a".repeat(ipc_limits::MAX_SQL_QUERY_TEXT_BYTES + 1);
    let err = sql::sql_get_schema(
        json!({ "kind": "sqlite", "inMemory": true }),
        sql_text,
        None,
    )
    .await
    .expect_err("expected oversized sql text to error");

    let msg = err.to_string();
    assert!(
        msg.contains("MAX_SQL_QUERY_TEXT_BYTES"),
        "expected error to mention the limit, got: {msg}"
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
async fn odbc_sqlite_in_memory_query_accepts_memory_alias() {
    let result = sql::sql_query(
        json!({ "kind": "odbc", "connectionString": "Driver=SQLite3;Database=memory" }),
        "SELECT CAST(1 AS INTEGER) AS one".to_string(),
        Vec::new(),
        None,
    )
    .await
    .expect("odbc sqlite query should succeed with Database=memory alias");

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
        if let Ok(canon) = dunce::canonicalize(base.home_dir()) {
            roots.push(canon);
        }
    }

    if let Some(user) = directories::UserDirs::new() {
        if let Some(doc) = user.document_dir() {
            if let Ok(canon) = dunce::canonicalize(doc) {
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
    let outside_dir_canon = dunce::canonicalize(outside_dir.path()).expect("canonicalize");

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
    let outside_dir_canon = dunce::canonicalize(outside_dir.path()).expect("canonicalize");
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

#[cfg(unix)]
#[tokio::test]
async fn sqlite_dotdot_traversal_escape_is_denied() {
    use std::path::Component;

    let roots = desktop_allowed_roots();
    assert!(!roots.is_empty(), "expected at least one allowed root");

    // Create a real database file outside the allowed roots.
    let outside_dir = tempfile::tempdir().expect("tempdir");
    let outside_dir_canon = dunce::canonicalize(outside_dir.path()).expect("canonicalize");
    if is_in_roots(&outside_dir_canon, &roots) {
        eprintln!(
            "skipping .. traversal sqlite test: tempdir {} is within allowed roots {roots:?}",
            outside_dir_canon.display()
        );
        return;
    }
    let outside_db = outside_dir.path().join("outside.sqlite");
    create_sqlite_db(&outside_db).await;

    // Craft a path that *appears* under the allowed root but uses `..` components to reach the
    // outside db path.
    let allowed_root = roots[0].clone();
    let mut escape_path = allowed_root.clone();
    let mut depth = 0usize;
    for comp in allowed_root.components() {
        if matches!(comp, Component::Normal(_)) {
            depth += 1;
        }
    }
    for _ in 0..depth {
        escape_path.push("..");
    }
    for comp in outside_db.components() {
        if matches!(comp, Component::Normal(_)) {
            escape_path.push(comp.as_os_str());
        }
    }

    let err = sql::sql_query(
        json!({ "kind": "sqlite", "path": escape_path }),
        "SELECT 1".to_string(),
        Vec::new(),
        None,
    )
    .await
    .expect_err("expected .. traversal escape to error");

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
    let outside_dir_canon = dunce::canonicalize(outside_dir.path()).expect("canonicalize");

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

#[tokio::test]
async fn odbc_sqlite_file_path_within_scope_is_allowed() {
    let roots = desktop_allowed_roots();
    assert!(!roots.is_empty(), "expected at least one allowed root");

    let allowed_dir = tempfile::tempdir_in(&roots[0]).expect("tempdir in allowed root");
    let db_path = allowed_dir.path().join("allowed.sqlite");
    create_sqlite_db(&db_path).await;

    let conn_str = format!("Driver={{SQLite3}};Database={};", db_path.display());
    let result = sql::sql_query(
        json!({ "kind": "odbc", "connectionString": conn_str }),
        "SELECT CAST(1 AS INTEGER) AS one".to_string(),
        Vec::new(),
        None,
    )
    .await
    .expect("odbc sqlite query should succeed for scoped db path");

    assert_eq!(result.columns, vec!["one".to_string()]);
    assert_eq!(result.rows, vec![vec![json!(1)]]);
}

#[cfg(unix)]
#[tokio::test]
async fn odbc_sqlite_symlink_escape_is_denied() {
    use std::os::unix::fs::symlink;

    let roots = desktop_allowed_roots();
    assert!(!roots.is_empty(), "expected at least one allowed root");

    // Create a real database file outside the allowed roots.
    let outside_dir = tempfile::tempdir().expect("tempdir");
    let outside_dir_canon = dunce::canonicalize(outside_dir.path()).expect("canonicalize");
    if is_in_roots(&outside_dir_canon, &roots) {
        eprintln!(
            "skipping symlink escape odbc sqlite test: tempdir {} is within allowed roots {roots:?}",
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

    let conn_str = format!("Driver=SQLite3;Database={};", link_path.display());
    let err = sql::sql_query(
        json!({ "kind": "odbc", "connectionString": conn_str }),
        "SELECT 1".to_string(),
        Vec::new(),
        None,
    )
    .await
    .expect_err("expected odbc sqlite symlink escape to error");

    assert!(
        err.to_string().contains("Access denied: SQLite database path is outside the allowed filesystem scope"),
        "unexpected error: {err}"
    );
}

#[cfg(unix)]
#[tokio::test]
async fn odbc_sqlite_dotdot_traversal_escape_is_denied() {
    use std::path::Component;

    let roots = desktop_allowed_roots();
    assert!(!roots.is_empty(), "expected at least one allowed root");

    // Create a real database file outside the allowed roots.
    let outside_dir = tempfile::tempdir().expect("tempdir");
    let outside_dir_canon = dunce::canonicalize(outside_dir.path()).expect("canonicalize");
    if is_in_roots(&outside_dir_canon, &roots) {
        eprintln!(
            "skipping .. traversal odbc sqlite test: tempdir {} is within allowed roots {roots:?}",
            outside_dir_canon.display()
        );
        return;
    }
    let outside_db = outside_dir.path().join("outside.sqlite");
    create_sqlite_db(&outside_db).await;

    // Craft a path that *appears* under the allowed root but uses `..` components to reach the
    // outside db path.
    let allowed_root = roots[0].clone();
    let mut escape_path = allowed_root.clone();
    let mut depth = 0usize;
    for comp in allowed_root.components() {
        if matches!(comp, Component::Normal(_)) {
            depth += 1;
        }
    }
    for _ in 0..depth {
        escape_path.push("..");
    }
    for comp in outside_db.components() {
        if matches!(comp, Component::Normal(_)) {
            escape_path.push(comp.as_os_str());
        }
    }

    let conn_str = format!("Driver=SQLite3;Database={};", escape_path.display());
    let err = sql::sql_query(
        json!({ "kind": "odbc", "connectionString": conn_str }),
        "SELECT 1".to_string(),
        Vec::new(),
        None,
    )
    .await
    .expect_err("expected odbc sqlite .. traversal escape to error");

    assert!(
        err.to_string().contains("Access denied: SQLite database path is outside the allowed filesystem scope"),
        "unexpected error: {err}"
    );
}

#[tokio::test]
async fn sqlite_attach_cannot_read_out_of_scope_db_via_multi_statement() {
    let roots = desktop_allowed_roots();
    assert!(!roots.is_empty(), "expected at least one allowed root");

    // Create a real database file outside the allowed roots.
    let outside_dir = tempfile::tempdir().expect("tempdir");
    let outside_dir_canon = dunce::canonicalize(outside_dir.path()).expect("canonicalize");

    if is_in_roots(&outside_dir_canon, &roots) {
        eprintln!(
            "skipping sqlite attach test: tempdir {} is within allowed roots {roots:?}",
            outside_dir_canon.display()
        );
        return;
    }

    let outside_db = outside_dir.path().join("outside.sqlite");
    create_sqlite_db(&outside_db).await;

    // Insert a sentinel value so a successful attach+select would exfiltrate it.
    let opts = sqlx::sqlite::SqliteConnectOptions::new()
        .filename(&outside_db)
        .create_if_missing(false);
    let mut conn = sqlx::SqliteConnection::connect_with(&opts)
        .await
        .expect("connect outside db");
    conn.execute("CREATE TABLE IF NOT EXISTS t (v INTEGER);")
        .await
        .expect("create table");
    conn.execute("DELETE FROM t;").await.expect("clear table");
    conn.execute("INSERT INTO t (v) VALUES (42);")
        .await
        .expect("insert sentinel");

    // Attempt to attach + query the out-of-scope db in a single SQL payload. This should not be
    // able to return the sentinel value.
    let db_path_sql = outside_db.to_string_lossy().replace('\'', "''");
    let sql_text = format!("ATTACH DATABASE '{db_path_sql}' AS ext; SELECT v FROM ext.t;");

    match sql::sql_query(
        json!({ "kind": "sqlite", "inMemory": true }),
        sql_text,
        Vec::new(),
        None,
    )
    .await
    {
        Ok(result) => {
            let leaked = result
                .rows
                .iter()
                .flat_map(|row| row.iter())
                .any(|v| *v == json!(42));
            assert!(
                !leaked,
                "expected ATTACH+SELECT to not return data from an out-of-scope sqlite db"
            );
        }
        Err(_) => {
            // Accept errors here; the important security property is that we don't successfully
            // return data from the out-of-scope database.
        }
    }
}

#[tokio::test]
async fn sqlite_get_schema_file_path_within_scope_is_allowed() {
    let roots = desktop_allowed_roots();
    assert!(!roots.is_empty(), "expected at least one allowed root");

    let allowed_dir = tempfile::tempdir_in(&roots[0]).expect("tempdir in allowed root");
    let db_path = allowed_dir.path().join("allowed.sqlite");
    create_sqlite_db(&db_path).await;

    let result = sql::sql_get_schema(
        json!({ "kind": "sqlite", "path": db_path }),
        "SELECT CAST(1 AS INTEGER) AS one".to_string(),
        None,
    )
    .await
    .expect("sql_get_schema should succeed for scoped db path");

    assert_eq!(result.columns, vec!["one".to_string()]);
    let types = result.types.expect("expected type map");
    assert_eq!(types.get("one"), Some(&SqlDataType::Number));
}

#[tokio::test]
async fn sqlite_get_schema_file_path_outside_scope_is_denied() {
    let roots = desktop_allowed_roots();
    assert!(!roots.is_empty(), "expected at least one allowed root");

    let outside_dir = tempfile::tempdir().expect("tempdir");
    let outside_dir_canon = dunce::canonicalize(outside_dir.path()).expect("canonicalize");

    if is_in_roots(&outside_dir_canon, &roots) {
        eprintln!(
            "skipping out-of-scope sqlite get_schema test: tempdir {} is within allowed roots {roots:?}",
            outside_dir_canon.display()
        );
        return;
    }

    let db_path = outside_dir.path().join("outside.sqlite");
    create_sqlite_db(&db_path).await;

    let err = sql::sql_get_schema(
        json!({ "kind": "sqlite", "path": db_path }),
        "SELECT 1".to_string(),
        None,
    )
    .await
    .expect_err("expected sql_get_schema to deny out-of-scope sqlite path");

    assert!(
        err.to_string().contains("Access denied: SQLite database path is outside the allowed filesystem scope"),
        "unexpected error: {err}"
    );
}

#[tokio::test]
async fn odbc_sqlite_get_schema_file_path_outside_scope_is_denied() {
    let roots = desktop_allowed_roots();
    assert!(!roots.is_empty(), "expected at least one allowed root");

    let outside_dir = tempfile::tempdir().expect("tempdir");
    let outside_dir_canon = dunce::canonicalize(outside_dir.path()).expect("canonicalize");

    if is_in_roots(&outside_dir_canon, &roots) {
        eprintln!(
            "skipping out-of-scope ODBC sqlite get_schema test: tempdir {} is within allowed roots {roots:?}",
            outside_dir_canon.display()
        );
        return;
    }

    let db_path = outside_dir.path().join("outside.sqlite");
    create_sqlite_db(&db_path).await;

    let conn_str = format!("Driver=SQLite3;Database={};", db_path.display());
    let err = sql::sql_get_schema(
        json!({ "kind": "odbc", "connectionString": conn_str }),
        "SELECT 1".to_string(),
        None,
    )
    .await
    .expect_err("expected sql_get_schema to deny out-of-scope odbc sqlite path");

    assert!(
        err.to_string().contains("Access denied: SQLite database path is outside the allowed filesystem scope"),
        "unexpected error: {err}"
    );
}

#[tokio::test]
async fn odbc_sqlite_get_schema_file_path_within_scope_is_allowed() {
    let roots = desktop_allowed_roots();
    assert!(!roots.is_empty(), "expected at least one allowed root");

    let allowed_dir = tempfile::tempdir_in(&roots[0]).expect("tempdir in allowed root");
    let db_path = allowed_dir.path().join("allowed.sqlite");
    create_sqlite_db(&db_path).await;

    let conn_str = format!("Driver={{SQLite3}};Database={};", db_path.display());
    let result = sql::sql_get_schema(
        json!({ "kind": "odbc", "connectionString": conn_str }),
        "SELECT CAST(1 AS INTEGER) AS one".to_string(),
        None,
    )
    .await
    .expect("odbc sqlite get_schema should succeed for scoped db path");

    assert_eq!(result.columns, vec!["one".to_string()]);
    let types = result.types.expect("expected type map");
    assert_eq!(types.get("one"), Some(&SqlDataType::Number));
}
