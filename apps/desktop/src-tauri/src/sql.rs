use anyhow::{anyhow, Context, Result};
use futures_util::TryStreamExt;
use serde::{Deserialize, Serialize};
use serde_json::Value as JsonValue;
use sqlx::{Column, Connection, Executor, Row, TypeInfo};
use std::collections::HashMap;
use std::path::{Path, PathBuf};
use std::str::FromStr;
use std::time::Duration;

use crate::ipc_limits::{
    enforce_json_byte_size, MAX_SQL_QUERY_CONNECTION_BYTES, MAX_SQL_QUERY_CREDENTIALS_BYTES,
    MAX_SQL_QUERY_PARAM_BYTES, MAX_SQL_QUERY_PARAMS, MAX_SQL_QUERY_TEXT_BYTES,
};

// Resource limits for `sql_query`.
//
// These are backend-enforced guards to prevent compromised/buggy webviews (or accidental queries
// like `SELECT * FROM huge_table`) from consuming unbounded memory/CPU in the Rust process.
pub const MAX_SQL_QUERY_ROWS: usize = 50_000;
pub const MAX_SQL_QUERY_CELLS: usize = 5_000_000;
/// Approximate total payload size guard (sum of string lengths + small overhead for scalars).
///
/// This is a secondary backstop against unexpectedly large cells (e.g. very long TEXT columns)
/// even when row/cell counts are within bounds.
pub const MAX_SQL_QUERY_BYTES: usize = 50 * 1024 * 1024; // 50 MiB
/// Maximum size for any individual cell value (primarily to prevent allocating a single huge TEXT
/// field into memory).
pub const MAX_SQL_QUERY_CELL_BYTES: usize = 10 * 1024 * 1024; // 10 MiB
pub const SQL_QUERY_TIMEOUT_MS: u64 = 10_000;

fn sql_query_timeout_ms() -> u64 {
    let default = SQL_QUERY_TIMEOUT_MS;
    let Some(value) = std::env::var("FORMULA_SQL_QUERY_TIMEOUT_MS")
        .ok()
        .and_then(|v| v.parse::<u64>().ok())
    else {
        return default;
    };
    // Allow relaxing limits only in debug builds; in release we only honor stricter settings.
    if cfg!(debug_assertions) {
        value
    } else {
        value.min(default)
    }
}

fn sql_query_max_bytes() -> usize {
    let default = MAX_SQL_QUERY_BYTES;
    let Some(value) = std::env::var("FORMULA_SQL_QUERY_MAX_BYTES")
        .ok()
        .and_then(|v| v.parse::<usize>().ok())
    else {
        return default;
    };
    if cfg!(debug_assertions) {
        value
    } else {
        value.min(default)
    }
}

fn sql_query_max_cell_bytes() -> usize {
    let default = MAX_SQL_QUERY_CELL_BYTES;
    let Some(value) = std::env::var("FORMULA_SQL_QUERY_MAX_CELL_BYTES")
        .ok()
        .and_then(|v| v.parse::<usize>().ok())
    else {
        return default;
    };
    if cfg!(debug_assertions) {
        value
    } else {
        value.min(default)
    }
}

fn sql_query_timeout_duration() -> Duration {
    Duration::from_millis(sql_query_timeout_ms())
}

fn sql_timeout_error(kind: &'static str) -> anyhow::Error {
    let timeout_ms = sql_query_timeout_ms();
    anyhow!(
        "{kind} query exceeded the maximum execution time ({timeout_ms}ms). \
         Try adding a LIMIT clause or optimizing the query."
    )
}

fn sql_value_estimated_bytes(value: &JsonValue) -> usize {
    match value {
        JsonValue::Null => 4,
        JsonValue::Bool(true) => 4,
        JsonValue::Bool(false) => 5,
        JsonValue::Number(_) => 16,
        JsonValue::String(s) => s.len(),
        JsonValue::Array(arr) => arr.iter().map(sql_value_estimated_bytes).sum(),
        JsonValue::Object(obj) => obj
            .iter()
            .map(|(k, v)| k.len() + sql_value_estimated_bytes(v))
            .sum(),
    }
}

async fn with_sql_query_timeout<T>(
    kind: &'static str,
    fut: impl std::future::Future<Output = Result<T>>,
) -> Result<T> {
    tokio::time::timeout(sql_query_timeout_duration(), fut)
        .await
        .map_err(|_| sql_timeout_error(kind))?
}

const SQLITE_SCOPE_DENIED_ERROR: &str =
    "Access denied: SQLite database path is outside the allowed filesystem scope";

fn validate_sqlite_db_path(path: &str) -> Result<PathBuf> {
    let raw = Path::new(path);
    let allowed_roots =
        crate::fs_scope::desktop_allowed_roots().context("determine allowed filesystem scope roots")?;

    match crate::fs_scope::canonicalize_in_allowed_roots_with_error(raw, &allowed_roots) {
        Ok(canonical) => Ok(canonical),
        Err(crate::fs_scope::CanonicalizeInAllowedRootsError::OutsideScope { .. }) => {
            Err(anyhow!(SQLITE_SCOPE_DENIED_ERROR))
        }
        Err(crate::fs_scope::CanonicalizeInAllowedRootsError::Canonicalize { source, .. }) => {
            Err(anyhow::Error::new(source)).context("canonicalize sqlite database path")
        }
    }
}

fn sqlite_connect_options_from_path(path: &str) -> Result<sqlx::sqlite::SqliteConnectOptions> {
    let trimmed = path.trim();
    if trimmed.eq_ignore_ascii_case(":memory:") || trimmed.eq_ignore_ascii_case("memory") {
        let mut opts = sqlx::sqlite::SqliteConnectOptions::from_str("sqlite::memory:")?;
        opts = opts.busy_timeout(sql_query_timeout_duration());
        return Ok(opts);
    }
    let canonical = validate_sqlite_db_path(path)?;
    let mut opts = sqlx::sqlite::SqliteConnectOptions::new().filename(canonical);
    opts = opts.busy_timeout(sql_query_timeout_duration());
    Ok(opts)
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "lowercase")]
pub enum SqlDataType {
    Any,
    String,
    Number,
    Boolean,
    Date,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct SqlQueryResult {
    pub columns: Vec<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub types: Option<HashMap<String, SqlDataType>>,
    pub rows: Vec<Vec<JsonValue>>,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct SqlSchemaResult {
    pub columns: Vec<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub types: Option<HashMap<String, SqlDataType>>,
}

#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct SqliteConnectionDescriptor {
    path: Option<String>,
    in_memory: Option<bool>,
}

#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct PostgresConnectionDescriptor {
    url: Option<String>,
    host: Option<String>,
    port: Option<u16>,
    database: Option<String>,
    #[serde(alias = "username")]
    user: Option<String>,
    ssl: Option<bool>,
}

#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct OdbcConnectionDescriptor {
    connection_string: String,
}

fn connection_kind(connection: &JsonValue) -> Result<String> {
    let kind = connection
        .get("kind")
        .and_then(|v| v.as_str())
        .map(|s| s.to_string())
        .ok_or_else(|| anyhow!("SQL connection descriptor must be an object with a string 'kind' field"))?;
    Ok(kind)
}

fn credential_string(credentials: Option<&JsonValue>, key: &str) -> Option<String> {
    let creds = credentials?;
    let obj = creds.as_object()?;
    obj.get(key)?.as_str().map(|s| s.to_string())
}

fn credential_password(credentials: Option<&JsonValue>) -> Option<String> {
    match credentials? {
        JsonValue::String(s) => Some(s.clone()),
        _ => credential_string(credentials, "password")
            .or_else(|| credential_string(credentials, "token"))
            .or_else(|| credential_string(credentials, "secret")),
    }
}

fn credential_username(credentials: Option<&JsonValue>) -> Option<String> {
    let creds = credentials?;
    let obj = creds.as_object()?;
    obj.get("user")
        .or_else(|| obj.get("username"))
        .and_then(|v| v.as_str())
        .map(|s| s.to_string())
}

fn normalize_odbc_key(input: &str) -> String {
    input
        .trim()
        .chars()
        .filter(|ch| !ch.is_whitespace() && *ch != '_' && *ch != '-')
        .map(|ch| ch.to_ascii_lowercase())
        .collect()
}

fn parse_odbc_connection_string(connection_string: &str) -> HashMap<String, String> {
    let mut parts = Vec::new();
    let mut current = String::new();
    let mut in_braces = false;

    for ch in connection_string.chars() {
        match ch {
            '{' => {
                in_braces = true;
                current.push(ch);
            }
            '}' => {
                in_braces = false;
                current.push(ch);
            }
            ';' if !in_braces => {
                let trimmed = current.trim();
                if !trimmed.is_empty() {
                    parts.push(trimmed.to_string());
                }
                current.clear();
            }
            _ => current.push(ch),
        }
    }

    let trimmed = current.trim();
    if !trimmed.is_empty() {
        parts.push(trimmed.to_string());
    }

    let mut out = HashMap::new();
    for part in parts {
        let Some((raw_key, raw_value)) = part.split_once('=') else {
            continue;
        };
        let key = normalize_odbc_key(raw_key);
        let mut value = raw_value.trim().to_string();
        if value.starts_with('{') && value.ends_with('}') && value.len() >= 2 {
            value = value[1..value.len() - 1].to_string();
        }
        if (value.starts_with('"') && value.ends_with('"')) || (value.starts_with('\'') && value.ends_with('\'')) {
            value = value[1..value.len() - 1].to_string();
        }
        if !key.is_empty() {
            out.insert(key, value);
        }
    }
    out
}

fn odbc_first<'a>(props: &'a HashMap<String, String>, keys: &[&str]) -> Option<&'a str> {
    for key in keys {
        if let Some(value) = props.get(*key) {
            if !value.trim().is_empty() {
                return Some(value.as_str());
            }
        }
    }
    None
}

fn odbc_driver_name(props: &HashMap<String, String>) -> Option<&str> {
    odbc_first(props, &["driver", "drv"])
}

fn parse_host_port(input: &str) -> (String, Option<u16>) {
    let trimmed = input.trim();
    if trimmed.is_empty() {
        return (String::new(), None);
    }

    if let Some((host, port_str)) = trimmed.rsplit_once(',') {
        if let Ok(port) = port_str.trim().parse::<u16>() {
            return (host.trim().trim_start_matches('[').trim_end_matches(']').to_string(), Some(port));
        }
    }

    if let Some((host, port_str)) = trimmed.rsplit_once(':') {
        let host_trimmed = host.trim();
        let port_trimmed = port_str.trim();
        if host_trimmed.starts_with('[') && host_trimmed.ends_with(']') {
            if let Ok(port) = port_trimmed.parse::<u16>() {
                return (
                    host_trimmed
                        .trim_start_matches('[')
                        .trim_end_matches(']')
                        .to_string(),
                    Some(port),
                );
            }
        } else if !host_trimmed.contains(':') {
            if let Ok(port) = port_trimmed.parse::<u16>() {
                return (host_trimmed.to_string(), Some(port));
            }
        }
    }

    (trimmed.trim_start_matches('[').trim_end_matches(']').to_string(), None)
}

fn sqlite_type_to_data_type(type_name: &str) -> SqlDataType {
    let normalized = type_name.trim().to_ascii_lowercase();
    if normalized.contains("int") || normalized.contains("real") || normalized.contains("floa") || normalized.contains("doub") || normalized.contains("num") {
        return SqlDataType::Number;
    }
    if normalized.contains("bool") {
        return SqlDataType::Boolean;
    }
    if normalized.contains("date") || normalized.contains("time") {
        return SqlDataType::Date;
    }
    if normalized.contains("char")
        || normalized.contains("text")
        || normalized.contains("clob")
        || normalized.contains("varchar")
    {
        return SqlDataType::String;
    }
    SqlDataType::Any
}

fn postgres_type_to_data_type(type_name: &str) -> SqlDataType {
    let normalized = type_name.trim().to_ascii_lowercase();
    match normalized.as_str() {
        "bool" => SqlDataType::Boolean,
        "int2" | "int4" | "int8" | "float4" | "float8" | "numeric" | "money" => SqlDataType::Number,
        "date" | "timestamp" | "timestamptz" | "time" | "timetz" => SqlDataType::Date,
        "text" | "varchar" | "bpchar" | "name" | "char" => SqlDataType::String,
        _ => {
            if normalized.starts_with("int") || normalized.starts_with("float") {
                SqlDataType::Number
            } else if normalized.contains("timestamp") || normalized.contains("date") || normalized.contains("time") {
                SqlDataType::Date
            } else {
                SqlDataType::Any
            }
        }
    }
}

async fn sqlite_schema(
    opts: &sqlx::sqlite::SqliteConnectOptions,
    sql: &str,
) -> Result<SqlSchemaResult> {
    let mut conn = sqlx::SqliteConnection::connect_with(opts)
        .await
        .context("connect sqlite")?;
    let describe = conn.describe(sql).await.context("describe sqlite query")?;

    let mut columns = Vec::new();
    let mut types = HashMap::new();
    for col in describe.columns() {
        let name = col.name().to_string();
        let ty = sqlite_type_to_data_type(col.type_info().name());
        columns.push(name.clone());
        types.insert(name, ty);
    }

    Ok(SqlSchemaResult {
        columns,
        types: Some(types),
    })
}

async fn postgres_schema(
    opts: &sqlx::postgres::PgConnectOptions,
    sql: &str,
) -> Result<SqlSchemaResult> {
    let mut conn = sqlx::PgConnection::connect_with(opts)
        .await
        .context("connect postgres")?;
    // Defense-in-depth: apply the same statement timeout used for `sql_query` so schema discovery
    // cannot hang indefinitely on slow planning/metadata operations.
    let timeout_ms = sql_query_timeout_ms();
    let _ = sqlx::query(&format!("SET statement_timeout = {timeout_ms}"))
        .execute(&mut conn)
        .await;
    let describe = conn.describe(sql).await.context("describe postgres query")?;

    let mut columns = Vec::new();
    let mut types = HashMap::new();
    for col in describe.columns() {
        let name = col.name().to_string();
        let ty = postgres_type_to_data_type(col.type_info().name());
        columns.push(name.clone());
        types.insert(name, ty);
    }

    Ok(SqlSchemaResult {
        columns,
        types: Some(types),
    })
}

fn bind_sqlite_param<'q>(
    query: sqlx::query::Query<'q, sqlx::Sqlite, sqlx::sqlite::SqliteArguments<'q>>,
    value: &'q JsonValue,
) -> sqlx::query::Query<'q, sqlx::Sqlite, sqlx::sqlite::SqliteArguments<'q>> {
    match value {
        JsonValue::Null => query.bind(Option::<String>::None),
        JsonValue::Bool(b) => query.bind(*b),
        JsonValue::Number(n) => {
            if let Some(i) = n.as_i64() {
                query.bind(i)
            } else if let Some(f) = n.as_f64() {
                query.bind(f)
            } else {
                query.bind(n.to_string())
            }
        }
        JsonValue::String(s) => query.bind(s.as_str()),
        other => query.bind(other.to_string()),
    }
}

fn bind_postgres_param<'q>(
    query: sqlx::query::Query<'q, sqlx::Postgres, sqlx::postgres::PgArguments>,
    value: &'q JsonValue,
) -> sqlx::query::Query<'q, sqlx::Postgres, sqlx::postgres::PgArguments> {
    match value {
        JsonValue::Null => query.bind(Option::<String>::None),
        JsonValue::Bool(b) => query.bind(*b),
        JsonValue::Number(n) => {
            if let Some(i) = n.as_i64() {
                query.bind(i)
            } else if let Some(f) = n.as_f64() {
                query.bind(f)
            } else {
                query.bind(n.to_string())
            }
        }
        JsonValue::String(s) => query.bind(s.as_str()),
        other => query.bind(other.to_string()),
    }
}

fn sqlite_cell_to_json(
    row: &sqlx::sqlite::SqliteRow,
    idx: usize,
    ty: &SqlDataType,
    max_cell_bytes: usize,
) -> Result<JsonValue> {
    // SQLite uses dynamic typing; use the declared schema type as a hint.
    match ty {
        SqlDataType::Boolean => {
            if let Ok(v) = row.try_get::<Option<bool>, _>(idx) {
                return Ok(v.map(JsonValue::from).unwrap_or(JsonValue::Null));
            }
            if let Ok(v) = row.try_get::<Option<i64>, _>(idx) {
                return Ok(v
                    .map(|n| JsonValue::Bool(n != 0))
                    .unwrap_or(JsonValue::Null));
            }
        }
        SqlDataType::Number => {
            if let Ok(v) = row.try_get::<Option<i64>, _>(idx) {
                return Ok(v.map(JsonValue::from).unwrap_or(JsonValue::Null));
            }
            if let Ok(v) = row.try_get::<Option<f64>, _>(idx) {
                return Ok(v
                    .and_then(serde_json::Number::from_f64)
                    .map(JsonValue::Number)
                    .unwrap_or(JsonValue::Null));
            }
        }
        SqlDataType::String | SqlDataType::Date => {
            if let Ok(v) = row.try_get::<Option<&str>, _>(idx) {
                if let Some(s) = v {
                    if s.len() > max_cell_bytes {
                        return Err(anyhow!(
                            "SQL query returned a cell larger than {max_cell_bytes} bytes. \
                             Truncate large values (e.g. SUBSTR/LEFT) or select fewer columns."
                        ));
                    }
                    return Ok(JsonValue::String(s.to_string()));
                }
                return Ok(JsonValue::Null);
            }
            if let Ok(v) = row.try_get::<Option<String>, _>(idx) {
                if let Some(s) = v {
                    if s.len() > max_cell_bytes {
                        return Err(anyhow!(
                            "SQL query returned a cell larger than {max_cell_bytes} bytes. \
                             Truncate large values (e.g. SUBSTR/LEFT) or select fewer columns."
                        ));
                    }
                    return Ok(JsonValue::String(s));
                }
                return Ok(JsonValue::Null);
            }
        }
        SqlDataType::Any => {}
    }

    // Fallback: attempt a few common decodes.
    if let Ok(v) = row.try_get::<Option<i64>, _>(idx) {
        return Ok(v.map(JsonValue::from).unwrap_or(JsonValue::Null));
    }
    if let Ok(v) = row.try_get::<Option<f64>, _>(idx) {
        return Ok(v
            .and_then(serde_json::Number::from_f64)
            .map(JsonValue::Number)
            .unwrap_or(JsonValue::Null));
    }
    if let Ok(v) = row.try_get::<Option<bool>, _>(idx) {
        return Ok(v.map(JsonValue::from).unwrap_or(JsonValue::Null));
    }
    if let Ok(v) = row.try_get::<Option<&str>, _>(idx) {
        if let Some(s) = v {
            if s.len() > max_cell_bytes {
                return Err(anyhow!(
                    "SQL query returned a cell larger than {max_cell_bytes} bytes. \
                     Truncate large values (e.g. SUBSTR/LEFT) or select fewer columns."
                ));
            }
            return Ok(JsonValue::String(s.to_string()));
        }
        return Ok(JsonValue::Null);
    }
    if let Ok(v) = row.try_get::<Option<String>, _>(idx) {
        if let Some(s) = v {
            if s.len() > max_cell_bytes {
                return Err(anyhow!(
                    "SQL query returned a cell larger than {max_cell_bytes} bytes. \
                     Truncate large values (e.g. SUBSTR/LEFT) or select fewer columns."
                ));
            }
            return Ok(JsonValue::String(s));
        }
        return Ok(JsonValue::Null);
    }
    Ok(JsonValue::Null)
}

fn postgres_cell_to_json(
    row: &sqlx::postgres::PgRow,
    idx: usize,
    pg_type_name: &str,
    max_cell_bytes: usize,
) -> Result<JsonValue> {
    // Prefer decoding based on the Postgres type name; fall back to a few generic attempts.
    match pg_type_name {
        "BOOL" => Ok(row
            .try_get::<Option<bool>, _>(idx)
            .ok()
            .flatten()
            .map(JsonValue::Bool)
            .unwrap_or(JsonValue::Null)),
        "INT2" => Ok(row
            .try_get::<Option<i16>, _>(idx)
            .ok()
            .flatten()
            .map(|v| JsonValue::from(i64::from(v)))
            .unwrap_or(JsonValue::Null)),
        "INT4" => Ok(row
            .try_get::<Option<i32>, _>(idx)
            .ok()
            .flatten()
            .map(|v| JsonValue::from(i64::from(v)))
            .unwrap_or(JsonValue::Null)),
        "INT8" => Ok(row
            .try_get::<Option<i64>, _>(idx)
            .ok()
            .flatten()
            .map(JsonValue::from)
            .unwrap_or(JsonValue::Null)),
        "FLOAT4" => Ok(row
            .try_get::<Option<f32>, _>(idx)
            .ok()
            .flatten()
            .map(|v| v as f64)
            .and_then(serde_json::Number::from_f64)
            .map(JsonValue::Number)
            .unwrap_or(JsonValue::Null)),
        "FLOAT8" => Ok(row
            .try_get::<Option<f64>, _>(idx)
            .ok()
            .flatten()
            .and_then(serde_json::Number::from_f64)
            .map(JsonValue::Number)
            .unwrap_or(JsonValue::Null)),
        "DATE" => {
            let Some(d) = row
                .try_get::<Option<sqlx::types::chrono::NaiveDate>, _>(idx)
                .ok()
                .flatten()
            else {
                return Ok(JsonValue::Null);
            };
            let s = format!("{}T00:00:00.000Z", d.format("%Y-%m-%d"));
            if s.len() > max_cell_bytes {
                return Err(anyhow!(
                    "SQL query returned a cell larger than {max_cell_bytes} bytes. \
                     Truncate large values (e.g. SUBSTR/LEFT) or select fewer columns."
                ));
            }
            Ok(JsonValue::String(s))
        }
        "TIMESTAMP" => {
            let Some(dt) = row
                .try_get::<Option<sqlx::types::chrono::NaiveDateTime>, _>(idx)
                .ok()
                .flatten()
            else {
                return Ok(JsonValue::Null);
            };
            let s = format!("{}Z", dt.format("%Y-%m-%dT%H:%M:%S%.f"));
            if s.len() > max_cell_bytes {
                return Err(anyhow!(
                    "SQL query returned a cell larger than {max_cell_bytes} bytes. \
                     Truncate large values (e.g. SUBSTR/LEFT) or select fewer columns."
                ));
            }
            Ok(JsonValue::String(s))
        }
        "TIMESTAMPTZ" => {
            let Some(dt) = row
                .try_get::<Option<sqlx::types::chrono::DateTime<sqlx::types::chrono::Utc>>, _>(idx)
                .ok()
                .flatten()
            else {
                return Ok(JsonValue::Null);
            };
            let s = dt.to_rfc3339();
            if s.len() > max_cell_bytes {
                return Err(anyhow!(
                    "SQL query returned a cell larger than {max_cell_bytes} bytes. \
                     Truncate large values (e.g. SUBSTR/LEFT) or select fewer columns."
                ));
            }
            Ok(JsonValue::String(s))
        }
        _ => {
            if let Ok(v) = row.try_get::<Option<&str>, _>(idx) {
                if let Some(s) = v {
                    if s.len() > max_cell_bytes {
                        return Err(anyhow!(
                            "SQL query returned a cell larger than {max_cell_bytes} bytes. \
                             Truncate large values (e.g. SUBSTR/LEFT) or select fewer columns."
                        ));
                    }
                    return Ok(JsonValue::String(s.to_string()));
                }
                return Ok(JsonValue::Null);
            }
            if let Ok(v) = row.try_get::<Option<String>, _>(idx) {
                if let Some(s) = v {
                    if s.len() > max_cell_bytes {
                        return Err(anyhow!(
                            "SQL query returned a cell larger than {max_cell_bytes} bytes. \
                             Truncate large values (e.g. SUBSTR/LEFT) or select fewer columns."
                        ));
                    }
                    return Ok(JsonValue::String(s));
                }
                return Ok(JsonValue::Null);
            }
            if let Ok(v) = row.try_get::<Option<i64>, _>(idx) {
                return Ok(v.map(JsonValue::from).unwrap_or(JsonValue::Null));
            }
            if let Ok(v) = row.try_get::<Option<i32>, _>(idx) {
                return Ok(v
                    .map(|n| JsonValue::from(i64::from(n)))
                    .unwrap_or(JsonValue::Null));
            }
            if let Ok(v) = row.try_get::<Option<f64>, _>(idx) {
                return Ok(v
                    .and_then(serde_json::Number::from_f64)
                    .map(JsonValue::Number)
                    .unwrap_or(JsonValue::Null));
            }
            if let Ok(v) = row.try_get::<Option<bool>, _>(idx) {
                return Ok(v.map(JsonValue::from).unwrap_or(JsonValue::Null));
            }
            Ok(JsonValue::Null)
        }
    }
}

async fn query_sqlite(
    opts: &sqlx::sqlite::SqliteConnectOptions,
    sql: &str,
    params: &[JsonValue],
) -> Result<SqlQueryResult> {
    with_sql_query_timeout("SQLite", async {
        // Fetch schema first (best-effort). We do this on a separate connection so schema discovery
        // failures don't prevent query execution.
        let schema =
            with_sql_query_timeout("SQLite schema discovery", sqlite_schema(opts, sql))
                .await
                .ok();

        let mut conn = sqlx::SqliteConnection::connect_with(opts)
            .await
            .context("connect sqlite")?;

        let mut query = sqlx::query(sql);
        for value in params {
            query = bind_sqlite_param(query, value);
        }

        let mut stream = query.fetch(&mut conn);

        let mut columns = schema
            .as_ref()
            .map(|s| s.columns.clone())
            .unwrap_or_default();
        let types = schema.as_ref().and_then(|s| s.types.clone());

        let mut column_types: Vec<SqlDataType> = Vec::new();
        if !columns.is_empty() {
            column_types = columns
                .iter()
                .map(|name| {
                    types
                        .as_ref()
                        .and_then(|m| m.get(name))
                        .cloned()
                        .unwrap_or(SqlDataType::Any)
                })
                .collect();
        }

        let mut out_rows: Vec<Vec<JsonValue>> = Vec::new();
        let mut row_count: usize = 0;
        let max_bytes = sql_query_max_bytes();
        let max_cell_bytes = sql_query_max_cell_bytes();
        let mut total_bytes: usize = 0;

        while let Some(row) = stream
            .try_next()
            .await
            .context("execute sqlite query")?
        {
            row_count += 1;
            if row_count > MAX_SQL_QUERY_ROWS {
                return Err(anyhow!(
                    "SQL query returned more than {MAX_SQL_QUERY_ROWS} rows. Please add a LIMIT clause."
                ));
            }

            if columns.is_empty() {
                columns = row
                    .columns()
                    .iter()
                    .map(|c| c.name().to_string())
                    .collect();
                column_types = columns
                    .iter()
                    .map(|name| {
                        types
                            .as_ref()
                            .and_then(|m| m.get(name))
                            .cloned()
                            .unwrap_or(SqlDataType::Any)
                    })
                    .collect();
            }

            let cell_count = row_count.saturating_mul(columns.len());
            if cell_count > MAX_SQL_QUERY_CELLS {
                return Err(anyhow!(
                    "SQL query returned too many values (>{MAX_SQL_QUERY_CELLS} cells). \
                     Please add a LIMIT clause or select fewer columns."
                ));
            }

            let mut out: Vec<JsonValue> = Vec::new();
            let _ = out.try_reserve(columns.len());
            for idx in 0..columns.len() {
                let value = sqlite_cell_to_json(&row, idx, &column_types[idx], max_cell_bytes)?;
                total_bytes = total_bytes.saturating_add(sql_value_estimated_bytes(&value));
                if total_bytes > max_bytes {
                    return Err(anyhow!(
                        "SQL query result exceeded the maximum size ({max_bytes} bytes). \
                         Please add a LIMIT clause or select fewer columns."
                    ));
                }
                out.push(value);
            }
            out_rows.push(out);
        }

        Ok(SqlQueryResult {
            columns,
            types,
            rows: out_rows,
        })
    })
    .await
}

async fn query_postgres(
    opts: &sqlx::postgres::PgConnectOptions,
    sql: &str,
    params: &[JsonValue],
) -> Result<SqlQueryResult> {
    with_sql_query_timeout("Postgres", async {
        let schema = postgres_schema(opts, sql).await.ok();

        let mut conn = sqlx::PgConnection::connect_with(opts)
            .await
            .context("connect postgres")?;

        // Defense-in-depth: enforce a server-side statement timeout as well.
        let timeout_ms = sql_query_timeout_ms();
        let _ = sqlx::query(&format!("SET statement_timeout = {timeout_ms}"))
            .execute(&mut conn)
            .await;

        let mut query = sqlx::query(sql);
        for value in params {
            query = bind_postgres_param(query, value);
        }

        let mut stream = query.fetch(&mut conn);

        let mut columns = schema
            .as_ref()
            .map(|s| s.columns.clone())
            .unwrap_or_default();
        let types = schema.as_ref().and_then(|s| s.types.clone());

        let mut out_rows: Vec<Vec<JsonValue>> = Vec::new();
        let mut row_count: usize = 0;
        let max_bytes = sql_query_max_bytes();
        let max_cell_bytes = sql_query_max_cell_bytes();
        let mut total_bytes: usize = 0;

        while let Some(row) = stream
            .try_next()
            .await
            .context("execute postgres query")?
        {
            row_count += 1;
            if row_count > MAX_SQL_QUERY_ROWS {
                return Err(anyhow!(
                    "SQL query returned more than {MAX_SQL_QUERY_ROWS} rows. Please add a LIMIT clause."
                ));
            }

            if columns.is_empty() {
                columns = row
                    .columns()
                    .iter()
                    .map(|c| c.name().to_string())
                    .collect();
            }

            let cell_count = row_count.saturating_mul(columns.len());
            if cell_count > MAX_SQL_QUERY_CELLS {
                return Err(anyhow!(
                    "SQL query returned too many values (>{MAX_SQL_QUERY_CELLS} cells). \
                     Please add a LIMIT clause or select fewer columns."
                ));
            }

            let mut out: Vec<JsonValue> = Vec::new();
            let _ = out.try_reserve(columns.len());
            for idx in 0..columns.len() {
                let type_name = row
                    .columns()
                    .get(idx)
                    .map(|c| c.type_info().name())
                    .unwrap_or("");
                let value = postgres_cell_to_json(&row, idx, type_name, max_cell_bytes)?;
                total_bytes = total_bytes.saturating_add(sql_value_estimated_bytes(&value));
                if total_bytes > max_bytes {
                    return Err(anyhow!(
                        "SQL query result exceeded the maximum size ({max_bytes} bytes). \
                         Please add a LIMIT clause or select fewer columns."
                    ));
                }
                out.push(value);
            }
            out_rows.push(out);
        }

        Ok(SqlQueryResult {
            columns,
            types,
            rows: out_rows,
        })
    })
    .await
}

pub async fn sql_query(
    connection: JsonValue,
    sql: String,
    params: Vec<JsonValue>,
    credentials: Option<JsonValue>,
) -> Result<SqlQueryResult> {
    if sql.as_bytes().len() > MAX_SQL_QUERY_TEXT_BYTES {
        return Err(anyhow!(
            "SQL query text exceeds MAX_SQL_QUERY_TEXT_BYTES ({MAX_SQL_QUERY_TEXT_BYTES} bytes)"
        ));
    }
    if params.len() > MAX_SQL_QUERY_PARAMS {
        return Err(anyhow!(
            "SQL query has too many parameters: {} (max MAX_SQL_QUERY_PARAMS = {MAX_SQL_QUERY_PARAMS})",
            params.len()
        ));
    }
    for (idx, value) in params.iter().enumerate() {
        enforce_json_byte_size(
            value,
            MAX_SQL_QUERY_PARAM_BYTES,
            &format!("SQL query parameter[{idx}]"),
            "MAX_SQL_QUERY_PARAM_BYTES",
        )
        .map_err(|e| anyhow!(e))?;
    }
    enforce_json_byte_size(
        &connection,
        MAX_SQL_QUERY_CONNECTION_BYTES,
        "SQL connection descriptor",
        "MAX_SQL_QUERY_CONNECTION_BYTES",
    )
    .map_err(|e| anyhow!(e))?;
    if let Some(credentials) = credentials.as_ref() {
        enforce_json_byte_size(
            credentials,
            MAX_SQL_QUERY_CREDENTIALS_BYTES,
            "SQL credentials",
            "MAX_SQL_QUERY_CREDENTIALS_BYTES",
        )
        .map_err(|e| anyhow!(e))?;
    }

    let kind = connection_kind(&connection)?;

    match kind.as_str() {
        "sqlite" => {
            let descriptor: SqliteConnectionDescriptor =
                serde_json::from_value(connection).context("invalid sqlite connection descriptor")?;
            let in_memory = descriptor.in_memory.unwrap_or(false);
            let mut opts = if in_memory {
                sqlx::sqlite::SqliteConnectOptions::from_str("sqlite::memory:")?
                    .busy_timeout(sql_query_timeout_duration())
            } else {
                let path = descriptor
                    .path
                    .ok_or_else(|| anyhow!("sqlite connection requires `path`"))?;
                sqlite_connect_options_from_path(&path)?
            };
            opts = opts.create_if_missing(false);
            query_sqlite(&opts, &sql, &params).await
        }
        "postgres" => {
            let descriptor: PostgresConnectionDescriptor =
                serde_json::from_value(connection).context("invalid postgres connection descriptor")?;
            use sqlx::postgres::{PgConnectOptions, PgSslMode};
            let mut opts = if let Some(url) = descriptor.url.clone() {
                PgConnectOptions::from_str(&url).context("invalid postgres url")?
            } else {
                let host = descriptor
                    .host
                    .clone()
                    .ok_or_else(|| anyhow!("postgres connection requires `host` or `url`"))?;
                let mut opts = PgConnectOptions::new().host(&host);
                if let Some(port) = descriptor.port {
                    opts = opts.port(port);
                }
                if let Some(database) = descriptor.database.clone() {
                    opts = opts.database(&database);
                }
                if let Some(user) = descriptor.user.clone() {
                    opts = opts.username(&user);
                }
                opts
            };

            if let Some(host) = descriptor.host {
                opts = opts.host(&host);
            }
            if let Some(port) = descriptor.port {
                opts = opts.port(port);
            }
            if let Some(database) = descriptor.database {
                opts = opts.database(&database);
            }
            if let Some(user) = descriptor.user {
                opts = opts.username(&user);
            } else if let Some(user) = credential_username(credentials.as_ref()) {
                opts = opts.username(&user);
            }

            if let Some(password) = credential_password(credentials.as_ref()) {
                opts = opts.password(&password);
            }

            if let Some(ssl) = descriptor.ssl {
                opts = opts.ssl_mode(if ssl { PgSslMode::Require } else { PgSslMode::Disable });
            }

            query_postgres(&opts, &sql, &params).await
        }
        "odbc" => {
            let descriptor: OdbcConnectionDescriptor =
                serde_json::from_value(connection).context("invalid odbc connection descriptor")?;
            let props = parse_odbc_connection_string(&descriptor.connection_string);
            let driver = match odbc_driver_name(&props) {
                Some(value) => value.to_string(),
                None => {
                    if let Some(dsn) = odbc_first(&props, &["dsn"]) {
                        return Err(anyhow!(
                            "ODBC DSN connections (DSN={dsn}) are not supported. Provide an explicit Driver and connection details (e.g. Driver={{PostgreSQL}};Server=...;Database=...; or Driver=SQLite3;Database=...;)."
                        ));
                    }
                    return Err(anyhow!("odbc connection string requires a `Driver` field"));
                }
            };
            let driver_lower = driver.to_ascii_lowercase();

            if driver_lower.contains("sqlite") {
                let path = odbc_first(&props, &["database", "dbq", "datasource"])
                    .map(|s| s.to_string())
                    .ok_or_else(|| anyhow!("odbc sqlite connection string requires a `Database` (or `Data Source`) field"))?;
                let in_memory = path.trim().eq_ignore_ascii_case(":memory:") || path.trim().eq_ignore_ascii_case("memory");
                let mut opts = if in_memory {
                    sqlx::sqlite::SqliteConnectOptions::from_str("sqlite::memory:")?
                        .busy_timeout(sql_query_timeout_duration())
                } else {
                    sqlite_connect_options_from_path(&path)?
                };
                opts = opts.create_if_missing(false);
                query_sqlite(&opts, &sql, &params).await
            } else if driver_lower.contains("postgres") {
                let server_raw = odbc_first(&props, &["server", "host", "hostname", "servername", "address"])
                    .map(|s| s.to_string())
                    .ok_or_else(|| anyhow!("odbc postgres connection string requires a `Server` (or `Host`) field"))?;
                let database = odbc_first(&props, &["database", "db", "dbname"])
                    .map(|s| s.to_string())
                    .ok_or_else(|| anyhow!("odbc postgres connection string requires a `Database` field"))?;
                let mut port = odbc_first(&props, &["port"]).and_then(|s| s.trim().parse::<u16>().ok());
                let (host, embedded_port) = if port.is_none() {
                    parse_host_port(&server_raw)
                } else {
                    (server_raw, None)
                };
                if port.is_none() {
                    port = embedded_port;
                }
                let user = credential_username(credentials.as_ref())
                    .or_else(|| odbc_first(&props, &["uid", "user", "username", "userid"]).map(|s| s.to_string()));
                let password = credential_password(credentials.as_ref()).or_else(|| {
                    odbc_first(&props, &["pwd", "password", "passwd"]).map(|s| s.to_string())
                });
                let ssl_mode = odbc_first(&props, &["sslmode"]).map(|s| s.to_ascii_lowercase());

                use sqlx::postgres::{PgConnectOptions, PgSslMode};
                let mut opts = PgConnectOptions::new().host(&host).database(&database);
                if let Some(port) = port {
                    opts = opts.port(port);
                }
                if let Some(user) = user {
                    opts = opts.username(&user);
                }
                if let Some(password) = password {
                    opts = opts.password(&password);
                }
                if let Some(ssl_mode) = ssl_mode {
                    if ssl_mode == "require" || ssl_mode == "verify-full" || ssl_mode == "verify-ca" {
                        opts = opts.ssl_mode(PgSslMode::Require);
                    } else if ssl_mode == "disable" {
                        opts = opts.ssl_mode(PgSslMode::Disable);
                    }
                }

                query_postgres(&opts, &sql, &params).await
            } else {
                Err(anyhow!(
                    "Unsupported ODBC driver '{driver}' (supported: SQLite, PostgreSQL)"
                ))
            }
        }
        "sql" | "sqlserver" | "mssql" => Err(anyhow!(
            "SQL Server connections are not supported yet (kind: 'sql'). Supported kinds: sqlite, postgres, odbc."
        )),
        other => Err(anyhow!(
            "Unsupported SQL connection kind '{other}' (supported: sqlite, postgres, odbc)"
        )),
    }
}

pub async fn sql_get_schema(
    connection: JsonValue,
    sql: String,
    credentials: Option<JsonValue>,
) -> Result<SqlSchemaResult> {
    if sql.as_bytes().len() > MAX_SQL_QUERY_TEXT_BYTES {
        return Err(anyhow!(
            "SQL query text exceeds MAX_SQL_QUERY_TEXT_BYTES ({MAX_SQL_QUERY_TEXT_BYTES} bytes)"
        ));
    }
    enforce_json_byte_size(
        &connection,
        MAX_SQL_QUERY_CONNECTION_BYTES,
        "SQL connection descriptor",
        "MAX_SQL_QUERY_CONNECTION_BYTES",
    )
    .map_err(|e| anyhow!(e))?;
    if let Some(credentials) = credentials.as_ref() {
        enforce_json_byte_size(
            credentials,
            MAX_SQL_QUERY_CREDENTIALS_BYTES,
            "SQL credentials",
            "MAX_SQL_QUERY_CREDENTIALS_BYTES",
        )
        .map_err(|e| anyhow!(e))?;
    }

    let kind = connection_kind(&connection)?;

    with_sql_query_timeout("SQL schema discovery", async move {
        match kind.as_str() {
            "sqlite" => {
                let descriptor: SqliteConnectionDescriptor =
                    serde_json::from_value(connection).context("invalid sqlite connection descriptor")?;
                let in_memory = descriptor.in_memory.unwrap_or(false);
                let mut opts = if in_memory {
                    sqlx::sqlite::SqliteConnectOptions::from_str("sqlite::memory:")?
                        .busy_timeout(sql_query_timeout_duration())
                } else {
                    let path = descriptor
                        .path
                        .ok_or_else(|| anyhow!("sqlite connection requires `path`"))?;
                    sqlite_connect_options_from_path(&path)?
                };
                opts = opts.create_if_missing(false);
                sqlite_schema(&opts, &sql).await
            }
            "postgres" => {
                let descriptor: PostgresConnectionDescriptor =
                    serde_json::from_value(connection).context("invalid postgres connection descriptor")?;
                use sqlx::postgres::{PgConnectOptions, PgSslMode};
                let mut opts = if let Some(url) = descriptor.url.clone() {
                    PgConnectOptions::from_str(&url).context("invalid postgres url")?
                } else {
                    let host = descriptor
                        .host
                        .clone()
                        .ok_or_else(|| anyhow!("postgres connection requires `host` or `url`"))?;
                    let mut opts = PgConnectOptions::new().host(&host);
                    if let Some(port) = descriptor.port {
                        opts = opts.port(port);
                    }
                    if let Some(database) = descriptor.database.clone() {
                        opts = opts.database(&database);
                    }
                    if let Some(user) = descriptor.user.clone() {
                        opts = opts.username(&user);
                    }
                    opts
                };

                if let Some(host) = descriptor.host {
                    opts = opts.host(&host);
                }
                if let Some(port) = descriptor.port {
                    opts = opts.port(port);
                }
                if let Some(database) = descriptor.database {
                    opts = opts.database(&database);
                }
                if let Some(user) = descriptor.user {
                    opts = opts.username(&user);
                } else if let Some(user) = credential_username(credentials.as_ref()) {
                    opts = opts.username(&user);
                }

                if let Some(password) = credential_password(credentials.as_ref()) {
                    opts = opts.password(&password);
                }
                if let Some(ssl) = descriptor.ssl {
                    opts = opts.ssl_mode(if ssl { PgSslMode::Require } else { PgSslMode::Disable });
                }

                postgres_schema(&opts, &sql).await
            }
            "odbc" => {
                let descriptor: OdbcConnectionDescriptor =
                    serde_json::from_value(connection).context("invalid odbc connection descriptor")?;
                let props = parse_odbc_connection_string(&descriptor.connection_string);
                let driver = match odbc_driver_name(&props) {
                    Some(value) => value.to_string(),
                    None => {
                        if let Some(dsn) = odbc_first(&props, &["dsn"]) {
                            return Err(anyhow!(
                                "ODBC DSN connections (DSN={dsn}) are not supported. Provide an explicit Driver and connection details (e.g. Driver={{PostgreSQL}};Server=...;Database=...; or Driver=SQLite3;Database=...;)."
                            ));
                        }
                        return Err(anyhow!("odbc connection string requires a `Driver` field"));
                    }
                };
                let driver_lower = driver.to_ascii_lowercase();

                if driver_lower.contains("sqlite") {
                    let path = odbc_first(&props, &["database", "dbq", "datasource"])
                        .map(|s| s.to_string())
                        .ok_or_else(|| anyhow!("odbc sqlite connection string requires a `Database` (or `Data Source`) field"))?;
                    let in_memory =
                        path.trim().eq_ignore_ascii_case(":memory:") || path.trim().eq_ignore_ascii_case("memory");
                    let mut opts = if in_memory {
                        sqlx::sqlite::SqliteConnectOptions::from_str("sqlite::memory:")?
                            .busy_timeout(sql_query_timeout_duration())
                    } else {
                        sqlite_connect_options_from_path(&path)?
                    };
                    opts = opts.create_if_missing(false);
                    sqlite_schema(&opts, &sql).await
                } else if driver_lower.contains("postgres") {
                    let server_raw = odbc_first(
                        &props,
                        &["server", "host", "hostname", "servername", "address"],
                    )
                    .map(|s| s.to_string())
                    .ok_or_else(|| anyhow!("odbc postgres connection string requires a `Server` (or `Host`) field"))?;
                    let database = odbc_first(&props, &["database", "db", "dbname"])
                        .map(|s| s.to_string())
                        .ok_or_else(|| anyhow!("odbc postgres connection string requires a `Database` field"))?;
                    let mut port = odbc_first(&props, &["port"]).and_then(|s| s.trim().parse::<u16>().ok());
                    let (host, embedded_port) = if port.is_none() {
                        parse_host_port(&server_raw)
                    } else {
                        (server_raw, None)
                    };
                    if port.is_none() {
                        port = embedded_port;
                    }
                    let user = credential_username(credentials.as_ref()).or_else(|| {
                        odbc_first(&props, &["uid", "user", "username", "userid"]).map(|s| s.to_string())
                    });
                    let password = credential_password(credentials.as_ref()).or_else(|| {
                        odbc_first(&props, &["pwd", "password", "passwd"]).map(|s| s.to_string())
                    });
                    let ssl_mode = odbc_first(&props, &["sslmode"]).map(|s| s.to_ascii_lowercase());

                    use sqlx::postgres::{PgConnectOptions, PgSslMode};
                    let mut opts = PgConnectOptions::new().host(&host).database(&database);
                    if let Some(port) = port {
                        opts = opts.port(port);
                    }
                    if let Some(user) = user {
                        opts = opts.username(&user);
                    }
                    if let Some(password) = password {
                        opts = opts.password(&password);
                    }
                    if let Some(ssl_mode) = ssl_mode {
                        if ssl_mode == "require" || ssl_mode == "verify-full" || ssl_mode == "verify-ca" {
                            opts = opts.ssl_mode(PgSslMode::Require);
                        } else if ssl_mode == "disable" {
                            opts = opts.ssl_mode(PgSslMode::Disable);
                        }
                    }

                    postgres_schema(&opts, &sql).await
                } else {
                    Err(anyhow!(
                        "Unsupported ODBC driver '{driver}' (supported: SQLite, PostgreSQL)"
                    ))
                }
            }
            "sql" | "sqlserver" | "mssql" => Err(anyhow!(
                "SQL Server connections are not supported yet (kind: 'sql'). Supported kinds: sqlite, postgres, odbc."
            )),
            other => Err(anyhow!(
                "Unsupported SQL connection kind '{other}' (supported: sqlite, postgres, odbc)"
            )),
        }
    })
    .await
}

#[cfg(test)]
mod tests {
    use super::*;
    use sqlx::Connection;
    use std::sync::{Mutex, OnceLock};

    static ENV_MUTEX: OnceLock<Mutex<()>> = OnceLock::new();

    fn env_mutex() -> &'static Mutex<()> {
        ENV_MUTEX.get_or_init(|| Mutex::new(()))
    }

    async fn make_shared_in_memory_sqlite() -> (sqlx::sqlite::SqliteConnectOptions, sqlx::SqliteConnection) {
        // SQLite shared-cache in-memory database. We keep one connection open for the lifetime of
        // each test so the database doesn't get dropped.
        let db_id = uuid::Uuid::new_v4();
        let dsn = format!("sqlite://file:formula_sql_test_{db_id}?mode=memory&cache=shared");
        // `query_sqlite` opens multiple connections (one for schema discovery and one for the
        // actual query). Configure a busy timeout so concurrent schema/query connections don't
        // sporadically fail under load on shared-cache in-memory databases.
        let opts = sqlx::sqlite::SqliteConnectOptions::from_str(&dsn)
            .expect("parse sqlite dsn")
            .busy_timeout(sql_query_timeout_duration());
        let keeper = sqlx::SqliteConnection::connect_with(&opts)
            .await
            .expect("connect sqlite");
        (opts, keeper)
    }

    async fn seed_numbers_table(conn: &mut sqlx::SqliteConnection, count: i64) {
        sqlx::query("CREATE TABLE numbers (n INTEGER NOT NULL);")
            .execute(&mut *conn)
            .await
            .expect("create table");
        // Use a recursive CTE to insert a moderate number of rows quickly/deterministically.
        // (SQLite's default recursion limit is 1000, so keep `count` comfortably below that.)
        sqlx::query(
            "WITH RECURSIVE cnt(x) AS (SELECT 1 UNION ALL SELECT x+1 FROM cnt WHERE x < ?1)
             INSERT INTO numbers SELECT x FROM cnt;",
        )
        .bind(count)
        .execute(&mut *conn)
        .await
        .expect("seed rows");
    }

    #[tokio::test]
    async fn sqlite_query_enforces_max_rows() {
        // Prevent concurrent env var overrides (timeout/max-bytes) from affecting this test.
        let _guard = env_mutex()
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());

        let (opts, mut keeper) = make_shared_in_memory_sqlite().await;
        seed_numbers_table(&mut keeper, 400).await;

        // 400x400 = 160k rows (> MAX_SQL_QUERY_ROWS).
        let sql = "SELECT a.n, b.n FROM numbers a CROSS JOIN numbers b";
        let err = query_sqlite(&opts, sql, &[]).await.unwrap_err();
        let msg = err.to_string();
        assert!(
            msg.contains("LIMIT") || msg.contains("limit"),
            "expected error to suggest adding LIMIT, got: {msg}"
        );
    }

    #[tokio::test]
    async fn sqlite_query_allows_exact_limit() {
        // Prevent concurrent env var overrides (timeout/max-bytes) from affecting this test.
        let _guard = env_mutex()
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());

        let (opts, mut keeper) = make_shared_in_memory_sqlite().await;
        seed_numbers_table(&mut keeper, 400).await;

        let sql = format!(
            "SELECT a.n, b.n FROM numbers a CROSS JOIN numbers b LIMIT {}",
            MAX_SQL_QUERY_ROWS
        );
        let result = query_sqlite(&opts, &sql, &[]).await.expect("query");
        assert_eq!(result.rows.len(), MAX_SQL_QUERY_ROWS);
    }

    #[tokio::test]
    async fn with_sql_query_timeout_times_out() {
        // Keep env overrides (if any) isolated from other tests in this crate.
        let _guard = env_mutex()
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());

        let key = "FORMULA_SQL_QUERY_TIMEOUT_MS";
        let prev = std::env::var(key).ok();
        std::env::set_var(key, "1");

        let err = with_sql_query_timeout("Test", async {
            tokio::time::sleep(Duration::from_millis(25)).await;
            Ok::<_, anyhow::Error>(())
        })
        .await
        .unwrap_err();

        if let Some(prev) = prev {
            std::env::set_var(key, prev);
        } else {
            std::env::remove_var(key);
        }

        assert!(
            err.to_string().contains("exceeded"),
            "expected timeout error message, got: {err}"
        );
    }

    #[tokio::test]
    async fn sqlite_query_enforces_max_bytes() {
        // Keep env overrides (if any) isolated from other tests in this crate.
        let _guard = env_mutex()
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());

        let key = "FORMULA_SQL_QUERY_MAX_BYTES";
        let prev = std::env::var(key).ok();
        std::env::set_var(key, "1000");

        let (opts, _keeper) = make_shared_in_memory_sqlite().await;
        // Deterministically generate a ~2KiB string.
        let sql = "SELECT replace(hex(zeroblob(2000)), '00', 'a')";
        let err = query_sqlite(&opts, sql, &[]).await.unwrap_err();

        if let Some(prev) = prev {
            std::env::set_var(key, prev);
        } else {
            std::env::remove_var(key);
        }

        assert!(
            err.to_string().contains("maximum size"),
            "expected size limit error, got: {err}"
        );
    }

    #[tokio::test]
    async fn sqlite_query_enforces_max_cell_bytes() {
        // Keep env overrides (if any) isolated from other tests in this crate.
        let _guard = env_mutex()
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());

        let key = "FORMULA_SQL_QUERY_MAX_CELL_BYTES";
        let prev = std::env::var(key).ok();
        std::env::set_var(key, "10");

        let (opts, _keeper) = make_shared_in_memory_sqlite().await;
        let sql = "SELECT replace(hex(zeroblob(20)), '00', 'a')";
        let err = query_sqlite(&opts, sql, &[]).await.unwrap_err();

        if let Some(prev) = prev {
            std::env::set_var(key, prev);
        } else {
            std::env::remove_var(key);
        }

        assert!(
            err.to_string().contains("cell larger"),
            "expected per-cell size limit error, got: {err}"
        );
    }
}
