use anyhow::{anyhow, Context, Result};
use serde::{Deserialize, Serialize};
use serde_json::Value as JsonValue;
use sqlx::{Column, Connection, Executor, Row, TypeInfo};
use std::collections::HashMap;
use std::str::FromStr;

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

fn sqlite_cell_to_json(row: &sqlx::sqlite::SqliteRow, idx: usize, ty: &SqlDataType) -> JsonValue {
    // SQLite uses dynamic typing; use the declared schema type as a hint.
    match ty {
        SqlDataType::Boolean => {
            if let Ok(v) = row.try_get::<Option<bool>, _>(idx) {
                return v.map(JsonValue::from).unwrap_or(JsonValue::Null);
            }
            if let Ok(v) = row.try_get::<Option<i64>, _>(idx) {
                return v
                    .map(|n| JsonValue::Bool(n != 0))
                    .unwrap_or(JsonValue::Null);
            }
        }
        SqlDataType::Number => {
            if let Ok(v) = row.try_get::<Option<i64>, _>(idx) {
                return v.map(JsonValue::from).unwrap_or(JsonValue::Null);
            }
            if let Ok(v) = row.try_get::<Option<f64>, _>(idx) {
                return v
                    .and_then(serde_json::Number::from_f64)
                    .map(JsonValue::Number)
                    .unwrap_or(JsonValue::Null);
            }
        }
        SqlDataType::String | SqlDataType::Date => {
            if let Ok(v) = row.try_get::<Option<String>, _>(idx) {
                return v.map(JsonValue::from).unwrap_or(JsonValue::Null);
            }
        }
        SqlDataType::Any => {}
    }

    // Fallback: attempt a few common decodes.
    if let Ok(v) = row.try_get::<Option<i64>, _>(idx) {
        return v.map(JsonValue::from).unwrap_or(JsonValue::Null);
    }
    if let Ok(v) = row.try_get::<Option<f64>, _>(idx) {
        return v
            .and_then(serde_json::Number::from_f64)
            .map(JsonValue::Number)
            .unwrap_or(JsonValue::Null);
    }
    if let Ok(v) = row.try_get::<Option<bool>, _>(idx) {
        return v.map(JsonValue::from).unwrap_or(JsonValue::Null);
    }
    if let Ok(v) = row.try_get::<Option<String>, _>(idx) {
        return v.map(JsonValue::from).unwrap_or(JsonValue::Null);
    }
    JsonValue::Null
}

fn postgres_cell_to_json(row: &sqlx::postgres::PgRow, idx: usize, pg_type_name: &str) -> JsonValue {
    // Prefer decoding based on the Postgres type name; fall back to a few generic attempts.
    match pg_type_name {
        "BOOL" => row
            .try_get::<Option<bool>, _>(idx)
            .ok()
            .flatten()
            .map(JsonValue::Bool)
            .unwrap_or(JsonValue::Null),
        "INT2" => row
            .try_get::<Option<i16>, _>(idx)
            .ok()
            .flatten()
            .map(|v| JsonValue::from(i64::from(v)))
            .unwrap_or(JsonValue::Null),
        "INT4" => row
            .try_get::<Option<i32>, _>(idx)
            .ok()
            .flatten()
            .map(|v| JsonValue::from(i64::from(v)))
            .unwrap_or(JsonValue::Null),
        "INT8" => row
            .try_get::<Option<i64>, _>(idx)
            .ok()
            .flatten()
            .map(JsonValue::from)
            .unwrap_or(JsonValue::Null),
        "FLOAT4" => row
            .try_get::<Option<f32>, _>(idx)
            .ok()
            .flatten()
            .map(|v| v as f64)
            .and_then(serde_json::Number::from_f64)
            .map(JsonValue::Number)
            .unwrap_or(JsonValue::Null),
        "FLOAT8" => row
            .try_get::<Option<f64>, _>(idx)
            .ok()
            .flatten()
            .and_then(serde_json::Number::from_f64)
            .map(JsonValue::Number)
            .unwrap_or(JsonValue::Null),
        "DATE" => row
            .try_get::<Option<sqlx::types::chrono::NaiveDate>, _>(idx)
            .ok()
            .flatten()
            .map(|d| JsonValue::from(format!("{}T00:00:00.000Z", d.format("%Y-%m-%d"))))
            .unwrap_or(JsonValue::Null),
        "TIMESTAMP" => row
            .try_get::<Option<sqlx::types::chrono::NaiveDateTime>, _>(idx)
            .ok()
            .flatten()
            .map(|dt| JsonValue::from(format!("{}Z", dt.format("%Y-%m-%dT%H:%M:%S%.f"))))
            .unwrap_or(JsonValue::Null),
        "TIMESTAMPTZ" => row
            .try_get::<Option<sqlx::types::chrono::DateTime<sqlx::types::chrono::Utc>>, _>(idx)
            .ok()
            .flatten()
            .map(|dt| JsonValue::from(dt.to_rfc3339()))
            .unwrap_or(JsonValue::Null),
        _ => {
            if let Ok(v) = row.try_get::<Option<String>, _>(idx) {
                return v.map(JsonValue::from).unwrap_or(JsonValue::Null);
            }
            if let Ok(v) = row.try_get::<Option<i64>, _>(idx) {
                return v.map(JsonValue::from).unwrap_or(JsonValue::Null);
            }
            if let Ok(v) = row.try_get::<Option<i32>, _>(idx) {
                return v.map(|n| JsonValue::from(i64::from(n))).unwrap_or(JsonValue::Null);
            }
            if let Ok(v) = row.try_get::<Option<f64>, _>(idx) {
                return v
                    .and_then(serde_json::Number::from_f64)
                    .map(JsonValue::Number)
                    .unwrap_or(JsonValue::Null);
            }
            if let Ok(v) = row.try_get::<Option<bool>, _>(idx) {
                return v.map(JsonValue::from).unwrap_or(JsonValue::Null);
            }
            JsonValue::Null
        }
    }
}

async fn query_sqlite(
    opts: &sqlx::sqlite::SqliteConnectOptions,
    sql: &str,
    params: &[JsonValue],
) -> Result<SqlQueryResult> {
    // Fetch schema first (best-effort). We do this on a separate connection so schema discovery
    // failures don't prevent query execution.
    let schema = sqlite_schema(opts, sql).await.ok();

    let mut conn = sqlx::SqliteConnection::connect_with(opts)
        .await
        .context("connect sqlite")?;

    let mut query = sqlx::query(sql);
    for value in params {
        query = bind_sqlite_param(query, value);
    }
    let rows = query.fetch_all(&mut conn).await.context("execute sqlite query")?;

    let columns = schema
        .as_ref()
        .map(|s| s.columns.clone())
        .or_else(|| rows.first().map(|r| r.columns().iter().map(|c| c.name().to_string()).collect()))
        .unwrap_or_default();

    let types = schema.as_ref().and_then(|s| s.types.clone());
    let column_types: Vec<SqlDataType> = columns
        .iter()
        .map(|name| {
            types
                .as_ref()
                .and_then(|m| m.get(name))
                .cloned()
                .unwrap_or(SqlDataType::Any)
        })
        .collect();

    let mut out_rows = Vec::with_capacity(rows.len());
    for row in rows {
        let mut out = Vec::with_capacity(columns.len());
        for idx in 0..columns.len() {
            out.push(sqlite_cell_to_json(&row, idx, &column_types[idx]));
        }
        out_rows.push(out);
    }

    Ok(SqlQueryResult {
        columns,
        types,
        rows: out_rows,
    })
}

async fn query_postgres(
    opts: &sqlx::postgres::PgConnectOptions,
    sql: &str,
    params: &[JsonValue],
) -> Result<SqlQueryResult> {
    let schema = postgres_schema(opts, sql).await.ok();

    let mut conn = sqlx::PgConnection::connect_with(opts)
        .await
        .context("connect postgres")?;

    let mut query = sqlx::query(sql);
    for value in params {
        query = bind_postgres_param(query, value);
    }
    let rows = query
        .fetch_all(&mut conn)
        .await
        .context("execute postgres query")?;

    let columns = schema
        .as_ref()
        .map(|s| s.columns.clone())
        .or_else(|| rows.first().map(|r| r.columns().iter().map(|c| c.name().to_string()).collect()))
        .unwrap_or_default();

    let types = schema.as_ref().and_then(|s| s.types.clone());

    let mut out_rows = Vec::with_capacity(rows.len());
    for row in rows {
        let mut out = Vec::with_capacity(columns.len());
        for idx in 0..columns.len() {
            let type_name = row
                .columns()
                .get(idx)
                .map(|c| c.type_info().name())
                .unwrap_or("");
            out.push(postgres_cell_to_json(&row, idx, type_name));
        }
        out_rows.push(out);
    }

    Ok(SqlQueryResult {
        columns,
        types,
        rows: out_rows,
    })
}

pub async fn sql_query(
    connection: JsonValue,
    sql: String,
    params: Vec<JsonValue>,
    credentials: Option<JsonValue>,
) -> Result<SqlQueryResult> {
    let kind = connection_kind(&connection)?;

    match kind.as_str() {
        "sqlite" => {
            let descriptor: SqliteConnectionDescriptor =
                serde_json::from_value(connection).context("invalid sqlite connection descriptor")?;
            let in_memory = descriptor.in_memory.unwrap_or(false);
            let mut opts = if in_memory {
                sqlx::sqlite::SqliteConnectOptions::from_str("sqlite::memory:")?
            } else {
                let path = descriptor
                    .path
                    .ok_or_else(|| anyhow!("sqlite connection requires `path`"))?;
                sqlx::sqlite::SqliteConnectOptions::new().filename(path)
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
                } else {
                    sqlx::sqlite::SqliteConnectOptions::new().filename(path)
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
    let kind = connection_kind(&connection)?;

    match kind.as_str() {
        "sqlite" => {
            let descriptor: SqliteConnectionDescriptor =
                serde_json::from_value(connection).context("invalid sqlite connection descriptor")?;
            let in_memory = descriptor.in_memory.unwrap_or(false);
            let mut opts = if in_memory {
                sqlx::sqlite::SqliteConnectOptions::from_str("sqlite::memory:")?
            } else {
                let path = descriptor
                    .path
                    .ok_or_else(|| anyhow!("sqlite connection requires `path`"))?;
                sqlx::sqlite::SqliteConnectOptions::new().filename(path)
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
                let in_memory = path.trim().eq_ignore_ascii_case(":memory:") || path.trim().eq_ignore_ascii_case("memory");
                let mut opts = if in_memory {
                    sqlx::sqlite::SqliteConnectOptions::from_str("sqlite::memory:")?
                } else {
                    sqlx::sqlite::SqliteConnectOptions::new().filename(path)
                };
                opts = opts.create_if_missing(false);
                sqlite_schema(&opts, &sql).await
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

                postgres_schema(&opts, &sql).await
            } else {
                Err(anyhow!(
                    "Unsupported ODBC driver '{driver}' (supported: SQLite, PostgreSQL)"
                ))
            }
        }
        other => Err(anyhow!(
            "Unsupported SQL connection kind '{other}' (supported: sqlite, postgres, odbc)"
        )),
    }
}
