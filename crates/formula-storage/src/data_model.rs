use crate::storage::{Result, StorageError};
use rusqlite::{params, Connection, OptionalExtension, Transaction};
use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::sync::Arc;
use uuid::Uuid;

/// Schema-only view of a persisted DAX data model.
///
/// This is intentionally decoupled from [`formula_dax::DataModel`] so callers can inspect the
/// model (tables/columns/measures/relationships) without loading large column chunks.
#[derive(Debug, Clone, PartialEq)]
pub struct DataModelSchema {
    pub tables: Vec<DataModelTableSchema>,
    pub relationships: Vec<formula_dax::Relationship>,
    pub measures: Vec<DataModelMeasureSchema>,
    pub calculated_columns: Vec<DataModelCalculatedColumnSchema>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct DataModelTableSchema {
    pub name: String,
    pub row_count: usize,
    pub columns: Vec<DataModelColumnSchema>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct DataModelColumnSchema {
    pub name: String,
    pub column_type: formula_columnar::ColumnType,
}

#[derive(Debug, Clone, PartialEq)]
pub struct DataModelMeasureSchema {
    pub name: String,
    pub expression: String,
}

#[derive(Debug, Clone, PartialEq)]
pub struct DataModelCalculatedColumnSchema {
    pub table: String,
    pub name: String,
    pub expression: String,
}

#[derive(Debug, Clone, PartialEq)]
pub struct DataModelChunk {
    pub chunk_index: usize,
    pub kind: DataModelChunkKind,
    pub data: Vec<u8>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DataModelChunkKind {
    Int,
    Float,
    Bool,
    Dict,
}

impl DataModelChunkKind {
    fn as_str(self) -> &'static str {
        match self {
            DataModelChunkKind::Int => "int",
            DataModelChunkKind::Float => "float",
            DataModelChunkKind::Bool => "bool",
            DataModelChunkKind::Dict => "dict",
        }
    }

    fn parse(raw: &str) -> Option<Self> {
        Some(match raw {
            "int" => DataModelChunkKind::Int,
            "float" => DataModelChunkKind::Float,
            "bool" => DataModelChunkKind::Bool,
            "dict" => DataModelChunkKind::Dict,
            _ => return None,
        })
    }
}

#[derive(Debug, Serialize, Deserialize)]
struct TableSchemaV1 {
    version: u32,
    page_size_rows: usize,
    cache_max_entries: usize,
}

#[derive(Debug, Serialize, Deserialize)]
struct ColumnEncodingV1 {
    version: u32,
    chunk_format_version: u32,
    dictionary_format_version: u32,
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize)]
#[serde(tag = "type", rename_all = "lowercase")]
enum ColumnTypeV1 {
    Number,
    String,
    Boolean,
    Datetime,
    Currency { scale: u8 },
    Percentage { scale: u8 },
}

impl From<formula_columnar::ColumnType> for ColumnTypeV1 {
    fn from(value: formula_columnar::ColumnType) -> Self {
        match value {
            formula_columnar::ColumnType::Number => ColumnTypeV1::Number,
            formula_columnar::ColumnType::String => ColumnTypeV1::String,
            formula_columnar::ColumnType::Boolean => ColumnTypeV1::Boolean,
            formula_columnar::ColumnType::DateTime => ColumnTypeV1::Datetime,
            formula_columnar::ColumnType::Currency { scale } => ColumnTypeV1::Currency { scale },
            formula_columnar::ColumnType::Percentage { scale } => {
                ColumnTypeV1::Percentage { scale }
            }
        }
    }
}

impl From<ColumnTypeV1> for formula_columnar::ColumnType {
    fn from(value: ColumnTypeV1) -> Self {
        match value {
            ColumnTypeV1::Number => formula_columnar::ColumnType::Number,
            ColumnTypeV1::String => formula_columnar::ColumnType::String,
            ColumnTypeV1::Boolean => formula_columnar::ColumnType::Boolean,
            ColumnTypeV1::Datetime => formula_columnar::ColumnType::DateTime,
            ColumnTypeV1::Currency { scale } => formula_columnar::ColumnType::Currency { scale },
            ColumnTypeV1::Percentage { scale } => {
                formula_columnar::ColumnType::Percentage { scale }
            }
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "type", content = "value", rename_all = "lowercase")]
enum ColumnarValueV1 {
    Null,
    Number(f64),
    Boolean(bool),
    String(String),
    Datetime(i64),
    Currency(i64),
    Percentage(i64),
}

impl From<&formula_columnar::Value> for ColumnarValueV1 {
    fn from(value: &formula_columnar::Value) -> Self {
        match value {
            formula_columnar::Value::Null => ColumnarValueV1::Null,
            formula_columnar::Value::Number(v) => ColumnarValueV1::Number(*v),
            formula_columnar::Value::Boolean(v) => ColumnarValueV1::Boolean(*v),
            formula_columnar::Value::String(v) => ColumnarValueV1::String(v.as_ref().to_string()),
            formula_columnar::Value::DateTime(v) => ColumnarValueV1::Datetime(*v),
            formula_columnar::Value::Currency(v) => ColumnarValueV1::Currency(*v),
            formula_columnar::Value::Percentage(v) => ColumnarValueV1::Percentage(*v),
        }
    }
}

impl From<ColumnarValueV1> for formula_columnar::Value {
    fn from(value: ColumnarValueV1) -> Self {
        match value {
            ColumnarValueV1::Null => formula_columnar::Value::Null,
            ColumnarValueV1::Number(v) => formula_columnar::Value::Number(v),
            ColumnarValueV1::Boolean(v) => formula_columnar::Value::Boolean(v),
            ColumnarValueV1::String(v) => formula_columnar::Value::String(Arc::<str>::from(v)),
            ColumnarValueV1::Datetime(v) => formula_columnar::Value::DateTime(v),
            ColumnarValueV1::Currency(v) => formula_columnar::Value::Currency(v),
            ColumnarValueV1::Percentage(v) => formula_columnar::Value::Percentage(v),
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
struct ColumnStatsV1 {
    distinct_count: u64,
    null_count: u64,
    min: Option<ColumnarValueV1>,
    max: Option<ColumnarValueV1>,
    sum: Option<f64>,
    avg_length: Option<f64>,
}

impl From<&formula_columnar::ColumnStats> for ColumnStatsV1 {
    fn from(value: &formula_columnar::ColumnStats) -> Self {
        Self {
            distinct_count: value.distinct_count,
            null_count: value.null_count,
            min: value.min.as_ref().map(ColumnarValueV1::from),
            max: value.max.as_ref().map(ColumnarValueV1::from),
            sum: value.sum,
            avg_length: value.avg_length,
        }
    }
}

impl ColumnStatsV1 {
    fn into_column_stats(
        self,
        column_type: formula_columnar::ColumnType,
    ) -> formula_columnar::ColumnStats {
        formula_columnar::ColumnStats {
            column_type,
            distinct_count: self.distinct_count,
            null_count: self.null_count,
            min: self.min.map(Into::into),
            max: self.max.map(Into::into),
            sum: self.sum,
            avg_length: self.avg_length,
        }
    }
}

const CHUNK_BLOB_VERSION: u8 = 1;
const DICTIONARY_BLOB_VERSION: u8 = 1;

pub(crate) fn save_data_model_tx(
    tx: &Transaction<'_>,
    workbook_id: Uuid,
    model: &formula_dax::DataModel,
) -> Result<()> {
    let workbook_id_str = workbook_id.to_string();

    // Replace the existing persisted model for this workbook.
    tx.execute(
        "DELETE FROM data_model_tables WHERE workbook_id = ?1",
        params![&workbook_id_str],
    )?;
    tx.execute(
        "DELETE FROM data_model_relationships WHERE workbook_id = ?1",
        params![&workbook_id_str],
    )?;
    tx.execute(
        "DELETE FROM data_model_measures WHERE workbook_id = ?1",
        params![&workbook_id_str],
    )?;
    tx.execute(
        "DELETE FROM data_model_calculated_columns WHERE workbook_id = ?1",
        params![&workbook_id_str],
    )?;

    for table in model.tables() {
        let columnar: Arc<formula_columnar::ColumnarTable> = match table.columnar_table() {
            Some(existing) => existing.clone(),
            None => Arc::new(build_columnar_from_dax_table(table)),
        };
        let columnar = columnar.as_ref();

        let options = columnar.options();
        let schema_json = serde_json::to_string(&TableSchemaV1 {
            version: 1,
            page_size_rows: options.page_size_rows,
            cache_max_entries: options.cache.max_entries,
        })?;

        tx.execute(
            r#"
            INSERT INTO data_model_tables (workbook_id, name, schema_json, row_count, metadata)
            VALUES (?1, ?2, ?3, ?4, NULL)
            "#,
            params![
                &workbook_id_str,
                table.name(),
                schema_json,
                columnar.row_count() as i64
            ],
        )?;
        let table_id = tx.last_insert_rowid();

        for (ordinal, col_schema) in columnar.schema().iter().enumerate() {
            let column_type_json =
                serde_json::to_string(&ColumnTypeV1::from(col_schema.column_type))?;

            let stats = columnar
                .stats(ordinal)
                .ok_or_else(|| StorageError::Sqlite(rusqlite::Error::InvalidQuery))?;
            let stats_json = serde_json::to_string(&ColumnStatsV1::from(stats))?;

            let encoding_json = serde_json::to_string(&ColumnEncodingV1 {
                version: 1,
                chunk_format_version: CHUNK_BLOB_VERSION as u32,
                dictionary_format_version: DICTIONARY_BLOB_VERSION as u32,
            })?;

            let dict_blob: Option<Vec<u8>> = columnar.dictionary(ordinal).map(encode_dictionary);

            tx.execute(
                r#"
                INSERT INTO data_model_columns (
                  table_id,
                  ordinal,
                  name,
                  column_type,
                  encoding_json,
                  stats_json,
                  dictionary
                ) VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7)
                "#,
                params![
                    table_id,
                    ordinal as i64,
                    &col_schema.name,
                    column_type_json,
                    encoding_json,
                    stats_json,
                    dict_blob
                ],
            )?;
            let column_id = tx.last_insert_rowid();

            let Some(chunks) = columnar.encoded_chunks(ordinal) else {
                continue;
            };
            for (chunk_index, chunk) in chunks.iter().enumerate() {
                let (kind, data) = encode_chunk(chunk);
                tx.execute(
                    r#"
                    INSERT INTO data_model_chunks (column_id, chunk_index, kind, data, metadata)
                    VALUES (?1, ?2, ?3, ?4, NULL)
                    "#,
                    params![column_id, chunk_index as i64, kind.as_str(), data],
                )?;
            }
        }
    }

    for relationship in model.relationships_definitions() {
        tx.execute(
            r#"
            INSERT INTO data_model_relationships (
              workbook_id,
              name,
              from_table,
              from_column,
              to_table,
              to_column,
              cardinality,
              cross_filter_direction,
              is_active,
              referential_integrity
            ) VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9, ?10)
            "#,
            params![
                &workbook_id_str,
                &relationship.name,
                &relationship.from_table,
                &relationship.from_column,
                &relationship.to_table,
                &relationship.to_column,
                cardinality_to_str(relationship.cardinality),
                cross_filter_to_str(relationship.cross_filter_direction),
                if relationship.is_active { 1i64 } else { 0i64 },
                if relationship.enforce_referential_integrity {
                    1i64
                } else {
                    0i64
                },
            ],
        )?;
    }

    for measure in model.measures_definitions() {
        tx.execute(
            r#"
            INSERT INTO data_model_measures (workbook_id, name, expression, metadata)
            VALUES (?1, ?2, ?3, NULL)
            "#,
            params![&workbook_id_str, &measure.name, &measure.expression],
        )?;
    }

    for calc in model.calculated_columns() {
        tx.execute(
            r#"
            INSERT INTO data_model_calculated_columns (workbook_id, table_name, name, expression, metadata)
            VALUES (?1, ?2, ?3, ?4, NULL)
            "#,
            params![&workbook_id_str, &calc.table, &calc.name, &calc.expression],
        )?;
    }

    Ok(())
}

pub(crate) fn load_data_model(
    conn: &Connection,
    workbook_id: Uuid,
) -> Result<formula_dax::DataModel> {
    let workbook_id_str = workbook_id.to_string();
    let mut model = formula_dax::DataModel::new();

    let mut table_stmt = conn.prepare(
        r#"
        SELECT id, name, schema_json, row_count
        FROM data_model_tables
        WHERE workbook_id = ?1
        ORDER BY id
        "#,
    )?;

    let tables = table_stmt.query_map(params![&workbook_id_str], |row| {
        Ok((
            row.get::<_, i64>(0)?,
            row.get::<_, String>(1)?,
            row.get::<_, Option<String>>(2).ok().flatten(),
            row.get::<_, Option<i64>>(3).ok().flatten(),
        ))
    })?;

    for table_row in tables {
        let Ok((table_id, table_name, schema_json, row_count)) = table_row else {
            continue;
        };
        let row_count = row_count.and_then(|count| usize::try_from(count).ok()).unwrap_or(0);
        let schema = schema_json
            .as_deref()
            .and_then(|raw| serde_json::from_str::<TableSchemaV1>(raw).ok());
        // `page_size_rows` is persisted as part of the table schema so we can map a row index to
        // its encoded chunk (`chunk_idx = row / page_size_rows`). If this value is corrupted (e.g.
        // `0` or mismatched with the persisted chunks), downstream accessors in `formula-columnar`
        // can panic on division/modulo by zero or out-of-bounds indexing.
        //
        // Prefer inferring a usable page size from the persisted chunks (best-effort), falling
        // back to the schema value when no chunks are available.
        let mut page_size_rows = schema.as_ref().map(|s| s.page_size_rows).unwrap_or(0);
        let cache_max_entries = schema
            .as_ref()
            .map(|s| s.cache_max_entries)
            .unwrap_or(formula_columnar::PageCacheConfig::default().max_entries);

        let mut col_stmt = conn.prepare(
            r#"
            SELECT id, name, column_type, encoding_json, stats_json, dictionary
            FROM data_model_columns
            WHERE table_id = ?1
            ORDER BY ordinal
            "#,
        )?;

        struct LoadedColumn {
            schema: formula_columnar::ColumnSchema,
            chunks: Vec<formula_columnar::EncodedChunk>,
            stats: formula_columnar::ColumnStats,
            dictionary: Option<Arc<Vec<Arc<str>>>>,
            max_chunk_len: usize,
        }
        let mut loaded_columns: Vec<LoadedColumn> = Vec::new();

        let cols = col_stmt.query_map(params![table_id], |row| {
            Ok((
                row.get::<_, i64>(0)?,
                row.get::<_, String>(1)?,
                row.get::<_, Option<String>>(2).ok().flatten(),
                row.get::<_, Option<String>>(4).ok().flatten(),
                row.get::<_, Option<Vec<u8>>>(5).ok().flatten(),
            ))
        })?;

        for col_row in cols {
            let Ok((column_id, name, column_type_raw, stats_json, dictionary_blob)) = col_row
            else {
                continue;
            };
            let Some(column_type_raw) = column_type_raw else {
                continue;
            };
            let Some(stats_json) = stats_json else {
                continue;
            };
            let column_type: ColumnTypeV1 = match serde_json::from_str(&column_type_raw) {
                Ok(column_type) => column_type,
                Err(_) => continue,
            };
            let column_type: formula_columnar::ColumnType = column_type.into();

            let stats_v1: ColumnStatsV1 = match serde_json::from_str(&stats_json) {
                Ok(stats) => stats,
                Err(_) => continue,
            };
            let stats = stats_v1.into_column_stats(column_type);

            let dictionary = match dictionary_blob.map(decode_dictionary).transpose() {
                Ok(dict) => dict,
                Err(_) => continue,
            };

            let mut chunk_stmt = conn.prepare(
                r#"
                SELECT chunk_index, kind, data
                FROM data_model_chunks
                WHERE column_id = ?1
                ORDER BY chunk_index
                "#,
            )?;

            let mut chunks: Vec<formula_columnar::EncodedChunk> = Vec::new();
            let mut max_chunk_len = 0usize;
            let rows_iter = match chunk_stmt.query_map(params![column_id], |row| {
                Ok((
                    row.get::<_, i64>(0)?,
                    row.get::<_, String>(1)?,
                    row.get::<_, Vec<u8>>(2)?,
                ))
            }) {
                Ok(iter) => iter,
                Err(_) => continue,
            };
            let mut chunk_failed = false;
            for chunk_row in rows_iter {
                let Ok((chunk_index, kind_raw, data)) = chunk_row else {
                    chunk_failed = true;
                    break;
                };
                let Ok(chunk_index) = usize::try_from(chunk_index) else {
                    chunk_failed = true;
                    break;
                };
                if chunk_index != chunks.len() {
                    chunk_failed = true;
                    break;
                }
                let Some(kind) = DataModelChunkKind::parse(&kind_raw) else {
                    chunk_failed = true;
                    break;
                };
                let decoded = match decode_chunk(kind, &data) {
                    Ok(decoded) => decoded,
                    Err(_) => {
                        chunk_failed = true;
                        break;
                    }
                };
                max_chunk_len = max_chunk_len.max(decoded.len());
                chunks.push(decoded);
            }
            if chunk_failed {
                continue;
            }

            let col_schema = formula_columnar::ColumnSchema {
                name: name.clone(),
                column_type,
            };
            loaded_columns.push(LoadedColumn {
                schema: col_schema,
                chunks,
                stats,
                dictionary,
                max_chunk_len,
            });
        }

        if loaded_columns.is_empty() {
            continue;
        }

        // Pick the most common maximum chunk length as the inferred page size. This avoids being
        // skewed by a single corrupted column that reports an absurd chunk length.
        let mut page_size_counts: HashMap<usize, usize> = HashMap::new();
        for col in &loaded_columns {
            if col.max_chunk_len > 0 {
                *page_size_counts.entry(col.max_chunk_len).or_insert(0) += 1;
            }
        }
        if !page_size_counts.is_empty() {
            let mut best_size = usize::MAX;
            let mut best_count = 0usize;
            for (size, count) in page_size_counts {
                // Prefer the most common chunk length. When there's a tie, pick the smaller page
                // size: it is more conservative (and less likely to cause other columns to be
                // rejected due to a single outlier chunk reporting an inflated length).
                if count > best_count || (count == best_count && size < best_size) {
                    best_size = size;
                    best_count = count;
                }
            }
            page_size_rows = best_size;
        }
        if page_size_rows == 0 {
            page_size_rows = 1;
        }

        fn max_rows_safe_for_chunks(
            chunks: &[formula_columnar::EncodedChunk],
            page_size_rows: usize,
        ) -> Option<usize> {
            if page_size_rows == 0 || chunks.is_empty() {
                return None;
            }
            if chunks.len() > 1 {
                for chunk in &chunks[..chunks.len() - 1] {
                    if chunk.len() != page_size_rows {
                        return None;
                    }
                }
            }
            let last_len = chunks.last()?.len();
            if last_len > page_size_rows {
                return None;
            }
            Some((chunks.len() - 1).saturating_mul(page_size_rows).saturating_add(last_len))
        }

        // Clamp the table row count to the largest safe chunk-derived row count so we don't create
        // a columnar table that can panic on out-of-bounds accesses when the persisted row count is
        // corrupted.
        let mut max_rows_any = 0usize;
        for col in &loaded_columns {
            if col.chunks.is_empty() {
                continue;
            }
            let Some(rows) = max_rows_safe_for_chunks(&col.chunks, page_size_rows) else {
                continue;
            };
            max_rows_any = max_rows_any.max(rows);
        }

        let row_count = if max_rows_any == 0 {
            0
        } else if row_count == 0 {
            max_rows_any
        } else {
            row_count.min(max_rows_any)
        };

        let mut schema_out = Vec::new();
        let mut columns_out = Vec::new();

        for col in loaded_columns {
            if !col.chunks.is_empty() {
                let Some(max_rows) = max_rows_safe_for_chunks(&col.chunks, page_size_rows) else {
                    continue;
                };
                if max_rows < row_count {
                    continue;
                }
            }

            schema_out.push(col.schema.clone());
            columns_out.push(formula_columnar::EncodedColumn {
                schema: col.schema,
                chunks: col.chunks,
                stats: col.stats,
                dictionary: col.dictionary,
            });
        }

        if schema_out.is_empty() {
            continue;
        }
        let options = formula_columnar::TableOptions {
            page_size_rows,
            cache: formula_columnar::PageCacheConfig {
                max_entries: cache_max_entries,
            },
        };

        let table =
            formula_columnar::ColumnarTable::from_encoded(schema_out, columns_out, row_count, options);
        if model
            .add_table(formula_dax::Table::from_columnar(table_name, table))
            .is_err()
        {
            continue;
        }
    }

    // Relationships.
    let mut rel_stmt = conn.prepare(
        r#"
        SELECT
          name,
          from_table,
          from_column,
          to_table,
          to_column,
          cardinality,
          cross_filter_direction,
          is_active,
          referential_integrity
        FROM data_model_relationships
        WHERE workbook_id = ?1
        ORDER BY id
        "#,
    )?;
    let rels = rel_stmt.query_map(params![&workbook_id_str], |row| {
        Ok((
            row.get::<_, String>(0)?,
            row.get::<_, String>(1)?,
            row.get::<_, String>(2)?,
            row.get::<_, String>(3)?,
            row.get::<_, String>(4)?,
            row.get::<_, String>(5)?,
            row.get::<_, String>(6)?,
            row.get::<_, i64>(7)?,
            row.get::<_, i64>(8)?,
        ))
    })?;

    for rel_row in rels {
        let Ok((
            name,
            from_table,
            from_column,
            to_table,
            to_column,
            card_raw,
            dir_raw,
            active,
            ri,
        )) = rel_row
        else {
            continue;
        };
        let Ok(cardinality) = parse_cardinality(&card_raw) else {
            continue;
        };
        let Ok(cross_filter_direction) = parse_cross_filter_direction(&dir_raw) else {
            continue;
        };
        let relationship = formula_dax::Relationship {
            name,
            from_table,
            from_column,
            to_table,
            to_column,
            cardinality,
            cross_filter_direction,
            is_active: active != 0,
            enforce_referential_integrity: ri != 0,
        };
        let _ = model.add_relationship(relationship);
    }

    // Measures.
    let mut measure_stmt = conn.prepare(
        r#"
        SELECT name, expression
        FROM data_model_measures
        WHERE workbook_id = ?1
        ORDER BY id
        "#,
    )?;
    let measures = measure_stmt.query_map(params![&workbook_id_str], |row| {
        Ok((row.get::<_, String>(0)?, row.get::<_, String>(1)?))
    })?;
    for row in measures {
        let Ok((name, expr)) = row else {
            continue;
        };
        let _ = model.add_measure(name, expr);
    }

    // Calculated columns (definition only; values are stored in the table data).
    let mut calc_stmt = conn.prepare(
        r#"
        SELECT table_name, name, expression
        FROM data_model_calculated_columns
        WHERE workbook_id = ?1
        ORDER BY id
        "#,
    )?;
    let calcs = calc_stmt.query_map(params![&workbook_id_str], |row| {
        Ok((
            row.get::<_, String>(0)?,
            row.get::<_, String>(1)?,
            row.get::<_, String>(2)?,
        ))
    })?;
    for row in calcs {
        let Ok((table, name, expr)) = row else {
            continue;
        };
        let _ = model.add_calculated_column_definition(table, name, expr);
    }

    Ok(model)
}

pub(crate) fn load_data_model_schema(
    conn: &Connection,
    workbook_id: Uuid,
) -> Result<DataModelSchema> {
    let workbook_id_str = workbook_id.to_string();

    let mut table_stmt = conn.prepare(
        r#"
        SELECT id, name, row_count
        FROM data_model_tables
        WHERE workbook_id = ?1
        ORDER BY id
        "#,
    )?;
    let table_rows = table_stmt.query_map(params![&workbook_id_str], |row| {
        Ok((
            row.get::<_, i64>(0)?,
            row.get::<_, String>(1)?,
            row.get::<_, Option<i64>>(2).ok().flatten(),
        ))
    })?;

    let mut tables = Vec::new();
    for row in table_rows {
        let Ok((table_id, name, row_count)) = row else {
            continue;
        };
        let row_count = row_count
            .and_then(|count| usize::try_from(count).ok())
            .unwrap_or(0);
        let mut col_stmt = conn.prepare(
            r#"
            SELECT name, column_type
            FROM data_model_columns
            WHERE table_id = ?1
            ORDER BY ordinal
            "#,
        )?;
        let col_rows = col_stmt.query_map(params![table_id], |row| {
            Ok((
                row.get::<_, Option<String>>(0).ok().flatten(),
                row.get::<_, Option<String>>(1).ok().flatten(),
            ))
        })?;
        let mut columns = Vec::new();
        for col in col_rows {
            let Ok((name, column_type_raw)) = col else {
                continue;
            };
            let (Some(name), Some(column_type_raw)) = (name, column_type_raw) else {
                continue;
            };
            let column_type: ColumnTypeV1 = match serde_json::from_str(&column_type_raw) {
                Ok(column_type) => column_type,
                Err(_) => continue,
            };
            columns.push(DataModelColumnSchema {
                name,
                column_type: column_type.into(),
            });
        }
        tables.push(DataModelTableSchema {
            name,
            row_count,
            columns,
        });
    }

    let mut rel_stmt = conn.prepare(
        r#"
        SELECT
          name,
          from_table,
          from_column,
          to_table,
          to_column,
          cardinality,
          cross_filter_direction,
          is_active,
          referential_integrity
        FROM data_model_relationships
        WHERE workbook_id = ?1
        ORDER BY id
        "#,
    )?;
    let rel_rows = rel_stmt.query_map(params![&workbook_id_str], |row| {
        Ok((
            row.get::<_, Option<String>>(0).ok().flatten(),
            row.get::<_, Option<String>>(1).ok().flatten(),
            row.get::<_, Option<String>>(2).ok().flatten(),
            row.get::<_, Option<String>>(3).ok().flatten(),
            row.get::<_, Option<String>>(4).ok().flatten(),
            row.get::<_, Option<String>>(5).ok().flatten(),
            row.get::<_, Option<String>>(6).ok().flatten(),
            row.get::<_, Option<i64>>(7).ok().flatten(),
            row.get::<_, Option<i64>>(8).ok().flatten(),
        ))
    })?;
    let mut relationships = Vec::new();
    for row in rel_rows {
        let Ok((name, from_table, from_column, to_table, to_column, card, dir, active, ri)) = row else {
            continue;
        };
        let (
            Some(name),
            Some(from_table),
            Some(from_column),
            Some(to_table),
            Some(to_column),
            Some(card),
            Some(dir),
        ) = (name, from_table, from_column, to_table, to_column, card, dir)
        else {
            continue;
        };
        let active = active.unwrap_or(0);
        let ri = ri.unwrap_or(0);
        let Ok(cardinality) = parse_cardinality(&card) else {
            continue;
        };
        let Ok(cross_filter_direction) = parse_cross_filter_direction(&dir) else {
            continue;
        };
        relationships.push(formula_dax::Relationship {
            name,
            from_table,
            from_column,
            to_table,
            to_column,
            cardinality,
            cross_filter_direction,
            is_active: active != 0,
            enforce_referential_integrity: ri != 0,
        });
    }

    let mut measure_stmt = conn.prepare(
        r#"
        SELECT name, expression
        FROM data_model_measures
        WHERE workbook_id = ?1
        ORDER BY id
        "#,
    )?;
    let measure_rows = measure_stmt.query_map(params![&workbook_id_str], |row| {
        Ok(DataModelMeasureSchema {
            name: row.get(0).unwrap_or_default(),
            expression: row.get(1).unwrap_or_default(),
        })
    })?;
    let mut measures = Vec::new();
    for row in measure_rows {
        let Ok(row) = row else {
            continue;
        };
        measures.push(row);
    }

    let mut calc_stmt = conn.prepare(
        r#"
        SELECT table_name, name, expression
        FROM data_model_calculated_columns
        WHERE workbook_id = ?1
        ORDER BY id
        "#,
    )?;
    let calc_rows = calc_stmt.query_map(params![&workbook_id_str], |row| {
        Ok(DataModelCalculatedColumnSchema {
            table: row.get(0).unwrap_or_default(),
            name: row.get(1).unwrap_or_default(),
            expression: row.get(2).unwrap_or_default(),
        })
    })?;
    let mut calculated_columns = Vec::new();
    for row in calc_rows {
        let Ok(row) = row else {
            continue;
        };
        calculated_columns.push(row);
    }

    Ok(DataModelSchema {
        tables,
        relationships,
        measures,
        calculated_columns,
    })
}

pub(crate) fn stream_column_chunks<F>(
    conn: &Connection,
    workbook_id: Uuid,
    table_name: &str,
    column_name: &str,
    mut f: F,
) -> Result<()>
where
    F: FnMut(DataModelChunk) -> Result<()>,
{
    let workbook_id_str = workbook_id.to_string();
    let table_id: Option<i64> = conn
        .query_row(
            "SELECT id FROM data_model_tables WHERE workbook_id = ?1 AND name = ?2",
            params![&workbook_id_str, table_name],
            |r| r.get(0),
        )
        .optional()?;
    let Some(table_id) = table_id else {
        return Ok(());
    };

    let column_id: Option<i64> = conn
        .query_row(
            "SELECT id FROM data_model_columns WHERE table_id = ?1 AND name = ?2",
            params![table_id, column_name],
            |r| r.get(0),
        )
        .optional()?;
    let Some(column_id) = column_id else {
        return Ok(());
    };

    let mut stmt = conn.prepare(
        r#"
        SELECT chunk_index, kind, data
        FROM data_model_chunks
        WHERE column_id = ?1
        ORDER BY chunk_index
        "#,
    )?;
    let rows = stmt.query_map(params![column_id], |row| {
        Ok((
            row.get::<_, i64>(0)?,
            row.get::<_, String>(1)?,
            row.get::<_, Vec<u8>>(2)?,
        ))
    })?;

    for row in rows {
        let (chunk_index, kind_raw, data) = match row {
            Ok(row) => row,
            Err(_) => continue,
        };
        let Ok(chunk_index) = usize::try_from(chunk_index) else {
            continue;
        };
        let Some(kind) = DataModelChunkKind::parse(&kind_raw) else {
            continue;
        };
        f(DataModelChunk {
            chunk_index,
            kind,
            data,
        })?;
    }
    Ok(())
}

fn cardinality_to_str(card: formula_dax::Cardinality) -> &'static str {
    match card {
        formula_dax::Cardinality::OneToMany => "one_to_many",
        formula_dax::Cardinality::OneToOne => "one_to_one",
        formula_dax::Cardinality::ManyToMany => "many_to_many",
    }
}

fn cross_filter_to_str(dir: formula_dax::CrossFilterDirection) -> &'static str {
    match dir {
        formula_dax::CrossFilterDirection::Single => "single",
        formula_dax::CrossFilterDirection::Both => "both",
    }
}

fn parse_cardinality(raw: &str) -> Result<formula_dax::Cardinality> {
    Ok(match raw {
        "one_to_many" => formula_dax::Cardinality::OneToMany,
        "one_to_one" => formula_dax::Cardinality::OneToOne,
        "many_to_many" => formula_dax::Cardinality::ManyToMany,
        _ => return Err(StorageError::Sqlite(rusqlite::Error::InvalidQuery)),
    })
}

fn parse_cross_filter_direction(raw: &str) -> Result<formula_dax::CrossFilterDirection> {
    Ok(match raw {
        "single" => formula_dax::CrossFilterDirection::Single,
        "both" => formula_dax::CrossFilterDirection::Both,
        _ => return Err(StorageError::Sqlite(rusqlite::Error::InvalidQuery)),
    })
}

fn build_columnar_from_dax_table(table: &formula_dax::Table) -> formula_columnar::ColumnarTable {
    use formula_columnar::{ColumnSchema, ColumnType, ColumnarTableBuilder, TableOptions, Value};

    let mut inferred: Vec<Option<ColumnType>> = vec![None; table.columns().len()];
    for col_idx in 0..table.columns().len() {
        let mut seen_number = false;
        let mut seen_bool = false;
        let mut seen_text = false;
        for row in 0..table.row_count() {
            let v = formula_dax::TableBackend::value_by_idx(table, row, col_idx)
                .unwrap_or(formula_dax::Value::Blank);
            match v {
                formula_dax::Value::Blank => {}
                formula_dax::Value::Number(_) => seen_number = true,
                formula_dax::Value::Boolean(_) => seen_bool = true,
                formula_dax::Value::Text(_) => seen_text = true,
            }
            if seen_text {
                break;
            }
        }
        inferred[col_idx] = Some(if seen_text {
            ColumnType::String
        } else if seen_number {
            ColumnType::Number
        } else if seen_bool {
            ColumnType::Boolean
        } else {
            ColumnType::String
        });
    }

    let schema: Vec<ColumnSchema> = table
        .columns()
        .iter()
        .enumerate()
        .map(|(idx, name)| ColumnSchema {
            name: name.clone(),
            column_type: inferred[idx].unwrap_or(ColumnType::String),
        })
        .collect();

    let options = TableOptions::default();
    let mut builder = ColumnarTableBuilder::new(schema, options);

    let mut row_buf: Vec<Value> = Vec::new();
    for row in 0..table.row_count() {
        row_buf.clear();
        for col_idx in 0..table.columns().len() {
            let v = formula_dax::TableBackend::value_by_idx(table, row, col_idx)
                .unwrap_or(formula_dax::Value::Blank);
            row_buf.push(match v {
                formula_dax::Value::Blank => Value::Null,
                formula_dax::Value::Number(n) => Value::Number(n.0),
                formula_dax::Value::Boolean(b) => Value::Boolean(b),
                formula_dax::Value::Text(s) => Value::String(s),
            });
        }
        builder.append_row(&row_buf);
    }

    builder.finalize()
}

fn encode_dictionary(dict: Arc<Vec<Arc<str>>>) -> Vec<u8> {
    let mut out = Vec::new();
    out.push(DICTIONARY_BLOB_VERSION);
    out.extend_from_slice(&(dict.len() as u32).to_le_bytes());
    for s in dict.iter() {
        let bytes = s.as_bytes();
        out.extend_from_slice(&(bytes.len() as u32).to_le_bytes());
        out.extend_from_slice(bytes);
    }
    out
}

fn decode_dictionary(bytes: Vec<u8>) -> Result<Arc<Vec<Arc<str>>>> {
    let mut cursor = Cursor::new(&bytes);
    let version = cursor.read_u8()?;
    if version != DICTIONARY_BLOB_VERSION {
        return Err(StorageError::Sqlite(rusqlite::Error::InvalidQuery));
    }
    let count = cursor.read_u32()? as usize;
    // Guard against corrupted headers that claim an absurd number of entries, which would
    // otherwise attempt to allocate huge vectors.
    let max_possible_entries = cursor.remaining() / 4;
    if count > max_possible_entries {
        return Err(StorageError::Sqlite(rusqlite::Error::InvalidQuery));
    }
    let mut out: Vec<Arc<str>> = Vec::new();
    let _ = out.try_reserve_exact(count);
    for _ in 0..count {
        let len = cursor.read_u32()? as usize;
        let raw = cursor.read_bytes(len)?;
        let s = std::str::from_utf8(raw)
            .map_err(|_| StorageError::Sqlite(rusqlite::Error::InvalidQuery))?;
        out.push(Arc::<str>::from(s));
    }
    Ok(Arc::new(out))
}

fn encode_chunk(chunk: &formula_columnar::EncodedChunk) -> (DataModelChunkKind, Vec<u8>) {
    let mut out = Vec::new();
    out.push(CHUNK_BLOB_VERSION);

    match chunk {
        formula_columnar::EncodedChunk::Int(c) => {
            out.extend_from_slice(&c.min.to_le_bytes());
            out.extend_from_slice(&(c.len as u32).to_le_bytes());
            encode_u64_sequence(&mut out, &c.offsets);
            encode_validity(&mut out, c.validity.as_ref(), c.len);
            (DataModelChunkKind::Int, out)
        }
        formula_columnar::EncodedChunk::Float(c) => {
            out.extend_from_slice(&(c.values.len() as u32).to_le_bytes());
            for v in &c.values {
                out.extend_from_slice(&v.to_le_bytes());
            }
            encode_validity(&mut out, c.validity.as_ref(), c.values.len());
            (DataModelChunkKind::Float, out)
        }
        formula_columnar::EncodedChunk::Bool(c) => {
            out.extend_from_slice(&(c.len as u32).to_le_bytes());
            out.extend_from_slice(&(c.data.len() as u32).to_le_bytes());
            out.extend_from_slice(&c.data);
            encode_validity(&mut out, c.validity.as_ref(), c.len);
            (DataModelChunkKind::Bool, out)
        }
        formula_columnar::EncodedChunk::Dict(c) => {
            out.extend_from_slice(&(c.len as u32).to_le_bytes());
            encode_u32_sequence(&mut out, &c.indices);
            encode_validity(&mut out, c.validity.as_ref(), c.len);
            (DataModelChunkKind::Dict, out)
        }
    }
}

fn decode_chunk(kind: DataModelChunkKind, bytes: &[u8]) -> Result<formula_columnar::EncodedChunk> {
    let mut cursor = Cursor::new(bytes);
    let version = cursor.read_u8()?;
    if version != CHUNK_BLOB_VERSION {
        return Err(StorageError::Sqlite(rusqlite::Error::InvalidQuery));
    }

    fn sanitize_validity(
        validity: Option<formula_columnar::BitVec>,
        len: usize,
    ) -> Option<formula_columnar::BitVec> {
        let Some(bits) = validity else {
            return None;
        };
        let required_words = len.saturating_add(63) / 64;
        if bits.as_words().len() < required_words {
            return None;
        }
        Some(bits)
    }

    fn rle_ends_are_sane(ends: &[u32], len: usize) -> bool {
        if len == 0 {
            return true;
        }
        let Some(&last) = ends.last() else {
            return false;
        };
        if (last as usize) < len {
            return false;
        }
        let mut prev: u32 = 0;
        for &end in ends {
            if end == 0 || end < prev {
                return false;
            }
            prev = end;
        }
        true
    }

    Ok(match kind {
        DataModelChunkKind::Int => {
            let min = cursor.read_i64()?;
            let len = cursor.read_u32()? as usize;
            let offsets = decode_u64_sequence(&mut cursor)?;
            match &offsets {
                formula_columnar::U64SequenceEncoding::Bitpacked { bit_width, .. } => {
                    if *bit_width > 64 {
                        return Err(StorageError::Sqlite(rusqlite::Error::InvalidQuery));
                    }
                }
                formula_columnar::U64SequenceEncoding::Rle(rle) => {
                    if !rle_ends_are_sane(&rle.ends, len) {
                        return Err(StorageError::Sqlite(rusqlite::Error::InvalidQuery));
                    }
                }
            }
            let validity = sanitize_validity(decode_validity(&mut cursor)?, len);
            formula_columnar::EncodedChunk::Int(formula_columnar::ValueEncodedChunk {
                min,
                len,
                offsets,
                validity,
            })
        }
        DataModelChunkKind::Float => {
            let len = cursor.read_u32()? as usize;
            let values_bytes = len
                .checked_mul(8)
                .ok_or_else(|| StorageError::Sqlite(rusqlite::Error::InvalidQuery))?;
            // Need at least 1 more byte for the validity tag.
            if cursor.remaining() < values_bytes.saturating_add(1) {
                return Err(StorageError::Sqlite(rusqlite::Error::InvalidQuery));
            }
            let mut values = Vec::new();
            let _ = values.try_reserve_exact(len);
            for _ in 0..len {
                values.push(cursor.read_f64()?);
            }
            let validity = sanitize_validity(decode_validity(&mut cursor)?, values.len());
            formula_columnar::EncodedChunk::Float(formula_columnar::FloatChunk { values, validity })
        }
        DataModelChunkKind::Bool => {
            let len = cursor.read_u32()? as usize;
            let data_len = cursor.read_u32()? as usize;
            let data = cursor.read_bytes(data_len)?.to_vec();
            let needed_bytes = len.saturating_add(7) / 8;
            if data.len() < needed_bytes {
                return Err(StorageError::Sqlite(rusqlite::Error::InvalidQuery));
            }
            let validity = sanitize_validity(decode_validity(&mut cursor)?, len);
            formula_columnar::EncodedChunk::Bool(formula_columnar::BoolChunk {
                len,
                data,
                validity,
            })
        }
        DataModelChunkKind::Dict => {
            let len = cursor.read_u32()? as usize;
            let indices = decode_u32_sequence(&mut cursor)?;
            match &indices {
                formula_columnar::U32SequenceEncoding::Bitpacked { bit_width, .. } => {
                    if *bit_width > 32 {
                        return Err(StorageError::Sqlite(rusqlite::Error::InvalidQuery));
                    }
                }
                formula_columnar::U32SequenceEncoding::Rle(rle) => {
                    if !rle_ends_are_sane(&rle.ends, len) {
                        return Err(StorageError::Sqlite(rusqlite::Error::InvalidQuery));
                    }
                }
            }
            let validity = sanitize_validity(decode_validity(&mut cursor)?, len);
            formula_columnar::EncodedChunk::Dict(formula_columnar::DictionaryEncodedChunk {
                len,
                indices,
                validity,
            })
        }
    })
}

fn encode_validity(out: &mut Vec<u8>, validity: Option<&formula_columnar::BitVec>, len: usize) {
    match validity {
        None => out.push(0),
        Some(bits) => {
            out.push(1);
            out.extend_from_slice(&(len as u32).to_le_bytes());
            let words = bits.as_words();
            out.extend_from_slice(&(words.len() as u32).to_le_bytes());
            for w in words {
                out.extend_from_slice(&w.to_le_bytes());
            }
        }
    }
}

fn decode_validity(cursor: &mut Cursor<'_>) -> Result<Option<formula_columnar::BitVec>> {
    let tag = cursor.read_u8()?;
    match tag {
        0 => Ok(None),
        1 => {
            let len = cursor.read_u32()? as usize;
            let words_len_raw = cursor.read_u32()? as usize;
            let required_words = len.saturating_add(63) / 64;
            if words_len_raw < required_words {
                return Err(StorageError::Sqlite(rusqlite::Error::InvalidQuery));
            }
            let words_len = required_words;
            let bytes_needed = words_len
                .checked_mul(8)
                .ok_or_else(|| StorageError::Sqlite(rusqlite::Error::InvalidQuery))?;
            if cursor.remaining() < bytes_needed {
                return Err(StorageError::Sqlite(rusqlite::Error::InvalidQuery));
            }
            let mut words = Vec::new();
            let _ = words.try_reserve_exact(words_len);
            for _ in 0..words_len {
                words.push(cursor.read_u64()?);
            }
            Ok(Some(formula_columnar::BitVec::from_words(words, len)))
        }
        _ => Err(StorageError::Sqlite(rusqlite::Error::InvalidQuery)),
    }
}

fn encode_u64_sequence(out: &mut Vec<u8>, seq: &formula_columnar::U64SequenceEncoding) {
    match seq {
        formula_columnar::U64SequenceEncoding::Bitpacked { bit_width, data } => {
            out.push(1);
            out.push(*bit_width);
            out.extend_from_slice(&(data.len() as u32).to_le_bytes());
            out.extend_from_slice(data);
        }
        formula_columnar::U64SequenceEncoding::Rle(rle) => {
            out.push(2);
            out.extend_from_slice(&(rle.values.len() as u32).to_le_bytes());
            for v in &rle.values {
                out.extend_from_slice(&v.to_le_bytes());
            }
            for e in &rle.ends {
                out.extend_from_slice(&e.to_le_bytes());
            }
        }
    }
}

fn decode_u64_sequence(cursor: &mut Cursor<'_>) -> Result<formula_columnar::U64SequenceEncoding> {
    let tag = cursor.read_u8()?;
    Ok(match tag {
        1 => {
            let bit_width = cursor.read_u8()?;
            if bit_width > 64 {
                return Err(StorageError::Sqlite(rusqlite::Error::InvalidQuery));
            }
            let data_len = cursor.read_u32()? as usize;
            let data = cursor.read_bytes(data_len)?.to_vec();
            formula_columnar::U64SequenceEncoding::Bitpacked { bit_width, data }
        }
        2 => {
            let run_count = cursor.read_u32()? as usize;
            let bytes_needed = run_count
                .checked_mul(12)
                .ok_or_else(|| StorageError::Sqlite(rusqlite::Error::InvalidQuery))?;
            if cursor.remaining() < bytes_needed {
                return Err(StorageError::Sqlite(rusqlite::Error::InvalidQuery));
            }
            let mut values = Vec::new();
            let _ = values.try_reserve_exact(run_count);
            for _ in 0..run_count {
                values.push(cursor.read_u64()?);
            }
            let mut ends = Vec::new();
            let _ = ends.try_reserve_exact(run_count);
            for _ in 0..run_count {
                ends.push(cursor.read_u32()?);
            }
            formula_columnar::U64SequenceEncoding::Rle(formula_columnar::RleEncodedU64 {
                values,
                ends,
            })
        }
        _ => return Err(StorageError::Sqlite(rusqlite::Error::InvalidQuery)),
    })
}

fn encode_u32_sequence(out: &mut Vec<u8>, seq: &formula_columnar::U32SequenceEncoding) {
    match seq {
        formula_columnar::U32SequenceEncoding::Bitpacked { bit_width, data } => {
            out.push(1);
            out.push(*bit_width);
            out.extend_from_slice(&(data.len() as u32).to_le_bytes());
            out.extend_from_slice(data);
        }
        formula_columnar::U32SequenceEncoding::Rle(rle) => {
            out.push(2);
            out.extend_from_slice(&(rle.values.len() as u32).to_le_bytes());
            for v in &rle.values {
                out.extend_from_slice(&v.to_le_bytes());
            }
            for e in &rle.ends {
                out.extend_from_slice(&e.to_le_bytes());
            }
        }
    }
}

fn decode_u32_sequence(cursor: &mut Cursor<'_>) -> Result<formula_columnar::U32SequenceEncoding> {
    let tag = cursor.read_u8()?;
    Ok(match tag {
        1 => {
            let bit_width = cursor.read_u8()?;
            if bit_width > 32 {
                return Err(StorageError::Sqlite(rusqlite::Error::InvalidQuery));
            }
            let data_len = cursor.read_u32()? as usize;
            let data = cursor.read_bytes(data_len)?.to_vec();
            formula_columnar::U32SequenceEncoding::Bitpacked { bit_width, data }
        }
        2 => {
            let run_count = cursor.read_u32()? as usize;
            let bytes_needed = run_count
                .checked_mul(8)
                .ok_or_else(|| StorageError::Sqlite(rusqlite::Error::InvalidQuery))?;
            if cursor.remaining() < bytes_needed {
                return Err(StorageError::Sqlite(rusqlite::Error::InvalidQuery));
            }
            let mut values = Vec::new();
            let _ = values.try_reserve_exact(run_count);
            for _ in 0..run_count {
                values.push(cursor.read_u32()?);
            }
            let mut ends = Vec::new();
            let _ = ends.try_reserve_exact(run_count);
            for _ in 0..run_count {
                ends.push(cursor.read_u32()?);
            }
            formula_columnar::U32SequenceEncoding::Rle(formula_columnar::RleEncodedU32 {
                values,
                ends,
            })
        }
        _ => return Err(StorageError::Sqlite(rusqlite::Error::InvalidQuery)),
    })
}

struct Cursor<'a> {
    buf: &'a [u8],
    pos: usize,
}

impl<'a> Cursor<'a> {
    fn new(buf: &'a [u8]) -> Self {
        Self { buf, pos: 0 }
    }

    fn read_bytes(&mut self, len: usize) -> Result<&'a [u8]> {
        let end = self
            .pos
            .checked_add(len)
            .ok_or_else(|| StorageError::Sqlite(rusqlite::Error::InvalidQuery))?;
        let slice = self
            .buf
            .get(self.pos..end)
            .ok_or_else(|| StorageError::Sqlite(rusqlite::Error::InvalidQuery))?;
        self.pos = end;
        Ok(slice)
    }

    fn remaining(&self) -> usize {
        self.buf.len().saturating_sub(self.pos)
    }

    fn read_u8(&mut self) -> Result<u8> {
        Ok(*self
            .read_bytes(1)?
            .first()
            .ok_or_else(|| StorageError::Sqlite(rusqlite::Error::InvalidQuery))?)
    }

    fn read_u32(&mut self) -> Result<u32> {
        let bytes = self.read_bytes(4)?;
        let bytes: [u8; 4] = bytes
            .try_into()
            .map_err(|_| StorageError::Sqlite(rusqlite::Error::InvalidQuery))?;
        Ok(u32::from_le_bytes(bytes))
    }

    fn read_u64(&mut self) -> Result<u64> {
        let bytes = self.read_bytes(8)?;
        let bytes: [u8; 8] = bytes
            .try_into()
            .map_err(|_| StorageError::Sqlite(rusqlite::Error::InvalidQuery))?;
        Ok(u64::from_le_bytes(bytes))
    }

    fn read_i64(&mut self) -> Result<i64> {
        let bytes = self.read_bytes(8)?;
        let bytes: [u8; 8] = bytes
            .try_into()
            .map_err(|_| StorageError::Sqlite(rusqlite::Error::InvalidQuery))?;
        Ok(i64::from_le_bytes(bytes))
    }

    fn read_f64(&mut self) -> Result<f64> {
        let bytes = self.read_bytes(8)?;
        let bytes: [u8; 8] = bytes
            .try_into()
            .map_err(|_| StorageError::Sqlite(rusqlite::Error::InvalidQuery))?;
        Ok(f64::from_le_bytes(bytes))
    }
}
