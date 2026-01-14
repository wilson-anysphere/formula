use js_sys::{Array, Object, Reflect};
use serde::{Deserialize, Serialize};
use serde_json::Value as JsonValue;
use wasm_bindgen::prelude::*;
use wasm_bindgen::JsCast;

use formula_dax::{
    dax_value_to_pivot_value, pivot, pivot_crosstab, Cardinality, CrossFilterDirection, DataModel,
    DaxEngine, DaxError, Expr, FilterContext, GroupByColumn, PivotMeasure, Relationship, RowContext,
    Table, Value,
};
use formula_model::pivots::{PivotKeyPart, PivotValue};

fn js_error(message: impl AsRef<str>) -> JsValue {
    js_sys::Error::new(message.as_ref()).into()
}

fn dax_error_to_js(err: DaxError) -> JsValue {
    js_sys::Error::new(&err.to_string()).into()
}

fn js_value_to_dax_value(value: JsValue) -> Result<Value, JsValue> {
    if value.is_null() || value.is_undefined() {
        return Ok(Value::Blank);
    }

    if let Some(n) = value.as_f64() {
        return Ok(Value::from(n));
    }
    if let Some(s) = value.as_string() {
        return Ok(Value::from(s));
    }
    if let Some(b) = value.as_bool() {
        return Ok(Value::from(b));
    }

    Err(js_error(
        "unsupported value type (expected null, number, string, or boolean)",
    ))
}

fn dax_value_to_js(value: Value) -> JsValue {
    match value {
        Value::Blank => JsValue::NULL,
        Value::Number(n) => JsValue::from_f64(n.0),
        Value::Text(s) => JsValue::from_str(&s),
        Value::Boolean(b) => JsValue::from_bool(b),
    }
}

fn dax_value_to_sort_key(value: &Value) -> PivotKeyPart {
    match value {
        Value::Blank => PivotKeyPart::Blank,
        Value::Number(n) => PivotKeyPart::Number(n.0.to_bits()),
        Value::Text(s) => PivotKeyPart::Text(s.to_string()),
        Value::Boolean(b) => PivotKeyPart::Bool(*b),
    }
}

fn object_set(obj: &Object, key: &str, value: &JsValue) -> Result<(), JsValue> {
    Reflect::set(obj, &JsValue::from_str(key), value).map(|_| ())
}

fn cardinality_from_js(raw: Option<&str>) -> Result<Cardinality, JsValue> {
    match raw.unwrap_or("OneToMany") {
        "OneToMany" => Ok(Cardinality::OneToMany),
        "OneToOne" => Ok(Cardinality::OneToOne),
        "ManyToMany" => Ok(Cardinality::ManyToMany),
        other => Err(js_error(format!(
            "unknown relationship.cardinality: {other} (expected OneToMany | OneToOne | ManyToMany)"
        ))),
    }
}

fn cross_filter_direction_from_js(raw: Option<&str>) -> Result<CrossFilterDirection, JsValue> {
    match raw.unwrap_or("Single") {
        "Single" => Ok(CrossFilterDirection::Single),
        "Both" => Ok(CrossFilterDirection::Both),
        other => Err(js_error(format!(
            "unknown relationship.crossFilterDirection: {other} (expected Single | Both)"
        ))),
    }
}

fn cardinality_to_js(cardinality: Cardinality) -> &'static str {
    match cardinality {
        Cardinality::OneToMany => "OneToMany",
        Cardinality::OneToOne => "OneToOne",
        Cardinality::ManyToMany => "ManyToMany",
    }
}

fn cross_filter_direction_to_js(dir: CrossFilterDirection) -> &'static str {
    match dir {
        CrossFilterDirection::Single => "Single",
        CrossFilterDirection::Both => "Both",
    }
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct RelationshipDto {
    name: String,
    from_table: String,
    from_column: String,
    to_table: String,
    to_column: String,
    #[serde(default)]
    cardinality: Option<String>,
    #[serde(default)]
    cross_filter_direction: Option<String>,
    #[serde(default)]
    is_active: Option<bool>,
    #[serde(default)]
    enforce_referential_integrity: Option<bool>,
}

#[derive(Debug, Deserialize)]
struct GroupByDto {
    table: String,
    column: String,
}

#[derive(Debug, Deserialize)]
struct PivotMeasureDto {
    name: String,
    expression: String,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
struct DaxTableSchemaDto {
    name: String,
    columns: Vec<String>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
struct DaxMeasureSchemaDto {
    name: String,
    expression: String,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
struct DaxRelationshipSchemaDto {
    name: String,
    from_table: String,
    from_column: String,
    to_table: String,
    to_column: String,
    cardinality: String,
    cross_filter_direction: String,
    is_active: bool,
    enforce_referential_integrity: bool,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
struct DaxModelSchemaDto {
    tables: Vec<DaxTableSchemaDto>,
    measures: Vec<DaxMeasureSchemaDto>,
    relationships: Vec<DaxRelationshipSchemaDto>,
}

/// JS-friendly wrapper around [`formula_dax::DataModel`].
#[wasm_bindgen]
pub struct DaxModel {
    model: DataModel,
}

#[wasm_bindgen]
impl DaxModel {
    #[wasm_bindgen(constructor)]
    pub fn new() -> Self {
        Self {
            model: DataModel::new(),
        }
    }

    /// Add an in-memory table from JS data.
    ///
    /// Values are converted as:
    /// - `null`/`undefined` → BLANK
    /// - `number` → Number
    /// - `string` → Text
    /// - `boolean` → Boolean
    #[wasm_bindgen(js_name = "addTable")]
    pub fn add_table(
        &mut self,
        name: &str,
        columns: Vec<String>,
        rows: JsValue,
    ) -> Result<(), JsValue> {
        let rows = rows
            .dyn_into::<Array>()
            .map_err(|_| js_error("rows must be an array of arrays"))?;

        let mut table = Table::new(name, columns);
        for row in rows.iter() {
            let row = row
                .dyn_into::<Array>()
                .map_err(|_| js_error("rows must be an array of arrays"))?;

            let mut out = Vec::with_capacity(row.length() as usize);
            for cell in row.iter() {
                out.push(js_value_to_dax_value(cell)?);
            }
            table.push_row(out).map_err(dax_error_to_js)?;
        }

        self.model.add_table(table).map_err(dax_error_to_js)
    }

    #[wasm_bindgen(js_name = "addRelationship")]
    pub fn add_relationship(&mut self, relationship: JsValue) -> Result<(), JsValue> {
        let dto: RelationshipDto = serde_wasm_bindgen::from_value(relationship)
            .map_err(|err| js_error(err.to_string()))?;

        let relationship = Relationship {
            name: dto.name,
            from_table: dto.from_table,
            from_column: dto.from_column,
            to_table: dto.to_table,
            to_column: dto.to_column,
            cardinality: cardinality_from_js(dto.cardinality.as_deref())?,
            cross_filter_direction: cross_filter_direction_from_js(
                dto.cross_filter_direction.as_deref(),
            )?,
            is_active: dto.is_active.unwrap_or(true),
            enforce_referential_integrity: dto.enforce_referential_integrity.unwrap_or(true),
        };

        self.model
            .add_relationship(relationship)
            .map_err(dax_error_to_js)
    }

    #[wasm_bindgen(js_name = "addMeasure")]
    pub fn add_measure(&mut self, name: &str, expression: &str) -> Result<(), JsValue> {
        self.model
            .add_measure(name, expression)
            .map_err(dax_error_to_js)
    }

    #[wasm_bindgen(js_name = "addCalculatedColumn")]
    pub fn add_calculated_column(
        &mut self,
        table: &str,
        name: &str,
        expression: &str,
    ) -> Result<(), JsValue> {
        self.model
            .add_calculated_column(table, name, expression)
            .map_err(dax_error_to_js)
    }

    /// Returns a lightweight schema for the current Data Model (tables/columns, measures, relationships).
    ///
    /// This is intended for pivot UIs that need to enumerate available fields.
    #[wasm_bindgen(js_name = "getSchema")]
    pub fn get_schema(&self) -> Result<JsValue, JsValue> {
        let mut tables: Vec<DaxTableSchemaDto> = self
            .model
            .tables()
            .map(|t| DaxTableSchemaDto {
                name: t.name().to_string(),
                columns: t.columns().to_vec(),
            })
            .collect();
        tables.sort_by(|a, b| a.name.cmp(&b.name));

        let mut measures: Vec<DaxMeasureSchemaDto> = self
            .model
            .measures_definitions()
            .map(|m| DaxMeasureSchemaDto {
                name: m.name.clone(),
                expression: m.expression.clone(),
            })
            .collect();
        measures.sort_by(|a, b| a.name.cmp(&b.name));

        let mut relationships: Vec<DaxRelationshipSchemaDto> = self
            .model
            .relationships_definitions()
            .map(|r| DaxRelationshipSchemaDto {
                name: r.name.clone(),
                from_table: r.from_table.clone(),
                from_column: r.from_column.clone(),
                to_table: r.to_table.clone(),
                to_column: r.to_column.clone(),
                cardinality: cardinality_to_js(r.cardinality).to_string(),
                cross_filter_direction: cross_filter_direction_to_js(r.cross_filter_direction)
                    .to_string(),
                is_active: r.is_active,
                enforce_referential_integrity: r.enforce_referential_integrity,
            })
            .collect();
        relationships.sort_by(|a, b| a.name.cmp(&b.name));

        serde_wasm_bindgen::to_value(&DaxModelSchemaDto {
            tables,
            measures,
            relationships,
        })
        .map_err(|err| js_error(err.to_string()))
    }

    #[wasm_bindgen(js_name = "evaluate")]
    pub fn evaluate(
        &self,
        expression_or_measure_name: &str,
        filter_context: Option<DaxFilterContext>,
    ) -> Result<JsValue, JsValue> {
        let filter = filter_context
            .map(|ctx| ctx.ctx)
            .unwrap_or_else(FilterContext::empty);

        // Provide a JS-friendly shortcut: if the input matches a known measure name,
        // evaluate it as a measure even if it isn't wrapped in `[brackets]`.
        match self
            .model
            .evaluate_measure(expression_or_measure_name, &filter)
        {
            Ok(value) => Ok(dax_value_to_js(value)),
            Err(DaxError::UnknownMeasure(_)) => DaxEngine::new()
                .evaluate(
                    &self.model,
                    expression_or_measure_name,
                    &filter,
                    &RowContext::default(),
                )
                .map(dax_value_to_js)
                .map_err(dax_error_to_js),
            Err(err) => Err(dax_error_to_js(err)),
        }
    }

    /// Evaluate an expression/measure under a filter context **without consuming** the
    /// `DaxFilterContext` JS object.
    ///
    /// `evaluate()` takes `filterContext?: DaxFilterContext` by value because wasm-bindgen does not
    /// currently support `Option<&T>` for exported classes. Passing a `DaxFilterContext` into
    /// `evaluate()` will therefore consume it.
    ///
    /// JS callers that want to reuse a filter context across multiple evaluations should prefer
    /// this API.
    #[wasm_bindgen(js_name = "evaluateWithFilter")]
    pub fn evaluate_with_filter(
        &self,
        expression_or_measure_name: &str,
        filter_context: &DaxFilterContext,
    ) -> Result<JsValue, JsValue> {
        let filter = filter_context.ctx.clone();
        match self
            .model
            .evaluate_measure(expression_or_measure_name, &filter)
        {
            Ok(value) => Ok(dax_value_to_js(value)),
            Err(DaxError::UnknownMeasure(_)) => DaxEngine::new()
                .evaluate(
                    &self.model,
                    expression_or_measure_name,
                    &filter,
                    &RowContext::default(),
                )
                .map(dax_value_to_js)
                .map_err(dax_error_to_js),
            Err(err) => Err(dax_error_to_js(err)),
        }
    }

    /// Returns the distinct values for `table[column]` under the provided filter context.
    ///
    /// This includes the relationship-generated virtual BLANK member when present.
    #[wasm_bindgen(js_name = "getDistinctColumnValues")]
    pub fn get_distinct_column_values(
        &self,
        table: &str,
        column: &str,
        filter_context: Option<DaxFilterContext>,
    ) -> Result<JsValue, JsValue> {
        let filter = filter_context
            .map(|ctx| ctx.ctx)
            .unwrap_or_else(FilterContext::empty);
        self.get_distinct_column_values_inner(table, column, &filter)
    }

    /// Like [`get_distinct_column_values`](Self::get_distinct_column_values), but borrows the
    /// provided filter context instead of consuming it.
    #[wasm_bindgen(js_name = "getDistinctColumnValuesWithFilter")]
    pub fn get_distinct_column_values_with_filter(
        &self,
        table: &str,
        column: &str,
        filter_context: &DaxFilterContext,
    ) -> Result<JsValue, JsValue> {
        let filter = filter_context.ctx.clone();
        self.get_distinct_column_values_inner(table, column, &filter)
    }

    /// Return a paged slice of distinct values for `table[column]`.
    ///
    /// This is useful for high-cardinality fields where returning every unique value would be too
    /// expensive to transport to JS at once.
    #[wasm_bindgen(js_name = "getDistinctColumnValuesPaged")]
    pub fn get_distinct_column_values_paged(
        &self,
        table: &str,
        column: &str,
        offset: u32,
        limit: u32,
        filter_context: Option<DaxFilterContext>,
    ) -> Result<JsValue, JsValue> {
        let filter = filter_context
            .map(|ctx| ctx.ctx)
            .unwrap_or_else(FilterContext::empty);
        self.get_distinct_column_values_paged_inner(table, column, offset, limit, &filter)
    }

    /// Like [`get_distinct_column_values_paged`](Self::get_distinct_column_values_paged), but borrows
    /// the provided filter context instead of consuming it.
    #[wasm_bindgen(js_name = "getDistinctColumnValuesPagedWithFilter")]
    pub fn get_distinct_column_values_paged_with_filter(
        &self,
        table: &str,
        column: &str,
        offset: u32,
        limit: u32,
        filter_context: &DaxFilterContext,
    ) -> Result<JsValue, JsValue> {
        let filter = filter_context.ctx.clone();
        self.get_distinct_column_values_paged_inner(table, column, offset, limit, &filter)
    }

    fn get_distinct_column_values_inner(
        &self,
        table: &str,
        column: &str,
        filter: &FilterContext,
    ) -> Result<JsValue, JsValue> {
        self.get_distinct_column_values_paged_inner(
            table,
            column,
            0,
            u32::MAX,
            filter,
        )
    }

    fn distinct_column_values_sorted(
        &self,
        table: &str,
        column: &str,
        filter: &FilterContext,
    ) -> Result<Vec<Value>, JsValue> {
        let expr = Expr::ColumnRef {
            table: table.to_string(),
            column: column.to_string(),
        };
        let values = DaxEngine::new()
            .distinct_column_values(&self.model, &expr, filter)
            .map_err(dax_error_to_js)?;

        let mut values: Vec<(PivotKeyPart, Value)> = values
            .into_iter()
            .map(|v| {
                let key = dax_value_to_sort_key(&v);
                (key, v)
            })
            .collect();
        values.sort_by(|(a, _), (b, _)| a.cmp(b));
        Ok(values.into_iter().map(|(_, v)| v).collect())
    }

    fn get_distinct_column_values_paged_inner(
        &self,
        table: &str,
        column: &str,
        offset: u32,
        limit: u32,
        filter: &FilterContext,
    ) -> Result<JsValue, JsValue> {
        let values = self.distinct_column_values_sorted(table, column, filter)?;
        let start = offset as usize;
        let end = start.saturating_add(limit as usize).min(values.len());
        let slice: &[Value] = if start >= values.len() {
            &[]
        } else {
            &values[start..end]
        };
        let out = Array::new();
        for value in slice {
            out.push(&dax_value_to_js(value.clone()));
        }
        Ok(out.into())
    }

    #[wasm_bindgen(js_name = "pivot")]
    pub fn pivot(
        &self,
        base_table: &str,
        group_by: JsValue,
        measures: JsValue,
        filter_context: Option<DaxFilterContext>,
    ) -> Result<JsValue, JsValue> {
        let group_by: Vec<GroupByDto> =
            serde_wasm_bindgen::from_value(group_by).map_err(|err| js_error(err.to_string()))?;
        let measures: Vec<PivotMeasureDto> =
            serde_wasm_bindgen::from_value(measures).map_err(|err| js_error(err.to_string()))?;

        let group_by: Vec<GroupByColumn> = group_by
            .into_iter()
            .map(|c| GroupByColumn::new(c.table, c.column))
            .collect();

        let mut pivot_measures = Vec::with_capacity(measures.len());
        for m in measures {
            pivot_measures.push(PivotMeasure::new(m.name, m.expression).map_err(dax_error_to_js)?);
        }

        let filter = filter_context
            .map(|ctx| ctx.ctx)
            .unwrap_or_else(FilterContext::empty);

        let result = pivot(&self.model, base_table, &group_by, &pivot_measures, &filter)
            .map_err(dax_error_to_js)?;

        let obj = Object::new();
        let cols = Array::new();
        for col in result.columns {
            cols.push(&JsValue::from_str(&col));
        }
        let rows = Array::new();
        for row in result.rows {
            let out = Array::new();
            for value in row {
                out.push(&dax_value_to_js(value));
            }
            rows.push(&out);
        }

        object_set(&obj, "columns", &cols.into())?;
        object_set(&obj, "rows", &rows.into())?;
        Ok(obj.into())
    }

    /// Compute an Excel-like pivot *crosstab* grid (row axis + column axis).
    ///
    /// This is a convenience wrapper around [`formula_dax::pivot_crosstab`]; it returns an object
    /// shaped as `{ data: any[][] }` where the first row is a header row.
    #[wasm_bindgen(js_name = "pivotCrosstab")]
    pub fn pivot_crosstab(
        &self,
        base_table: &str,
        row_fields: JsValue,
        column_fields: JsValue,
        measures: JsValue,
        filter_context: Option<DaxFilterContext>,
    ) -> Result<JsValue, JsValue> {
        let row_fields: Vec<GroupByDto> = serde_wasm_bindgen::from_value(row_fields)
            .map_err(|err| js_error(err.to_string()))?;
        let column_fields: Vec<GroupByDto> = serde_wasm_bindgen::from_value(column_fields)
            .map_err(|err| js_error(err.to_string()))?;
        let measures: Vec<PivotMeasureDto> =
            serde_wasm_bindgen::from_value(measures).map_err(|err| js_error(err.to_string()))?;

        let row_fields: Vec<GroupByColumn> = row_fields
            .into_iter()
            .map(|c| GroupByColumn::new(c.table, c.column))
            .collect();
        let column_fields: Vec<GroupByColumn> = column_fields
            .into_iter()
            .map(|c| GroupByColumn::new(c.table, c.column))
            .collect();

        let mut pivot_measures = Vec::with_capacity(measures.len());
        for m in measures {
            pivot_measures.push(PivotMeasure::new(m.name, m.expression).map_err(dax_error_to_js)?);
        }

        let filter = filter_context
            .map(|ctx| ctx.ctx)
            .unwrap_or_else(FilterContext::empty);

        let result = pivot_crosstab(
            &self.model,
            base_table,
            &row_fields,
            &column_fields,
            &pivot_measures,
            &filter,
        )
        .map_err(dax_error_to_js)?;

        let data = Array::new();
        for row in result.data {
            let out = Array::new();
            for value in row {
                out.push(&dax_value_to_js(value));
            }
            data.push(&out);
        }

        let obj = Object::new();
        object_set(&obj, "data", &data.into())?;
        Ok(obj.into())
    }

    /// Borrowing variant of [`pivot_crosstab`](Self::pivot_crosstab) that does not consume the filter context.
    #[wasm_bindgen(js_name = "pivotCrosstabWithFilter")]
    pub fn pivot_crosstab_with_filter(
        &self,
        base_table: &str,
        row_fields: JsValue,
        column_fields: JsValue,
        measures: JsValue,
        filter_context: &DaxFilterContext,
    ) -> Result<JsValue, JsValue> {
        let filter = filter_context.ctx.clone();
        self.pivot_crosstab(
            base_table,
            row_fields,
            column_fields,
            measures,
            Some(DaxFilterContext { ctx: filter }),
        )
    }

    /// Pivot query variant that borrows the provided `DaxFilterContext` instead of consuming it.
    #[wasm_bindgen(js_name = "pivotWithFilter")]
    pub fn pivot_with_filter(
        &self,
        base_table: &str,
        group_by: JsValue,
        measures: JsValue,
        filter_context: &DaxFilterContext,
    ) -> Result<JsValue, JsValue> {
        let filter = filter_context.ctx.clone();
        self.pivot(
            base_table,
            group_by,
            measures,
            Some(DaxFilterContext { ctx: filter }),
        )
    }

    /// Apply `CALCULATE`-style filter arguments to an existing filter context, returning the
    /// resulting [`DaxFilterContext`].
    ///
    /// This enables expressing filters that can't be represented as simple column equals/in sets
    /// (for example: `Customers[Region] <> BLANK()`).
    #[wasm_bindgen(js_name = "applyCalculateFilters")]
    pub fn apply_calculate_filters(
        &self,
        base: Option<DaxFilterContext>,
        filter_args: Vec<String>,
    ) -> Result<DaxFilterContext, JsValue> {
        let base_filter = base.map(|ctx| ctx.ctx).unwrap_or_else(FilterContext::empty);
        let args: Vec<&str> = filter_args.iter().map(|s| s.as_str()).collect();
        let filter = DaxEngine::new()
            .apply_calculate_filters(&self.model, &base_filter, &args)
            .map_err(dax_error_to_js)?;
        Ok(DaxFilterContext { ctx: filter })
    }
}

/// JS wrapper around [`formula_dax::FilterContext`].
#[wasm_bindgen]
pub struct DaxFilterContext {
    ctx: FilterContext,
}

#[wasm_bindgen]
impl DaxFilterContext {
    #[wasm_bindgen(constructor)]
    pub fn new() -> Self {
        Self {
            ctx: FilterContext::empty(),
        }
    }

    #[wasm_bindgen(js_name = "setColumnEquals")]
    pub fn set_column_equals(
        &mut self,
        table: &str,
        column: &str,
        value: JsValue,
    ) -> Result<(), JsValue> {
        let value = js_value_to_dax_value(value)?;
        self.ctx.set_column_equals(table, column, value);
        Ok(())
    }

    /// Clone this filter context.
    ///
    /// Useful when passing a filter context into APIs that take ownership (e.g. `evaluate()`).
    #[wasm_bindgen(js_name = "clone")]
    pub fn clone_js(&self) -> DaxFilterContext {
        DaxFilterContext {
            ctx: self.ctx.clone(),
        }
    }

    /// Set a multi-value filter for a column (`column IN { ... }`).
    ///
    /// `values` must be an array of scalars (`null`/`undefined`, number, string, boolean).
    #[wasm_bindgen(js_name = "setColumnIn")]
    pub fn set_column_in(
        &mut self,
        table: &str,
        column: &str,
        values: Vec<JsValue>,
    ) -> Result<(), JsValue> {
        let mut out = Vec::with_capacity(values.len());
        for value in values {
            out.push(js_value_to_dax_value(value)?);
        }
        self.ctx.set_column_in(table, column, out);
        Ok(())
    }

    /// Clear any filter for a specific column.
    #[wasm_bindgen(js_name = "clearColumnFilter")]
    pub fn clear_column_filter(&mut self, table: &str, column: &str) {
        self.ctx.clear_column_filter_public(table, column);
    }

    #[wasm_bindgen(js_name = "clearTableFilters")]
    pub fn clear_table_filters(&mut self, table: &str) {
        self.ctx.clear_table_filters_public(table);
    }
}

// -----------------------------------------------------------------------------
// Compatibility: a minimal serde-friendly pivot API added in an earlier iteration.
// -----------------------------------------------------------------------------

#[derive(Clone, Debug, Deserialize)]
pub struct TableSchemaDto {
    pub name: String,
    pub columns: Vec<String>,
}

#[derive(Clone, Debug, Deserialize)]
pub struct MeasureDto {
    pub name: String,
    pub expression: String,
}

#[derive(Clone, Debug, Deserialize)]
pub struct GroupByColumnDto {
    pub table: String,
    pub column: String,
}

#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PivotRequestDto {
    pub base_table: String,
    pub group_by: Vec<GroupByColumnDto>,
    pub measures: Vec<MeasureDto>,
}

#[derive(Clone, Debug, Serialize, PartialEq)]
pub struct PivotResultDto {
    pub columns: Vec<String>,
    pub rows: Vec<Vec<JsonValue>>,
}

/// Pivot output where every cell uses the canonical pivot scalar representation (`{ type, value }`).
#[derive(Clone, Debug, Serialize, PartialEq)]
pub struct PivotResultPivotValuesDto {
    pub columns: Vec<String>,
    pub rows: Vec<Vec<PivotValue>>,
}

#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PivotCrosstabRequestDto {
    pub base_table: String,
    pub row_fields: Vec<GroupByColumnDto>,
    pub column_fields: Vec<GroupByColumnDto>,
    pub measures: Vec<MeasureDto>,
}

#[derive(Clone, Debug, Serialize, PartialEq)]
pub struct PivotGridDto {
    pub data: Vec<Vec<JsonValue>>,
}

/// Crosstab output where every cell uses the canonical pivot scalar representation (`{ type, value }`).
#[derive(Clone, Debug, Serialize, PartialEq)]
pub struct PivotGridPivotValuesDto {
    pub data: Vec<Vec<PivotValue>>,
}

fn json_scalar_to_dax_value(value: &JsonValue) -> Result<Value, JsValue> {
    match value {
        JsonValue::Null => Ok(Value::Blank),
        JsonValue::Bool(b) => Ok(Value::from(*b)),
        JsonValue::Number(n) => n
            .as_f64()
            .ok_or_else(|| super::js_err("invalid number".to_string()))
            .map(Value::from),
        JsonValue::String(s) => Ok(Value::from(s.clone())),
        _ => Err(super::js_err("expected a JSON scalar".to_string())),
    }
}

fn dax_value_to_json_scalar(value: &Value) -> JsonValue {
    match value {
        Value::Blank => JsonValue::Null,
        Value::Boolean(b) => JsonValue::Bool(*b),
        Value::Number(n) => serde_json::Number::from_f64(n.0)
            .map(JsonValue::Number)
            .unwrap_or(JsonValue::Null),
        Value::Text(s) => JsonValue::String(s.to_string()),
    }
}

#[wasm_bindgen(js_name = "DaxDataModel")]
pub struct WasmDaxDataModel {
    inner: DataModel,
}

#[wasm_bindgen(js_class = "DaxDataModel")]
impl WasmDaxDataModel {
    #[wasm_bindgen(constructor)]
    pub fn new() -> WasmDaxDataModel {
        WasmDaxDataModel {
            inner: DataModel::new(),
        }
    }

    #[wasm_bindgen(js_name = "addTable")]
    pub fn add_table(&mut self, schema: JsValue, rows: JsValue) -> Result<(), JsValue> {
        let schema: TableSchemaDto =
            serde_wasm_bindgen::from_value(schema).map_err(|err| super::js_err(err.to_string()))?;
        let rows: Vec<Vec<JsonValue>> =
            serde_wasm_bindgen::from_value(rows).map_err(|err| super::js_err(err.to_string()))?;

        let mut table = Table::new(&schema.name, schema.columns);
        for row in rows {
            let mut values = Vec::with_capacity(row.len());
            for cell in row {
                values.push(json_scalar_to_dax_value(&cell)?);
            }
            table
                .push_row(values)
                .map_err(|err| super::js_err(err.to_string()))?;
        }

        self.inner
            .add_table(table)
            .map_err(|err| super::js_err(err.to_string()))
    }

    #[wasm_bindgen(js_name = "addMeasure")]
    pub fn add_measure(&mut self, measure: JsValue) -> Result<(), JsValue> {
        let measure: MeasureDto = serde_wasm_bindgen::from_value(measure)
            .map_err(|err| super::js_err(err.to_string()))?;
        self.inner
            .add_measure(measure.name, measure.expression)
            .map_err(|err| super::js_err(err.to_string()))
    }

    #[wasm_bindgen(js_name = "insertRow")]
    pub fn insert_row(&mut self, table: String, row: JsValue) -> Result<(), JsValue> {
        let row: Vec<JsonValue> =
            serde_wasm_bindgen::from_value(row).map_err(|err| super::js_err(err.to_string()))?;

        let mut values = Vec::with_capacity(row.len());
        for cell in row {
            values.push(json_scalar_to_dax_value(&cell)?);
        }

        self.inner
            .insert_row(&table, values)
            .map_err(|err| super::js_err(err.to_string()))
    }

    #[wasm_bindgen(js_name = "addRelationship")]
    pub fn add_relationship(&mut self, relationship: JsValue) -> Result<(), JsValue> {
        let dto: RelationshipDto = serde_wasm_bindgen::from_value(relationship)
            .map_err(|err| super::js_err(err.to_string()))?;

        let relationship = Relationship {
            name: dto.name,
            from_table: dto.from_table,
            from_column: dto.from_column,
            to_table: dto.to_table,
            to_column: dto.to_column,
            cardinality: cardinality_from_js(dto.cardinality.as_deref())?,
            cross_filter_direction: cross_filter_direction_from_js(
                dto.cross_filter_direction.as_deref(),
            )?,
            is_active: dto.is_active.unwrap_or(true),
            // Unlike `DaxModel`, default to *not* enforcing referential integrity so fact rows with
            // missing dimension keys still participate via the virtual blank row.
            enforce_referential_integrity: dto.enforce_referential_integrity.unwrap_or(false),
        };

        self.inner
            .add_relationship(relationship)
            .map_err(|err| super::js_err(err.to_string()))
    }

    #[wasm_bindgen(js_name = "pivot")]
    pub fn pivot(&self, request: JsValue) -> Result<JsValue, JsValue> {
        let request: PivotRequestDto = serde_wasm_bindgen::from_value(request)
            .map_err(|err| super::js_err(err.to_string()))?;

        let group_by: Vec<GroupByColumn> = request
            .group_by
            .into_iter()
            .map(|col| GroupByColumn::new(col.table, col.column))
            .collect();

        let measures: Vec<PivotMeasure> = request
            .measures
            .into_iter()
            .map(|m| PivotMeasure::new(m.name, m.expression))
            .collect::<Result<Vec<_>, _>>()
            .map_err(|err| super::js_err(err.to_string()))?;

        let result = formula_dax::pivot(
            &self.inner,
            &request.base_table,
            &group_by,
            &measures,
            &FilterContext::empty(),
        )
        .map_err(|err| super::js_err(err.to_string()))?;

        let rows = result
            .rows
            .into_iter()
            .map(|row| row.iter().map(dax_value_to_json_scalar).collect())
            .collect();

        let out = PivotResultDto {
            columns: result.columns,
            rows,
        };

        use serde::ser::Serialize as _;
        out.serialize(&serde_wasm_bindgen::Serializer::json_compatible())
            .map_err(|err| super::js_err(err.to_string()))
    }

    /// Like [`pivot`](Self::pivot), but returns values using the canonical pivot scalar
    /// representation (tagged `{ type, value }` payloads).
    #[wasm_bindgen(js_name = "pivotPivotValues")]
    pub fn pivot_pivot_values(&self, request: JsValue) -> Result<JsValue, JsValue> {
        let request: PivotRequestDto = serde_wasm_bindgen::from_value(request)
            .map_err(|err| super::js_err(err.to_string()))?;

        let group_by: Vec<GroupByColumn> = request
            .group_by
            .into_iter()
            .map(|col| GroupByColumn::new(col.table, col.column))
            .collect();

        let measures: Vec<PivotMeasure> = request
            .measures
            .into_iter()
            .map(|m| PivotMeasure::new(m.name, m.expression))
            .collect::<Result<Vec<_>, _>>()
            .map_err(|err| super::js_err(err.to_string()))?;

        let result = formula_dax::pivot(
            &self.inner,
            &request.base_table,
            &group_by,
            &measures,
            &FilterContext::empty(),
        )
        .map_err(|err| super::js_err(err.to_string()))?;

        let rows = result
            .rows
            .into_iter()
            .map(|row| row.into_iter().map(|v| dax_value_to_pivot_value(&v)).collect())
            .collect();

        let out = PivotResultPivotValuesDto {
            columns: result.columns,
            rows,
        };

        use serde::ser::Serialize as _;
        out.serialize(&serde_wasm_bindgen::Serializer::json_compatible())
            .map_err(|err| super::js_err(err.to_string()))
    }

    #[wasm_bindgen(js_name = "pivotCrosstab")]
    pub fn pivot_crosstab(&self, request: JsValue) -> Result<JsValue, JsValue> {
        let request: PivotCrosstabRequestDto = serde_wasm_bindgen::from_value(request)
            .map_err(|err| super::js_err(err.to_string()))?;

        let row_fields: Vec<GroupByColumn> = request
            .row_fields
            .into_iter()
            .map(|col| GroupByColumn::new(col.table, col.column))
            .collect();
        let column_fields: Vec<GroupByColumn> = request
            .column_fields
            .into_iter()
            .map(|col| GroupByColumn::new(col.table, col.column))
            .collect();

        let measures: Vec<PivotMeasure> = request
            .measures
            .into_iter()
            .map(|m| PivotMeasure::new(m.name, m.expression))
            .collect::<Result<Vec<_>, _>>()
            .map_err(|err| super::js_err(err.to_string()))?;

        let grid = pivot_crosstab(
            &self.inner,
            &request.base_table,
            &row_fields,
            &column_fields,
            &measures,
            &FilterContext::empty(),
        )
        .map_err(|err| super::js_err(err.to_string()))?;

        let data = grid
            .data
            .into_iter()
            .map(|row| row.iter().map(dax_value_to_json_scalar).collect())
            .collect();

        let out = PivotGridDto { data };

        use serde::ser::Serialize as _;
        out.serialize(&serde_wasm_bindgen::Serializer::json_compatible())
            .map_err(|err| super::js_err(err.to_string()))
    }

    /// Like [`pivot_crosstab`](Self::pivot_crosstab), but returns values using the canonical pivot
    /// scalar representation (tagged `{ type, value }` payloads).
    #[wasm_bindgen(js_name = "pivotCrosstabPivotValues")]
    pub fn pivot_crosstab_pivot_values(&self, request: JsValue) -> Result<JsValue, JsValue> {
        let request: PivotCrosstabRequestDto = serde_wasm_bindgen::from_value(request)
            .map_err(|err| super::js_err(err.to_string()))?;

        let row_fields: Vec<GroupByColumn> = request
            .row_fields
            .into_iter()
            .map(|col| GroupByColumn::new(col.table, col.column))
            .collect();
        let column_fields: Vec<GroupByColumn> = request
            .column_fields
            .into_iter()
            .map(|col| GroupByColumn::new(col.table, col.column))
            .collect();

        let measures: Vec<PivotMeasure> = request
            .measures
            .into_iter()
            .map(|m| PivotMeasure::new(m.name, m.expression))
            .collect::<Result<Vec<_>, _>>()
            .map_err(|err| super::js_err(err.to_string()))?;

        let grid = pivot_crosstab(
            &self.inner,
            &request.base_table,
            &row_fields,
            &column_fields,
            &measures,
            &FilterContext::empty(),
        )
        .map_err(|err| super::js_err(err.to_string()))?;

        let data = grid
            .data
            .into_iter()
            .map(|row| row.into_iter().map(|v| dax_value_to_pivot_value(&v)).collect())
            .collect();

        let out = PivotGridPivotValuesDto { data };

        use serde::ser::Serialize as _;
        out.serialize(&serde_wasm_bindgen::Serializer::json_compatible())
            .map_err(|err| super::js_err(err.to_string()))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn dax_value_json_scalar_roundtrip_smoke() {
        let values = vec![
            Value::Blank,
            Value::from(true),
            Value::from(false),
            Value::from(42.0),
            Value::from("Hello"),
        ];

        for value in values {
            let json = dax_value_to_json_scalar(&value);
            let back = json_scalar_to_dax_value(&json).unwrap();
            assert_eq!(back, value);
        }
    }
}
