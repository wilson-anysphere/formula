use serde::{Deserialize, Serialize};
use serde_json::Value as JsonValue;
use wasm_bindgen::prelude::*;

use formula_dax::{FilterContext, GroupByColumn, PivotMeasure, Value as DaxValue};

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

fn json_scalar_to_dax_value(value: &JsonValue) -> Result<DaxValue, JsValue> {
    match value {
        JsonValue::Null => Ok(DaxValue::Blank),
        JsonValue::Bool(b) => Ok(DaxValue::from(*b)),
        JsonValue::Number(n) => n
            .as_f64()
            .ok_or_else(|| super::js_err("invalid number".to_string()))
            .map(DaxValue::from),
        JsonValue::String(s) => Ok(DaxValue::from(s.clone())),
        _ => Err(super::js_err("expected a JSON scalar".to_string())),
    }
}

fn dax_value_to_json_scalar(value: &DaxValue) -> JsonValue {
    match value {
        DaxValue::Blank => JsonValue::Null,
        DaxValue::Boolean(b) => JsonValue::Bool(*b),
        DaxValue::Number(n) => serde_json::Number::from_f64(n.0)
            .map(JsonValue::Number)
            .unwrap_or(JsonValue::Null),
        DaxValue::Text(s) => JsonValue::String(s.to_string()),
    }
}

#[wasm_bindgen(js_name = "DaxDataModel")]
pub struct WasmDaxDataModel {
    inner: formula_dax::DataModel,
}

#[wasm_bindgen(js_class = "DaxDataModel")]
impl WasmDaxDataModel {
    #[wasm_bindgen(constructor)]
    pub fn new() -> WasmDaxDataModel {
        WasmDaxDataModel {
            inner: formula_dax::DataModel::new(),
        }
    }

    #[wasm_bindgen(js_name = "addTable")]
    pub fn add_table(&mut self, schema: JsValue, rows: JsValue) -> Result<(), JsValue> {
        let schema: TableSchemaDto =
            serde_wasm_bindgen::from_value(schema).map_err(|err| super::js_err(err.to_string()))?;
        let rows: Vec<Vec<JsonValue>> =
            serde_wasm_bindgen::from_value(rows).map_err(|err| super::js_err(err.to_string()))?;

        let mut table = formula_dax::Table::new(&schema.name, schema.columns);
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
        let measure: MeasureDto =
            serde_wasm_bindgen::from_value(measure).map_err(|err| super::js_err(err.to_string()))?;
        self.inner
            .add_measure(measure.name, measure.expression)
            .map_err(|err| super::js_err(err.to_string()))
    }

    #[wasm_bindgen(js_name = "pivot")]
    pub fn pivot(&self, request: JsValue) -> Result<JsValue, JsValue> {
        let request: PivotRequestDto =
            serde_wasm_bindgen::from_value(request).map_err(|err| super::js_err(err.to_string()))?;

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
}

