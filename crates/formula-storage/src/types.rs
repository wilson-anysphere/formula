use serde::{Deserialize, Serialize};
use serde_json::Value;
use uuid::Uuid;

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "camelCase")]
pub enum SheetVisibility {
    Visible,
    Hidden,
    VeryHidden,
}

impl SheetVisibility {
    pub fn as_str(&self) -> &'static str {
        match self {
            SheetVisibility::Visible => "visible",
            SheetVisibility::Hidden => "hidden",
            SheetVisibility::VeryHidden => "veryHidden",
        }
    }

    pub fn parse(value: &str) -> Self {
        match value {
            "hidden" => SheetVisibility::Hidden,
            "veryHidden" => SheetVisibility::VeryHidden,
            _ => SheetVisibility::Visible,
        }
    }
}

impl Default for SheetVisibility {
    fn default() -> Self {
        SheetVisibility::Visible
    }
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
pub struct WorkbookMeta {
    pub id: Uuid,
    pub name: String,
    pub metadata: Option<Value>,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
pub struct SheetMeta {
    pub id: Uuid,
    pub workbook_id: Uuid,
    pub name: String,
    pub position: i64,
    #[serde(default)]
    pub visibility: SheetVisibility,
    pub tab_color: Option<String>,
    pub xlsx_sheet_id: Option<i64>,
    pub xlsx_rel_id: Option<String>,
    pub frozen_rows: i64,
    pub frozen_cols: i64,
    pub zoom: f64,
    pub metadata: Option<Value>,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
#[serde(tag = "type", content = "value")]
pub enum CellValue {
    /// An empty cell.
    Empty,
    Number(f64),
    Text(String),
    Boolean(bool),
    Error(String),
}

impl CellValue {
    pub fn is_empty(&self) -> bool {
        matches!(self, CellValue::Empty)
    }
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
pub struct Style {
    pub font_id: Option<i64>,
    pub fill_id: Option<i64>,
    pub border_id: Option<i64>,
    pub number_format: Option<String>,
    pub alignment: Option<Value>,
    pub protection: Option<Value>,
}

impl Style {
    pub(crate) fn canonical_alignment(&self) -> Option<String> {
        self.alignment.as_ref().map(canonical_json)
    }

    pub(crate) fn canonical_protection(&self) -> Option<String> {
        self.protection.as_ref().map(canonical_json)
    }
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
pub struct CellData {
    pub value: CellValue,
    pub formula: Option<String>,
    pub style: Option<Style>,
}

impl CellData {
    pub fn empty() -> Self {
        Self {
            value: CellValue::Empty,
            formula: None,
            style: None,
        }
    }

    pub fn is_truly_empty(&self) -> bool {
        self.value.is_empty() && self.formula.is_none() && self.style.is_none()
    }
}

/// A snapshot of a cell as persisted in SQLite.
///
/// This is used when reading from the `cells` table and when recording entries
/// in `change_log`.
#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
pub struct CellSnapshot {
    pub value: CellValue,
    pub formula: Option<String>,
    pub style_id: Option<i64>,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
pub struct NamedRange {
    pub workbook_id: Uuid,
    pub name: String,
    /// Either `"workbook"` or a sheet name.
    pub scope: String,
    pub reference: String,
}

fn canonical_json(value: &Value) -> String {
    // `serde_json::Value` has no canonical serialization by default because
    // object key order is preserved. For style deduplication we want stable
    // strings, so we sort object keys recursively.
    let mut value = value.clone();
    sort_json_keys(&mut value);
    // Sorting keys makes `to_string` deterministic.
    serde_json::to_string(&value).unwrap_or_else(|_| "null".to_string())
}

fn sort_json_keys(value: &mut Value) {
    match value {
        Value::Object(map) => {
            let mut keys: Vec<String> = map.keys().cloned().collect();
            keys.sort();
            let mut new_map = serde_json::Map::new();
            for k in keys {
                if let Some(mut v) = map.remove(&k) {
                    sort_json_keys(&mut v);
                    new_map.insert(k, v);
                }
            }
            *map = new_map;
        }
        Value::Array(values) => {
            for v in values.iter_mut() {
                sort_json_keys(v);
            }
        }
        _ => {}
    }
}
