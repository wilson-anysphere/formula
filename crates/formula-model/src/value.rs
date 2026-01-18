use serde::{Deserialize, Serialize};

use crate::drawings::ImageId;
pub use crate::rich_text::RichText;
use crate::{CellRef, ErrorValue};
use std::collections::BTreeMap;
use std::fmt;

pub(crate) fn text_eq_case_insensitive(a: &str, b: &str) -> bool {
    if a.is_ascii() && b.is_ascii() {
        return a.eq_ignore_ascii_case(b);
    }
    a.chars()
        .flat_map(|c| c.to_uppercase())
        .eq(b.chars().flat_map(|c| c.to_uppercase()))
}

fn map_get_case_insensitive<'a, V>(map: &'a BTreeMap<String, V>, key: &str) -> Option<&'a V> {
    if let Some(value) = map.get(key) {
        return Some(value);
    }
    map.iter()
        .find(|(k, _)| text_eq_case_insensitive(k, key))
        .map(|(_, v)| v)
}

/// Versioned, JSON-friendly representation of a cell value.
///
/// The enum uses an explicit `{type, value}` tagged layout for stable IPC.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(tag = "type", content = "value", rename_all = "snake_case")]
pub enum CellValue {
    /// Empty / unset cell value.
    Empty,
    /// IEEE-754 double precision number.
    Number(f64),
    /// Plain string (not rich text).
    String(String),
    /// Boolean.
    Boolean(bool),
    /// Excel error value.
    Error(ErrorValue),
    /// Rich text value (stub).
    RichText(RichText),
    /// Excel rich value (data type) representing an entity.
    ///
    /// This is a JSON-friendly representation of Excel's "entity" rich values
    /// (e.g. Stocks / Geography) where the primary scalar representation is a
    /// display string.
    Entity(LinkedEntityValue),
    /// Excel rich value representing a record.
    ///
    /// Record values are treated similarly to entities for scalar export: they
    /// degrade to their display string in legacy IO paths.
    Record(RecordValue),
    /// In-cell image value (Excel "image in cell" / `IMAGE()`).
    ///
    /// This is a first-class value variant so that workbook JSON/IPC payloads can reference
    /// media stored in [`Workbook::images`](crate::Workbook::images) without requiring a full
    /// XLSX round-trip implementation.
    Image(ImageValue),
    /// Array result (stub).
    Array(ArrayValue),
    /// Marker for a cell that is part of a spilled array (stub).
    Spill(SpillValue),
}

impl Default for CellValue {
    fn default() -> Self {
        CellValue::Empty
    }
}

impl CellValue {
    /// Returns true if the value is [`CellValue::Empty`].
    pub fn is_empty(&self) -> bool {
        matches!(self, CellValue::Empty)
    }
}

impl From<f64> for CellValue {
    fn from(value: f64) -> Self {
        CellValue::Number(value)
    }
}

impl From<bool> for CellValue {
    fn from(value: bool) -> Self {
        CellValue::Boolean(value)
    }
}

impl From<String> for CellValue {
    fn from(value: String) -> Self {
        CellValue::String(value)
    }
}

impl From<&str> for CellValue {
    fn from(value: &str) -> Self {
        CellValue::String(value.to_string())
    }
}

impl From<ErrorValue> for CellValue {
    fn from(value: ErrorValue) -> Self {
        CellValue::Error(value)
    }
}

impl From<RichText> for CellValue {
    fn from(value: RichText) -> Self {
        CellValue::RichText(value)
    }
}

/// JSON-friendly representation of an Excel rich "entity" value.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct LinkedEntityValue {
    /// Entity type discriminator (e.g. `"stock"`, `"geography"`).
    #[serde(
        default,
        alias = "entity_type",
        skip_serializing_if = "String::is_empty"
    )]
    pub entity_type: String,
    /// Entity identifier (e.g. `"AAPL"`).
    #[serde(default, alias = "entity_id", skip_serializing_if = "String::is_empty")]
    pub entity_id: String,
    /// User-visible string representation (what Excel renders in the grid).
    ///
    /// Accept the legacy `"display"` key as an alias for backward compatibility.
    #[serde(default, alias = "display", alias = "display_value")]
    pub display_value: String,
    // Back-compat: earlier prototypes used `"fields"` to match the engine's internal naming.
    #[serde(default, alias = "fields", skip_serializing_if = "BTreeMap::is_empty")]
    pub properties: BTreeMap<String, CellValue>,
}

/// Backwards-compatible name for [`LinkedEntityValue`].
pub type EntityValue = LinkedEntityValue;

impl LinkedEntityValue {
    pub fn new(display: impl Into<String>) -> Self {
        Self {
            display_value: display.into(),
            ..Self::default()
        }
    }

    #[must_use]
    pub fn with_entity_type(mut self, entity_type: impl Into<String>) -> Self {
        self.entity_type = entity_type.into();
        self
    }

    #[must_use]
    pub fn with_entity_id(mut self, entity_id: impl Into<String>) -> Self {
        self.entity_id = entity_id.into();
        self
    }

    #[must_use]
    pub fn with_properties(mut self, properties: BTreeMap<String, CellValue>) -> Self {
        self.properties = properties;
        self
    }

    #[must_use]
    pub fn with_property(mut self, name: impl Into<String>, value: impl Into<CellValue>) -> Self {
        self.properties.insert(name.into(), value.into());
        self
    }
}

impl fmt::Display for LinkedEntityValue {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(&self.display_value)
    }
}

impl From<LinkedEntityValue> for CellValue {
    fn from(value: LinkedEntityValue) -> Self {
        CellValue::Entity(value)
    }
}

/// JSON-friendly representation of an Excel rich "record" value.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct RecordValue {
    #[serde(default, skip_serializing_if = "BTreeMap::is_empty")]
    pub fields: BTreeMap<String, CellValue>,
    #[serde(
        default,
        alias = "display_field",
        skip_serializing_if = "Option::is_none"
    )]
    pub display_field: Option<String>,
    /// Optional precomputed display string (legacy / fallback).
    ///
    /// This field exists to keep older IPC payloads working; when `displayField`
    /// is present and points to a scalar field, UIs should prefer that value.
    #[serde(
        default,
        alias = "display",
        alias = "display_value",
        skip_serializing_if = "String::is_empty"
    )]
    pub display_value: String,
}

impl RecordValue {
    pub fn new(display: impl Into<String>) -> Self {
        Self {
            display_value: display.into(),
            ..Self::default()
        }
    }

    #[must_use]
    pub fn with_fields(mut self, fields: BTreeMap<String, CellValue>) -> Self {
        self.fields = fields;
        self
    }

    #[must_use]
    pub fn with_field(mut self, name: impl Into<String>, value: impl Into<CellValue>) -> Self {
        self.fields.insert(name.into(), value.into());
        self
    }

    #[must_use]
    pub fn with_display_field(mut self, display_field: impl Into<String>) -> Self {
        self.display_field = Some(display_field.into());
        self
    }

    pub(crate) fn get_field_case_insensitive(&self, field: &str) -> Option<&CellValue> {
        map_get_case_insensitive(&self.fields, field)
    }

    fn display_text(&self) -> Option<String> {
        let field = self.display_field.as_deref()?;
        let value = self.get_field_case_insensitive(field)?;
        match value {
            CellValue::Empty => Some(String::new()),
            CellValue::String(s) => Some(s.clone()),
            CellValue::Number(n) => Some(n.to_string()),
            CellValue::Boolean(b) => Some(if *b { "TRUE" } else { "FALSE" }.to_string()),
            CellValue::Error(e) => Some(e.as_str().to_string()),
            CellValue::RichText(rt) => Some(rt.text.clone()),
            CellValue::Entity(entity) => Some(entity.display_value.clone()),
            CellValue::Record(record) => Some(record.to_string()),
            CellValue::Image(image) => Some(
                image
                    .alt_text
                    .as_deref()
                    .filter(|s| !s.is_empty())
                    .unwrap_or("[Image]")
                    .to_string(),
            ),
            _ => None,
        }
    }
}

impl fmt::Display for RecordValue {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        if let Some(display) = self.display_text().as_deref() {
            f.write_str(display)
        } else {
            f.write_str(&self.display_value)
        }
    }
}

impl From<RecordValue> for CellValue {
    fn from(value: RecordValue) -> Self {
        CellValue::Record(value)
    }
}

/// JSON-friendly representation of an in-cell image value.
///
/// This references an entry in [`Workbook::images`](crate::Workbook::images) by id, with optional
/// metadata that can be used for display/round-tripping.
///
/// ## JSON schema
///
/// Stored inside the [`CellValue`] `{type, value}` envelope:
///
/// ```json
/// {
///   "type": "image",
///   "value": {
///     "imageId": "image1.png",
///     "altText": "Logo",
///     "width": 128,
///     "height": 64
///   }
/// }
/// ```
///
/// `width` / `height` are expressed in CSS pixels (px).
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ImageValue {
    /// Image identifier in [`Workbook::images`](crate::Workbook::images).
    #[serde(alias = "image_id")]
    pub image_id: ImageId,
    /// Optional alt text for accessibility / scalar degradation.
    #[serde(default, alias = "alt_text", skip_serializing_if = "Option::is_none")]
    pub alt_text: Option<String>,
    /// Optional display width in pixels.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub width: Option<u32>,
    /// Optional display height in pixels.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub height: Option<u32>,
}

impl From<ImageValue> for CellValue {
    fn from(value: ImageValue) -> Self {
        CellValue::Image(value)
    }
}

/// Stub representation of a dynamic array result.
///
/// For now this stores a 2D matrix. The calculation engine may later choose a
/// more compact representation.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct ArrayValue {
    /// 2D array in row-major order.
    pub data: Vec<Vec<CellValue>>,
}

/// Stub marker for cells that belong to a spilled range.
#[derive(Copy, Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub struct SpillValue {
    /// Origin cell containing the spilling formula.
    pub origin: CellRef,
}

#[cfg(test)]
mod tests {
    use super::{CellValue, EntityValue, RecordValue};
    use serde_json::json;

    #[test]
    fn entity_value_deserializes_legacy_display_aliases() {
        let entity: EntityValue = serde_json::from_value(json!({
            "display": "Entity display"
        }))
        .expect("deserialize legacy entity");
        assert_eq!(entity.display_value, "Entity display");
        let serialized = serde_json::to_value(&entity).expect("serialize entity");
        assert_eq!(
            serialized.get("displayValue").and_then(|v| v.as_str()),
            Some("Entity display")
        );

        let entity: EntityValue = serde_json::from_value(json!({
            "display_value": "Entity display"
        }))
        .expect("deserialize snake_case entity display");
        assert_eq!(entity.display_value, "Entity display");
    }

    #[test]
    fn entity_value_deserializes_snake_case_entity_metadata() {
        let entity: EntityValue = serde_json::from_value(json!({
            "entity_type": "stock",
            "entity_id": "AAPL",
            "displayValue": "Apple"
        }))
        .expect("deserialize snake_case entity metadata");
        assert_eq!(entity.entity_type, "stock");
        assert_eq!(entity.entity_id, "AAPL");
        assert_eq!(entity.display_value, "Apple");
    }

    #[test]
    fn entity_value_deserializes_fields_alias_for_properties() {
        let entity: EntityValue = serde_json::from_value(json!({
            "display": "Entity display",
            "fields": {
                "Price": { "type": "number", "value": 178.5 }
            }
        }))
        .expect("deserialize legacy entity fields");
        assert_eq!(entity.display_value, "Entity display");
        assert_eq!(
            entity.properties.get("Price"),
            Some(&CellValue::Number(178.5))
        );
    }

    #[test]
    fn record_value_deserializes_legacy_display_aliases() {
        let record: RecordValue = serde_json::from_value(json!({
            "display": "Record display"
        }))
        .expect("deserialize legacy record");
        assert_eq!(record.display_value, "Record display");

        let record: RecordValue = serde_json::from_value(json!({
            "display_value": "Record display"
        }))
        .expect("deserialize snake_case record display");
        assert_eq!(record.display_value, "Record display");
    }

    #[test]
    fn record_value_deserializes_display_field_alias() {
        let record: RecordValue = serde_json::from_value(json!({
            "display_field": "Name",
            "fields": {
                "Name": { "type": "string", "value": "Alice" }
            }
        }))
        .expect("deserialize record");
        assert_eq!(record.display_field.as_deref(), Some("Name"));
        assert_eq!(record.to_string(), "Alice");
    }

    #[test]
    fn record_display_field_matches_case_insensitively() {
        let record: RecordValue = serde_json::from_value(json!({
            "displayField": "name",
            "fields": {
                "Name": { "type": "string", "value": "Alice" }
            }
        }))
        .expect("deserialize record");
        assert_eq!(record.display_field.as_deref(), Some("name"));
        assert_eq!(record.to_string(), "Alice");
    }
}
