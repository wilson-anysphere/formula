use serde::{Deserialize, Serialize};

/// An Excel-style PivotTable *calculated field*.
///
/// In Excel, a calculated field is a named formula that behaves like an extra source column:
/// it is evaluated for each record in the pivot cache and can then be used as a field in the
/// pivot configuration (most commonly in the "Values" area).
///
/// The Formula pivot engine persists the raw formula text and treats the calculated field as
/// an additional cache field named [`CalculatedField::name`] when building or refreshing the
/// pivot cache.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CalculatedField {
    pub name: String,
    pub formula: String,
}

/// An Excel-style PivotTable *calculated item*.
///
/// In Excel, a calculated item creates a synthetic member (an "item") inside a specific pivot
/// field. The item is defined by a name and a formula that typically references other items
/// within the same field (for example, creating `Q1` from `Jan + Feb + Mar` inside a `Month`
/// field).
///
/// The Formula pivot engine interprets a calculated item as a post-aggregation transform:
/// after the pivot is grouped by [`CalculatedItem::field`], the engine evaluates
/// [`CalculatedItem::formula`] against the existing items for that field and inserts a new item
/// named [`CalculatedItem::name`] into the pivot results.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CalculatedItem {
    pub field: String,
    pub name: String,
    pub formula: String,
}
