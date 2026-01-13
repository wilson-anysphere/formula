use serde::{Deserialize, Serialize};

/// Chart style stored in a separate OPC part (e.g. `xl/charts/style1.xml`).
///
/// Excel can reference this from `xl/charts/_rels/chartN.xml.rels` via a
/// `chartStyle` relationship. For now we preserve the raw XML for round-trip
/// and extract a small set of commonly useful fields.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct ChartStylePartModel {
    /// Optional numeric style id (commonly stored as `@id` on the root element).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub id: Option<u32>,
    /// Raw XML payload for debugging / lossless round-trip.
    pub raw_xml: String,
}

/// Chart color style stored in a separate OPC part (e.g. `xl/charts/colors1.xml`).
///
/// Excel can reference this from `xl/charts/_rels/chartN.xml.rels` via a
/// `chartColorStyle` relationship.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct ChartColorStylePartModel {
    /// Optional numeric color style id (commonly stored as `@id` on the root element).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub id: Option<u32>,
    /// Best-effort list of palette entries encountered in the part.
    ///
    /// Values are stored as raw OOXML strings such as `"FF00AA"` for `a:srgbClr`
    /// or `"scheme:accent1"` for `a:schemeClr`.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub colors: Vec<String>,
    /// Raw XML payload for debugging / lossless round-trip.
    pub raw_xml: String,
}
