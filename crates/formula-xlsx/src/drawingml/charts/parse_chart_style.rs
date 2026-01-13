use formula_model::charts::ChartStylePartModel;
use roxmltree::Document;

#[derive(Debug, thiserror::Error)]
pub enum ChartStyleParseError {
    #[error("part is not valid UTF-8: {part_name}: {source}")]
    XmlNonUtf8 {
        part_name: String,
        #[source]
        source: std::str::Utf8Error,
    },
    #[error("failed to parse XML: {part_name}: {source}")]
    XmlParse {
        part_name: String,
        #[source]
        source: roxmltree::Error,
    },
}

/// Parse a `chartStyle` part (`xl/charts/style*.xml`).
///
/// This is best-effort: we preserve the raw XML and extract the root `@id`
/// attribute when present.
pub fn parse_chart_style(
    style_xml: &[u8],
    part_name: &str,
) -> Result<ChartStylePartModel, ChartStyleParseError> {
    let xml = std::str::from_utf8(style_xml).map_err(|e| ChartStyleParseError::XmlNonUtf8 {
        part_name: part_name.to_string(),
        source: e,
    })?;

    let doc = Document::parse(xml).map_err(|e| ChartStyleParseError::XmlParse {
        part_name: part_name.to_string(),
        source: e,
    })?;

    let root = doc.root_element();
    let id = root
        .attribute("id")
        .or_else(|| root.attribute("Id"))
        .and_then(|v| v.parse::<u32>().ok());

    Ok(ChartStylePartModel {
        id,
        raw_xml: xml.to_string(),
    })
}
