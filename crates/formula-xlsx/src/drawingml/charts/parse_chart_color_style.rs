use formula_model::charts::ChartColorStylePartModel;
use roxmltree::Document;

#[derive(Debug, thiserror::Error)]
pub enum ChartColorStyleParseError {
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

/// Parse a `chartColorStyle` part (`xl/charts/colors*.xml`).
///
/// This is best-effort: we preserve the raw XML, extract the root `@id`
/// attribute when present, and collect color entries from `a:srgbClr/@val`
/// or `a:schemeClr/@val` elements.
pub fn parse_chart_color_style(
    colors_xml: &[u8],
    part_name: &str,
) -> Result<ChartColorStylePartModel, ChartColorStyleParseError> {
    let xml =
        std::str::from_utf8(colors_xml).map_err(|e| ChartColorStyleParseError::XmlNonUtf8 {
            part_name: part_name.to_string(),
            source: e,
        })?;

    let doc = Document::parse(xml).map_err(|e| ChartColorStyleParseError::XmlParse {
        part_name: part_name.to_string(),
        source: e,
    })?;

    let root = doc.root_element();
    let id = root
        .attribute("id")
        .or_else(|| root.attribute("Id"))
        .and_then(|v| v.parse::<u32>().ok());

    let mut colors = Vec::new();
    for node in doc.descendants().filter(|n| n.is_element()) {
        match node.tag_name().name() {
            "srgbClr" => {
                if let Some(val) = node.attribute("val") {
                    colors.push(val.to_string());
                }
            }
            "schemeClr" => {
                if let Some(val) = node.attribute("val") {
                    colors.push(format!("scheme:{val}"));
                }
            }
            _ => {}
        }
    }

    Ok(ChartColorStylePartModel {
        id,
        colors,
        raw_xml: xml.to_string(),
    })
}
