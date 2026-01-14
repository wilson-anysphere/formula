use formula_model::{parse_argb_hex_color, Alignment, CfStyleOverride, HorizontalAlignment, VerticalAlignment};
use quick_xml::events::Event;
use quick_xml::Reader;
use roxmltree::Document;

mod cell_styles;
pub use cell_styles::{StylesPart, StylesPartError};
mod style_editor;
pub use style_editor::XlsxStylesEditor;

#[derive(Debug, thiserror::Error)]
pub enum StylesError {
    #[error("xml parse error: {0}")]
    Xml(#[from] roxmltree::Error),
    #[error("quick-xml parse error: {0}")]
    QuickXml(#[from] quick_xml::Error),
    #[error("quick-xml attribute error: {0}")]
    Attr(#[from] quick_xml::events::attributes::AttrError),
    #[error("utf-8 error: {0}")]
    Utf8(#[from] std::str::Utf8Error),
}

#[derive(Clone, Debug, Default)]
pub struct Styles {
    pub dxfs: Vec<CfStyleOverride>,
}

impl Styles {
    pub fn parse(xml: &str) -> Result<Self, StylesError> {
        let doc = Document::parse(xml)?;
        let mut styles = Styles::default();
        let main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        if let Some(dxfs) = doc
            .descendants()
            .find(|n| n.tag_name().name() == "dxfs" && n.tag_name().namespace() == Some(main_ns))
        {
            for dxf in dxfs
                .children()
                .filter(|n| n.is_element() && n.tag_name().name() == "dxf")
            {
                styles.dxfs.push(parse_dxf(dxf, main_ns));
            }
        }
        Ok(styles)
    }
}

fn parse_dxf(dxf: roxmltree::Node<'_, '_>, main_ns: &str) -> CfStyleOverride {
    let mut out = CfStyleOverride::default();
    if let Some(font) = dxf
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "font" && n.tag_name().namespace() == Some(main_ns))
    {
        if let Some(b) = font
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "b")
        {
            out.bold = Some(parse_dxf_bool(b));
        }
        if let Some(i) = font
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "i")
        {
            out.italic = Some(parse_dxf_bool(i));
        }
        if let Some(color) = font.children().find(|n| n.is_element() && n.tag_name().name() == "color") {
            if let Some(rgb) = color.attribute("rgb") {
                out.font_color = parse_argb_hex_color(rgb);
            }
        }
    }

    if let Some(fill) = dxf
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "fill" && n.tag_name().namespace() == Some(main_ns))
    {
        if let Some(pattern_fill) = fill
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "patternFill")
        {
            if let Some(fg) = pattern_fill
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "fgColor")
            {
                if let Some(rgb) = fg.attribute("rgb") {
                    out.fill = parse_argb_hex_color(rgb);
                }
            }
        }
    }

    out
}

fn parse_dxf_bool(el: roxmltree::Node<'_, '_>) -> bool {
    match el.attribute("val") {
        None => true,
        Some(v) => !(v == "0" || v.eq_ignore_ascii_case("false")),
    }
}

pub struct DxfProvider<'a> {
    pub styles: &'a Styles,
}

impl<'a> formula_model::DifferentialFormatProvider for DxfProvider<'a> {
    fn get_dxf(&self, dxf_id: u32) -> Option<CfStyleOverride> {
        self.styles.dxfs.get(dxf_id as usize).cloned()
    }
}

/// Parse `xl/styles.xml` cellXfs alignments.
///
/// Returns a vector indexed by `xf` index. Missing alignments default to `Alignment::default()`.
pub fn parse_cell_xfs_alignments(styles_xml: &str) -> Result<Vec<Alignment>, StylesError> {
    let mut reader = Reader::from_str(styles_xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut in_cell_xfs = false;
    let mut current_xf: Option<Alignment> = None;
    let mut xfs = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if e.local_name().as_ref() == b"cellXfs" => {
                in_cell_xfs = true;
            }
            Event::End(e) if e.local_name().as_ref() == b"cellXfs" => {
                break;
            }
            Event::Start(e) if in_cell_xfs && e.local_name().as_ref() == b"xf" => {
                current_xf = Some(Alignment::default());
            }
            Event::Empty(e) if in_cell_xfs && e.local_name().as_ref() == b"xf" => {
                xfs.push(Alignment::default());
            }
            Event::End(e) if in_cell_xfs && e.local_name().as_ref() == b"xf" => {
                xfs.push(current_xf.take().unwrap_or_default());
            }
            Event::Empty(e) if in_cell_xfs && e.local_name().as_ref() == b"alignment" => {
                if let Some(xf) = current_xf.as_mut() {
                    for attr in e.attributes() {
                        let attr = attr?;
                        let value = std::str::from_utf8(&attr.value)?;
                        match attr.key.as_ref() {
                            b"horizontal" => xf.horizontal = Some(parse_horizontal(value)),
                            b"vertical" => xf.vertical = Some(parse_vertical(value)),
                            b"wrapText" => {
                                xf.wrap_text = value == "1" || value.eq_ignore_ascii_case("true")
                            }
                            b"textRotation" => {
                                let rotation = value.parse::<i16>().unwrap_or_default();
                                xf.rotation = if rotation == 0 { None } else { Some(rotation) };
                            }
                            b"indent" => {
                                xf.indent = value.parse::<u16>().ok();
                            }
                            _ => {}
                        }
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(xfs)
}

fn parse_horizontal(value: &str) -> HorizontalAlignment {
    match value {
        "left" => HorizontalAlignment::Left,
        "center" => HorizontalAlignment::Center,
        "right" => HorizontalAlignment::Right,
        "fill" => HorizontalAlignment::Fill,
        "justify" => HorizontalAlignment::Justify,
        _ => HorizontalAlignment::General,
    }
}

fn parse_vertical(value: &str) -> VerticalAlignment {
    match value {
        "top" => VerticalAlignment::Top,
        "center" => VerticalAlignment::Center,
        "bottom" => VerticalAlignment::Bottom,
        _ => VerticalAlignment::Bottom,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn styles_parse_dxf_respects_bold_italic_val_zero_as_false() {
        let styles_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dxfs count="1">
    <dxf>
      <font>
        <b val="0"/>
        <i val="false"/>
      </font>
    </dxf>
  </dxfs>
</styleSheet>
"#;

        let styles = Styles::parse(styles_xml).unwrap();
        assert_eq!(styles.dxfs.len(), 1);
        assert_eq!(styles.dxfs[0].bold, Some(false));
        assert_eq!(styles.dxfs[0].italic, Some(false));
    }
}
