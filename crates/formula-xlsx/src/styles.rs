use formula_model::{parse_argb_hex_color, CfStyleOverride};
use roxmltree::Document;

#[derive(Debug, thiserror::Error)]
pub enum StylesError {
    #[error("xml parse error: {0}")]
    Xml(#[from] roxmltree::Error),
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
        if font
            .children()
            .any(|n| n.is_element() && n.tag_name().name() == "b")
        {
            out.bold = Some(true);
        }
        if font
            .children()
            .any(|n| n.is_element() && n.tag_name().name() == "i")
        {
            out.italic = Some(true);
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

pub struct DxfProvider<'a> {
    pub styles: &'a Styles,
}

impl<'a> formula_model::DifferentialFormatProvider for DxfProvider<'a> {
    fn get_dxf(&self, dxf_id: u32) -> Option<CfStyleOverride> {
        self.styles.dxfs.get(dxf_id as usize).cloned()
    }
}
