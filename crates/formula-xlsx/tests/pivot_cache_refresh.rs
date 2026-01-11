use std::io::Cursor;

use formula_xlsx::XlsxPackage;
use pretty_assertions::assert_eq;
use quick_xml::events::Event;
use quick_xml::Reader;

#[derive(Debug, Clone, PartialEq, Eq)]
enum CacheValue {
    String(String),
    Number(String),
    Bool(bool),
    Missing,
}

fn parse_cache_records(xml: &str) -> Vec<Vec<CacheValue>> {
    let mut reader = Reader::from_reader(Cursor::new(xml.as_bytes()));
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut rows: Vec<Vec<CacheValue>> = Vec::new();
    let mut current: Option<Vec<CacheValue>> = None;

    loop {
        match reader.read_event_into(&mut buf).expect("parse xml") {
            Event::Start(e) if e.local_name().as_ref() == b"r" => current = Some(Vec::new()),
            Event::End(e) if e.local_name().as_ref() == b"r" => {
                rows.push(current.take().expect("row started"));
            }
            Event::Empty(e) if current.is_some() => {
                let tag = e.local_name();
                let mut v: Option<String> = None;
                for attr in e.attributes().with_checks(false) {
                    let attr = attr.expect("attr");
                    if attr.key.local_name().as_ref() == b"v" {
                        v = Some(attr.unescape_value().expect("value").into_owned());
                    }
                }

                let row = current.as_mut().expect("row started");
                match tag.as_ref() {
                    b"s" => row.push(CacheValue::String(v.unwrap_or_default())),
                    b"n" => row.push(CacheValue::Number(v.unwrap_or_default())),
                    b"b" => row.push(CacheValue::Bool(v.as_deref() == Some("1"))),
                    b"m" => row.push(CacheValue::Missing),
                    _ => {}
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    rows
}

#[test]
fn refreshes_pivot_cache_records_from_worksheet_range() {
    let fixture = include_bytes!("fixtures/pivot-cache-refresh.xlsx");
    let mut pkg = XlsxPackage::from_bytes(fixture).expect("read fixture");

    let updated_sheet = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Region</t></is></c>
      <c r="B1" t="inlineStr"><is><t>Product</t></is></c>
      <c r="C1" t="inlineStr"><is><t>Revenue</t></is></c>
    </row>
    <row r="2">
      <c r="A2" t="inlineStr"><is><t>East</t></is></c>
      <c r="B2" t="inlineStr"><is><t>A</t></is></c>
      <c r="C2"><v>100</v></c>
    </row>
    <row r="3">
      <c r="A3" t="inlineStr"><is><t>East</t></is></c>
      <c r="B3" t="inlineStr"><is><t>B</t></is></c>
      <c r="C3"><v>150</v></c>
    </row>
    <row r="4">
      <c r="A4" t="inlineStr"><is><t>West</t></is></c>
      <c r="B4" t="inlineStr"><is><t>A</t></is></c>
      <c r="C4"><v>200</v></c>
    </row>
    <row r="5">
      <c r="A5" t="inlineStr"><is><t>West</t></is></c>
      <c r="B5" t="inlineStr"><is><t>B</t></is></c>
      <c r="C5"><v>250</v></c>
    </row>
    <row r="6">
      <c r="A6" t="inlineStr"><is><t>North</t></is></c>
      <c r="B6" t="inlineStr"><is><t>C</t></is></c>
      <c r="C6"><v>300</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    pkg.set_part("xl/worksheets/sheet1.xml", updated_sheet.as_bytes().to_vec());

    pkg.refresh_pivot_cache_from_worksheet("xl/pivotCache/pivotCacheDefinition1.xml")
        .expect("refresh cache");

    let cache_definition_xml = std::str::from_utf8(
        pkg.part("xl/pivotCache/pivotCacheDefinition1.xml")
            .expect("cache definition exists"),
    )
    .expect("utf-8");
    let doc = roxmltree::Document::parse(cache_definition_xml).expect("parse definition xml");
    assert_eq!(doc.root_element().attribute("recordCount"), Some("5"));
    let cache_fields = doc
        .root_element()
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "cacheFields")
        .expect("cacheFields missing");
    assert_eq!(cache_fields.attribute("count"), Some("3"));
    let field_names: Vec<_> = cache_fields
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "cacheField")
        .filter_map(|n| n.attribute("name"))
        .collect();
    assert_eq!(field_names, vec!["Region", "Product", "Revenue"]);

    let cache_records_xml = std::str::from_utf8(
        pkg.part("xl/pivotCache/pivotCacheRecords1.xml")
            .expect("cache records exists"),
    )
    .expect("utf-8");
    let doc = roxmltree::Document::parse(cache_records_xml).expect("parse records xml");
    assert_eq!(doc.root_element().attribute("count"), Some("5"));

    let rows = parse_cache_records(cache_records_xml);
    assert_eq!(
        rows,
        vec![
            vec![
                CacheValue::String("East".to_string()),
                CacheValue::String("A".to_string()),
                CacheValue::Number("100".to_string())
            ],
            vec![
                CacheValue::String("East".to_string()),
                CacheValue::String("B".to_string()),
                CacheValue::Number("150".to_string())
            ],
            vec![
                CacheValue::String("West".to_string()),
                CacheValue::String("A".to_string()),
                CacheValue::Number("200".to_string())
            ],
            vec![
                CacheValue::String("West".to_string()),
                CacheValue::String("B".to_string()),
                CacheValue::Number("250".to_string())
            ],
            vec![
                CacheValue::String("North".to_string()),
                CacheValue::String("C".to_string()),
                CacheValue::Number("300".to_string())
            ],
        ]
    );
}

