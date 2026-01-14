use std::io::Cursor;

use formula_xlsx::XlsxPackage;
use pretty_assertions::assert_eq;
use quick_xml::events::Event;
use quick_xml::Reader;
use std::io::Write;
use zip::write::FileOptions;
use zip::ZipWriter;

#[derive(Debug, Clone, PartialEq, Eq)]
enum CacheValue {
    String(String),
    Number(String),
    DateTime(String),
    Bool(bool),
    Missing,
}

const REL_TYPE_SHARED_STRINGS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
const REL_TYPE_STYLES: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";

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
                    b"d" => row.push(CacheValue::DateTime(v.unwrap_or_default())),
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

    pkg.set_part(
        "xl/worksheets/sheet1.xml",
        updated_sheet.as_bytes().to_vec(),
    );

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

#[test]
fn refreshes_all_pivot_caches_from_worksheets() {
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

    pkg.refresh_all_pivot_caches_from_worksheets()
        .expect("refresh caches");

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

#[test]
fn refreshes_pivot_cache_records_with_custom_shared_strings_part() {
    let fixture = include_bytes!("fixtures/pivot-cache-refresh.xlsx");
    let mut pkg = XlsxPackage::from_bytes(fixture).expect("read fixture");

    let shared_strings_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="9" uniqueCount="9">
  <si><t>Region</t></si>
  <si><t>Product</t></si>
  <si><t>Revenue</t></si>
  <si><t>East</t></si>
  <si><t>A</t></si>
  <si><t>B</t></si>
  <si><t>West</t></si>
  <si><t>North</t></si>
  <si><t>C</t></si>
</sst>"#;
    pkg.set_part(
        "xl/custom/sharedStrings.xml",
        shared_strings_xml.as_bytes().to_vec(),
    );

    let workbook_rels_part = "xl/_rels/workbook.xml.rels";
    let mut rels_xml = std::str::from_utf8(
        pkg.part(workbook_rels_part)
            .expect("fixture should include workbook.xml.rels"),
    )
    .expect("utf-8")
    .to_string();

    let insert = format!(
        r#"  <Relationship Id="rId9999" Type="{REL_TYPE_SHARED_STRINGS}" Target="custom/sharedStrings.xml"/>"#
    );
    let close = "</Relationships>";
    let pos = rels_xml
        .rfind(close)
        .expect("workbook.xml.rels should include </Relationships>");
    rels_xml.insert_str(pos, &format!("{insert}\n"));
    pkg.set_part(workbook_rels_part, rels_xml.into_bytes());

    let updated_sheet = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1" t="s"><v>1</v></c>
      <c r="C1" t="s"><v>2</v></c>
    </row>
    <row r="2">
      <c r="A2" t="s"><v>3</v></c>
      <c r="B2" t="s"><v>4</v></c>
      <c r="C2"><v>100</v></c>
    </row>
    <row r="3">
      <c r="A3" t="s"><v>3</v></c>
      <c r="B3" t="s"><v>5</v></c>
      <c r="C3"><v>150</v></c>
    </row>
    <row r="4">
      <c r="A4" t="s"><v>6</v></c>
      <c r="B4" t="s"><v>4</v></c>
      <c r="C4"><v>200</v></c>
    </row>
    <row r="5">
      <c r="A5" t="s"><v>6</v></c>
      <c r="B5" t="s"><v>5</v></c>
      <c r="C5"><v>250</v></c>
    </row>
    <row r="6">
      <c r="A6" t="s"><v>7</v></c>
      <c r="B6" t="s"><v>8</v></c>
      <c r="C6"><v>300</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    pkg.set_part(
        "xl/worksheets/sheet1.xml",
        updated_sheet.as_bytes().to_vec(),
    );

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

#[test]
fn refreshes_pivot_cache_records_with_embedded_sheet_in_ref() {
    let fixture = include_bytes!("fixtures/pivot-cache-refresh.xlsx");
    let mut pkg = XlsxPackage::from_bytes(fixture).expect("read fixture");

    // Some producers omit `worksheetSource@sheet` and instead embed the sheet name in the ref.
    let cache_definition_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" refreshOnLoad="1" recordCount="4">
  <cacheSource type="worksheet">
    <worksheetSource ref="Sheet1!A1:C6"/>
  </cacheSource>
  <cacheFields count="3">
    <cacheField name="Region" numFmtId="0"/>
    <cacheField name="Product" numFmtId="0"/>
    <cacheField name="Sales" numFmtId="0"/>
  </cacheFields>
</pivotCacheDefinition>"#;
    pkg.set_part(
        "xl/pivotCache/pivotCacheDefinition1.xml",
        cache_definition_xml.as_bytes().to_vec(),
    );

    pkg.refresh_pivot_cache_from_worksheet("xl/pivotCache/pivotCacheDefinition1.xml")
        .expect("refresh cache");

    let cache_definition_xml = std::str::from_utf8(
        pkg.part("xl/pivotCache/pivotCacheDefinition1.xml")
            .expect("cache definition exists"),
    )
    .expect("utf-8");
    let doc = roxmltree::Document::parse(cache_definition_xml).expect("parse definition xml");
    assert_eq!(doc.root_element().attribute("recordCount"), Some("4"));

    let cache_records_xml = std::str::from_utf8(
        pkg.part("xl/pivotCache/pivotCacheRecords1.xml")
            .expect("cache records exists"),
    )
    .expect("utf-8");
    let doc = roxmltree::Document::parse(cache_records_xml).expect("parse records xml");
    assert_eq!(doc.root_element().attribute("count"), Some("4"));

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
        ]
    );
}

#[test]
fn refreshes_pivot_cache_records_with_absolute_a1_range_in_ref() {
    let fixture = include_bytes!("fixtures/pivot-cache-refresh.xlsx");
    let mut pkg = XlsxPackage::from_bytes(fixture).expect("read fixture");

    // Excel often writes A1 ranges with `$` absolute markers. Ensure we can still parse the
    // worksheetSource ref when the sheet name is embedded.
    let cache_definition_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" refreshOnLoad="1" recordCount="4">
  <cacheSource type="worksheet">
    <worksheetSource ref="Sheet1!$A$1:$C$6"/>
  </cacheSource>
  <cacheFields count="3">
    <cacheField name="Region" numFmtId="0"/>
    <cacheField name="Product" numFmtId="0"/>
    <cacheField name="Sales" numFmtId="0"/>
  </cacheFields>
</pivotCacheDefinition>"#;
    pkg.set_part(
        "xl/pivotCache/pivotCacheDefinition1.xml",
        cache_definition_xml.as_bytes().to_vec(),
    );

    pkg.refresh_pivot_cache_from_worksheet("xl/pivotCache/pivotCacheDefinition1.xml")
        .expect("refresh cache");

    let cache_definition_xml = std::str::from_utf8(
        pkg.part("xl/pivotCache/pivotCacheDefinition1.xml")
            .expect("cache definition exists"),
    )
    .expect("utf-8");
    let doc = roxmltree::Document::parse(cache_definition_xml).expect("parse definition xml");
    assert_eq!(doc.root_element().attribute("recordCount"), Some("4"));

    let cache_records_xml = std::str::from_utf8(
        pkg.part("xl/pivotCache/pivotCacheRecords1.xml")
            .expect("cache records exists"),
    )
    .expect("utf-8");
    let doc = roxmltree::Document::parse(cache_records_xml).expect("parse records xml");
    assert_eq!(doc.root_element().attribute("count"), Some("4"));

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
        ]
    );
}

#[test]
fn refreshes_pivot_cache_records_with_quoted_sheet_in_ref() {
    let fixture = include_bytes!("fixtures/pivot-cache-refresh.xlsx");
    let mut pkg = XlsxPackage::from_bytes(fixture).expect("read fixture");

    // Excel quotes sheet names with spaces/special characters using `'...'` and escapes any
    // embedded quote as `''`.
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet 1's" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;
    pkg.set_part("xl/workbook.xml", workbook_xml.as_bytes().to_vec());

    let cache_definition_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" refreshOnLoad="1" recordCount="4">
  <cacheSource type="worksheet">
    <worksheetSource ref="'Sheet 1''s'!A1:C6"/>
  </cacheSource>
  <cacheFields count="3">
    <cacheField name="Region" numFmtId="0"/>
    <cacheField name="Product" numFmtId="0"/>
    <cacheField name="Sales" numFmtId="0"/>
  </cacheFields>
</pivotCacheDefinition>"#;
    pkg.set_part(
        "xl/pivotCache/pivotCacheDefinition1.xml",
        cache_definition_xml.as_bytes().to_vec(),
    );

    pkg.refresh_pivot_cache_from_worksheet("xl/pivotCache/pivotCacheDefinition1.xml")
        .expect("refresh cache");

    let cache_definition_xml = std::str::from_utf8(
        pkg.part("xl/pivotCache/pivotCacheDefinition1.xml")
            .expect("cache definition exists"),
    )
    .expect("utf-8");
    let doc = roxmltree::Document::parse(cache_definition_xml).expect("parse definition xml");
    assert_eq!(doc.root_element().attribute("recordCount"), Some("4"));

    let cache_records_xml = std::str::from_utf8(
        pkg.part("xl/pivotCache/pivotCacheRecords1.xml")
            .expect("cache records exists"),
    )
    .expect("utf-8");
    let doc = roxmltree::Document::parse(cache_records_xml).expect("parse records xml");
    assert_eq!(doc.root_element().attribute("count"), Some("4"));
}

fn build_date_pivot_fixture(date1904: bool) -> Vec<u8> {
    let workbook_pr = if date1904 {
        r#"<workbookPr date1904="1"/>"#
    } else {
        ""
    };

    let workbook_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  {workbook_pr}
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#
    );

    let workbook_rels = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="{REL_TYPE_STYLES}" Target="styles.xml"/>
</Relationships>"#
    );

    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="165" formatCode="yyyy-mm-dd"/>
  </numFmts>
</styleSheet>"#;

    let (serial1, serial2) = if date1904 { (0, 1) } else { (1, 2) };
    let worksheet_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Date</t></is></c>
      <c r="B1" t="inlineStr"><is><t>Value</t></is></c>
    </row>
    <row r="2">
      <c r="A2"><v>{serial1}</v></c>
      <c r="B2"><v>10</v></c>
    </row>
    <row r="3">
      <c r="A3"><v>{serial2}</v></c>
      <c r="B3"><v>20</v></c>
    </row>
  </sheetData>
</worksheet>"#
    );

    let cache_definition_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" recordCount="0">
  <cacheSource type="worksheet">
    <worksheetSource sheet="Sheet1" ref="A1:B3"/>
  </cacheSource>
  <cacheFields count="2">
    <cacheField name="Date" numFmtId="165"/>
    <cacheField name="Value" numFmtId="0"/>
  </cacheFields>
</pivotCacheDefinition>"#;

    let cache_definition_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords1.xml"/>
</Relationships>"#;

    let cache_records_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0"/>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file("xl/styles.xml", options).unwrap();
    zip.write_all(styles_xml.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(worksheet_xml.as_bytes()).unwrap();

    zip.start_file("xl/pivotCache/pivotCacheDefinition1.xml", options)
        .unwrap();
    zip.write_all(cache_definition_xml.as_bytes()).unwrap();

    zip.start_file(
        "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels",
        options,
    )
    .unwrap();
    zip.write_all(cache_definition_rels.as_bytes()).unwrap();

    zip.start_file("xl/pivotCache/pivotCacheRecords1.xml", options)
        .unwrap();
    zip.write_all(cache_records_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

#[test]
fn refreshes_pivot_cache_records_with_date_time_fields_as_d_tags() {
    let fixture = build_date_pivot_fixture(false);
    let mut pkg = XlsxPackage::from_bytes(&fixture).expect("read pkg");

    pkg.refresh_pivot_cache_from_worksheet("xl/pivotCache/pivotCacheDefinition1.xml")
        .expect("refresh cache");

    let cache_definition_xml = std::str::from_utf8(
        pkg.part("xl/pivotCache/pivotCacheDefinition1.xml")
            .expect("cache definition exists"),
    )
    .expect("utf-8");
    let doc = roxmltree::Document::parse(cache_definition_xml).expect("parse definition xml");
    assert_eq!(doc.root_element().attribute("recordCount"), Some("2"));

    let cache_records_xml = std::str::from_utf8(
        pkg.part("xl/pivotCache/pivotCacheRecords1.xml")
            .expect("cache records exists"),
    )
    .expect("utf-8");
    let doc = roxmltree::Document::parse(cache_records_xml).expect("parse records xml");
    assert_eq!(doc.root_element().attribute("count"), Some("2"));

    let rows = parse_cache_records(cache_records_xml);
    assert_eq!(
        rows,
        vec![
            vec![
                CacheValue::DateTime("1900-01-01T00:00:00Z".to_string()),
                CacheValue::Number("10".to_string())
            ],
            vec![
                CacheValue::DateTime("1900-01-02T00:00:00Z".to_string()),
                CacheValue::Number("20".to_string())
            ],
        ]
    );
}

#[test]
fn refreshes_pivot_cache_records_uses_workbook_date1904_for_datetime_serials() {
    let fixture = build_date_pivot_fixture(true);
    let mut pkg = XlsxPackage::from_bytes(&fixture).expect("read pkg");

    pkg.refresh_pivot_cache_from_worksheet("xl/pivotCache/pivotCacheDefinition1.xml")
        .expect("refresh cache");

    let cache_records_xml = std::str::from_utf8(
        pkg.part("xl/pivotCache/pivotCacheRecords1.xml")
            .expect("cache records exists"),
    )
    .expect("utf-8");

    let rows = parse_cache_records(cache_records_xml);
    assert_eq!(
        rows,
        vec![
            vec![
                CacheValue::DateTime("1904-01-01T00:00:00Z".to_string()),
                CacheValue::Number("10".to_string())
            ],
            vec![
                CacheValue::DateTime("1904-01-02T00:00:00Z".to_string()),
                CacheValue::Number("20".to_string())
            ],
        ]
    );
}
