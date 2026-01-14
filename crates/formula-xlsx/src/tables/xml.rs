use formula_model::table::{AutoFilter, Table, TableColumn, TableStyleInfo};
use formula_model::{normalize_formula_text, Range};
use quick_xml::de::from_str;
use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::Writer;
use serde::Deserialize;
use std::io::{Cursor, Write};

#[derive(Debug, Deserialize)]
#[serde(rename = "table")]
struct TableXml {
    #[serde(rename = "@id")]
    id: u32,
    #[serde(rename = "@name")]
    name: String,
    #[serde(rename = "@displayName")]
    display_name: String,
    #[serde(rename = "@ref")]
    reference: String,
    #[serde(rename = "@headerRowCount")]
    header_row_count: Option<u32>,
    #[serde(rename = "@totalsRowCount")]
    totals_row_count: Option<u32>,
    #[serde(rename = "tableColumns")]
    table_columns: TableColumnsXml,
    #[serde(rename = "tableStyleInfo")]
    style_info: Option<TableStyleInfoXml>,
}

#[derive(Debug, Deserialize)]
struct TableColumnsXml {
    #[serde(rename = "@count")]
    _count: Option<u32>,
    #[serde(rename = "tableColumn", default)]
    columns: Vec<TableColumnXml>,
}

#[derive(Debug, Deserialize)]
struct TableColumnXml {
    #[serde(rename = "@id")]
    id: u32,
    #[serde(rename = "@name")]
    name: String,
    #[serde(rename = "calculatedColumnFormula")]
    calculated_column_formula: Option<String>,
    #[serde(rename = "totalsRowFormula")]
    totals_row_formula: Option<String>,
}

#[derive(Debug, Deserialize)]
struct TableStyleInfoXml {
    #[serde(rename = "@name")]
    name: String,
    #[serde(rename = "@showFirstColumn")]
    show_first_column: Option<u8>,
    #[serde(rename = "@showLastColumn")]
    show_last_column: Option<u8>,
    #[serde(rename = "@showRowStripes")]
    show_row_stripes: Option<u8>,
    #[serde(rename = "@showColumnStripes")]
    show_column_stripes: Option<u8>,
}

pub fn parse_table(xml: &str) -> Result<Table, String> {
    let table: TableXml = from_str(xml).map_err(|e| e.to_string())?;

    let range = Range::from_a1(&table.reference).map_err(|e| e.to_string())?;

    let columns = table
        .table_columns
        .columns
        .into_iter()
        .map(|c| TableColumn {
            id: c.id,
            name: c.name,
            formula: c
                .calculated_column_formula
                .as_deref()
                .and_then(normalize_formula_text)
                .map(|f| crate::formula_text::strip_xlfn_prefixes(&f)),
            totals_formula: c
                .totals_row_formula
                .as_deref()
                .and_then(normalize_formula_text)
                .map(|f| crate::formula_text::strip_xlfn_prefixes(&f)),
        })
        .collect();

    let style = table.style_info.map(|s| TableStyleInfo {
        name: s.name,
        show_first_column: s.show_first_column.unwrap_or(0) != 0,
        show_last_column: s.show_last_column.unwrap_or(0) != 0,
        show_row_stripes: s.show_row_stripes.unwrap_or(0) != 0,
        show_column_stripes: s.show_column_stripes.unwrap_or(0) != 0,
    });

    let auto_filter: Option<AutoFilter> = crate::autofilter::parse_worksheet_autofilter(xml)
        .map_err(|e| e.to_string())?;

    Ok(Table {
        id: table.id,
        name: table.name,
        display_name: table.display_name,
        range,
        header_row_count: table.header_row_count.unwrap_or(1),
        totals_row_count: table.totals_row_count.unwrap_or(0),
        columns,
        style,
        auto_filter,
        relationship_id: None,
        part_path: None,
    })
}

pub fn write_table_xml(table: &Table) -> Result<String, String> {
    fn write_text_element<W: Write>(
        writer: &mut Writer<W>,
        name: &str,
        text: &str,
    ) -> Result<(), quick_xml::Error> {
        writer.write_event(Event::Start(BytesStart::new(name)))?;
        writer.write_event(Event::Text(BytesText::new(text)))?;
        writer.write_event(Event::End(BytesEnd::new(name)))?;
        Ok(())
    }

    let mut writer = Writer::new(Cursor::new(Vec::new()));

    let id = table.id.to_string();
    let reference = table.range.to_string();
    let header_row_count = table.header_row_count.to_string();
    let totals_row_count = table.totals_row_count.to_string();

    let mut table_tag = BytesStart::new("table");
    table_tag.push_attribute(("xmlns", crate::xml::SPREADSHEETML_NS));
    table_tag.push_attribute(("id", id.as_str()));
    table_tag.push_attribute(("name", table.name.as_str()));
    table_tag.push_attribute(("displayName", table.display_name.as_str()));
    table_tag.push_attribute(("ref", reference.as_str()));
    table_tag.push_attribute(("headerRowCount", header_row_count.as_str()));
    table_tag.push_attribute(("totalsRowCount", totals_row_count.as_str()));
    writer
        .write_event(Event::Start(table_tag))
        .map_err(|e| e.to_string())?;

    if let Some(filter) = &table.auto_filter {
        let autofilter_xml =
            crate::autofilter::write_autofilter(filter).map_err(|e| e.to_string())?;
        writer
            .get_mut()
            .write_all(autofilter_xml.as_bytes())
            .map_err(|e| e.to_string())?;
    }

    let col_count = (table.columns.len() as u32).to_string();
    let mut table_columns = BytesStart::new("tableColumns");
    table_columns.push_attribute(("count", col_count.as_str()));
    writer
        .write_event(Event::Start(table_columns))
        .map_err(|e| e.to_string())?;

    for col in &table.columns {
        let col_id = col.id.to_string();
        let mut table_column = BytesStart::new("tableColumn");
        table_column.push_attribute(("id", col_id.as_str()));
        table_column.push_attribute(("name", col.name.as_str()));

        let calculated_column_formula = col
            .formula
            .as_deref()
            .and_then(normalize_formula_text)
            .map(|formula| crate::formula_text::add_xlfn_prefixes(&formula));
        let totals_row_formula = col
            .totals_formula
            .as_deref()
            .and_then(normalize_formula_text)
            .map(|formula| crate::formula_text::add_xlfn_prefixes(&formula));

        if calculated_column_formula.is_none() && totals_row_formula.is_none() {
            writer
                .write_event(Event::Empty(table_column))
                .map_err(|e| e.to_string())?;
            continue;
        }

        writer
            .write_event(Event::Start(table_column))
            .map_err(|e| e.to_string())?;

        if let Some(formula) = calculated_column_formula {
            write_text_element(&mut writer, "calculatedColumnFormula", &formula)
                .map_err(|e| e.to_string())?;
        }
        if let Some(formula) = totals_row_formula {
            write_text_element(&mut writer, "totalsRowFormula", &formula).map_err(|e| e.to_string())?;
        }

        writer
            .write_event(Event::End(BytesEnd::new("tableColumn")))
            .map_err(|e| e.to_string())?;
    }

    writer
        .write_event(Event::End(BytesEnd::new("tableColumns")))
        .map_err(|e| e.to_string())?;

    if let Some(style) = &table.style {
        let show_first_column = (style.show_first_column as u8).to_string();
        let show_last_column = (style.show_last_column as u8).to_string();
        let show_row_stripes = (style.show_row_stripes as u8).to_string();
        let show_column_stripes = (style.show_column_stripes as u8).to_string();

        let mut style_info = BytesStart::new("tableStyleInfo");
        style_info.push_attribute(("name", style.name.as_str()));
        style_info.push_attribute(("showFirstColumn", show_first_column.as_str()));
        style_info.push_attribute(("showLastColumn", show_last_column.as_str()));
        style_info.push_attribute(("showRowStripes", show_row_stripes.as_str()));
        style_info.push_attribute(("showColumnStripes", show_column_stripes.as_str()));

        writer
            .write_event(Event::Empty(style_info))
            .map_err(|e| e.to_string())?;
    }

    writer
        .write_event(Event::End(BytesEnd::new("table")))
        .map_err(|e| e.to_string())?;

    let bytes = writer.into_inner().into_inner();
    Ok(String::from_utf8_lossy(&bytes).to_string())
}

#[cfg(test)]
mod tests {
    use super::*;
    use formula_model::table::TableArea;
    use formula_model::table::{Table, TableColumn};
    use formula_model::autofilter::{DateComparison, NumberComparison};
    use formula_model::{FilterCriterion, FilterJoin};

    #[test]
    fn round_trips_table_xml() {
        let table = parse_table(include_str!("../../tests/fixtures/table1.xml")).unwrap();
        assert_eq!(table.name, "Table1");
        assert_eq!(table.columns.len(), 4);
        assert_eq!(
            table.column_range_in_area("Qty", TableArea::Data).unwrap().to_string(),
            "B2:B4"
        );

        let out = write_table_xml(&table).unwrap();
        let reparsed = parse_table(&out).unwrap();
        assert_eq!(reparsed, table);
    }

    #[test]
    fn round_trips_table_autofilter_advanced_and_preserves_raw_xml() {
        let xml = include_str!("../../tests/fixtures/table_autofilter_advanced.xml");
        let table = parse_table(xml).unwrap();
        let filter = table.auto_filter.as_ref().expect("expected autoFilter");
        assert_eq!(filter.filter_columns.len(), 3);

        let col0 = &filter.filter_columns[0];
        assert_eq!(col0.col_id, 0);
        assert_eq!(col0.join, FilterJoin::All);
        assert_eq!(
            col0.criteria,
            vec![FilterCriterion::Number(NumberComparison::GreaterThan(10.0))]
        );

        let col1 = &filter.filter_columns[1];
        assert_eq!(col1.col_id, 1);
        assert_eq!(
            col1.criteria,
            vec![FilterCriterion::Date(DateComparison::Today)]
        );

        let col2 = &filter.filter_columns[2];
        assert_eq!(col2.col_id, 2);
        assert!(
            col2.raw_xml.iter().any(|x| x.contains("<top10")),
            "expected top10 to be captured into raw_xml, got: {:?}",
            col2.raw_xml
        );

        assert!(
            filter.raw_xml.iter().any(|x| x.contains("<extLst")),
            "expected extLst to be captured into autoFilter raw_xml, got: {:?}",
            filter.raw_xml
        );
        let sort_state = filter.sort_state.as_ref().expect("expected sortState");
        assert_eq!(sort_state.conditions.len(), 1);
        assert_eq!(sort_state.conditions[0].range.to_string(), "A2:A10");
        assert!(sort_state.conditions[0].descending);

        let written = write_table_xml(&table).unwrap();
        assert!(written.contains("<customFilters"), "missing customFilters: {written}");
        assert!(
            written.contains(r#"operator="greaterThan""#),
            "missing customFilter operator: {written}"
        );
        assert!(
            written.contains("<dynamicFilter") && written.contains(r#"type="today""#),
            "missing dynamicFilter today: {written}"
        );
        assert!(
            written.contains(r#"<filterColumn colId="2"><top10"#),
            "expected raw_xml top10 to be the primary filter element for colId=2, got: {written}"
        );
        assert!(written.contains("<top10"), "missing top10 raw_xml: {written}");
        assert!(written.contains("<extLst"), "missing extLst raw_xml: {written}");

        let reparsed = parse_table(&written).unwrap();
        assert_eq!(reparsed, table);
    }

    #[test]
    fn parse_table_strips_leading_equals_and_xlfn_prefixes_in_column_formulas() {
        let xml = r#"<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="A1:A2" headerRowCount="1" totalsRowCount="0"><tableColumns count="1"><tableColumn id="1" name="Seq"><calculatedColumnFormula>=_xlfn.SEQUENCE(2)</calculatedColumnFormula><totalsRowFormula>=_xlfn.ISFORMULA(A1)</totalsRowFormula></tableColumn></tableColumns></table>"#;
        let table = parse_table(xml).unwrap();
        assert_eq!(table.columns.len(), 1);
        assert_eq!(table.columns[0].formula.as_deref(), Some("SEQUENCE(2)"));
        assert_eq!(table.columns[0].totals_formula.as_deref(), Some("ISFORMULA(A1)"));
    }

    #[test]
    fn write_table_xml_prefixes_xlfn_functions_in_column_formulas() {
        let table = Table {
            id: 1,
            name: "Table1".to_string(),
            display_name: "Table1".to_string(),
            range: Range::from_a1("A1:B2").expect("range"),
            header_row_count: 1,
            totals_row_count: 0,
            columns: vec![
                TableColumn {
                    id: 1,
                    name: "Seq".to_string(),
                    formula: Some("SEQUENCE(2)".to_string()),
                    totals_formula: None,
                },
                TableColumn {
                    id: 2,
                    name: "IsFormula".to_string(),
                    formula: None,
                    totals_formula: Some("ISFORMULA(A1)".to_string()),
                },
            ],
            style: None,
            auto_filter: None,
            relationship_id: None,
            part_path: None,
        };

        let xml = write_table_xml(&table).expect("write table xml");
        assert!(
            xml.contains("<calculatedColumnFormula>_xlfn.SEQUENCE(2)</calculatedColumnFormula>"),
            "expected calculatedColumnFormula to be prefixed, got xml: {xml}"
        );
        assert!(
            xml.contains("<totalsRowFormula>_xlfn.ISFORMULA(A1)</totalsRowFormula>"),
            "expected totalsRowFormula to be prefixed, got xml: {xml}"
        );
    }

    #[test]
    fn write_table_xml_escapes_special_chars_in_column_names_and_formulas() {
        let table = Table {
            id: 1,
            name: "Table1".to_string(),
            display_name: "Table1".to_string(),
            range: Range::from_a1("A1:A2").expect("range"),
            header_row_count: 1,
            totals_row_count: 0,
            columns: vec![TableColumn {
                id: 1,
                name: "A&B".to_string(),
                formula: Some("A1<5".to_string()),
                totals_formula: Some("IF(A1<5,1,0)".to_string()),
            }],
            style: None,
            auto_filter: None,
            relationship_id: None,
            part_path: None,
        };

        let xml = write_table_xml(&table).expect("write table xml");
        assert!(
            xml.contains(r#"name="A&amp;B""#),
            "expected tableColumn/@name to be XML-escaped, got xml: {xml}"
        );
        assert!(
            xml.contains("<calculatedColumnFormula>A1&lt;5</calculatedColumnFormula>"),
            "expected calculatedColumnFormula to be XML-escaped, got xml: {xml}"
        );
        assert!(
            xml.contains("<totalsRowFormula>IF(A1&lt;5,1,0)</totalsRowFormula>"),
            "expected totalsRowFormula to be XML-escaped, got xml: {xml}"
        );
    }
}
