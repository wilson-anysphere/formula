use formula_model::table::{
    AutoFilter, FilterColumn, SortCondition, SortState, Table, TableColumn, TableStyleInfo,
};
use formula_model::Range;
use quick_xml::{de::from_str, se::to_string};
use serde::{Deserialize, Serialize};

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
    #[serde(rename = "autoFilter")]
    auto_filter: Option<AutoFilterXml>,
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

#[derive(Debug, Deserialize)]
struct AutoFilterXml {
    #[serde(rename = "@ref")]
    reference: String,
    #[serde(rename = "filterColumn", default)]
    filter_columns: Vec<FilterColumnXml>,
    #[serde(rename = "sortState")]
    sort_state: Option<SortStateXml>,
}

#[derive(Debug, Deserialize)]
struct FilterColumnXml {
    #[serde(rename = "@colId")]
    col_id: u32,
    #[serde(rename = "filters")]
    filters: Option<FiltersXml>,
}

#[derive(Debug, Deserialize)]
struct FiltersXml {
    #[serde(rename = "filter", default)]
    filters: Vec<FilterXml>,
}

#[derive(Debug, Deserialize)]
struct FilterXml {
    #[serde(rename = "@val")]
    val: String,
}

#[derive(Debug, Deserialize)]
struct SortStateXml {
    #[serde(rename = "sortCondition", default)]
    sort_conditions: Vec<SortConditionXml>,
}

#[derive(Debug, Deserialize)]
struct SortConditionXml {
    #[serde(rename = "@ref")]
    reference: String,
    #[serde(rename = "@descending")]
    descending: Option<u8>,
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
            formula: c.calculated_column_formula,
            totals_formula: c.totals_row_formula,
        })
        .collect();

    let style = table.style_info.map(|s| TableStyleInfo {
        name: s.name,
        show_first_column: s.show_first_column.unwrap_or(0) != 0,
        show_last_column: s.show_last_column.unwrap_or(0) != 0,
        show_row_stripes: s.show_row_stripes.unwrap_or(0) != 0,
        show_column_stripes: s.show_column_stripes.unwrap_or(0) != 0,
    });

    let auto_filter = table.auto_filter.map(|af| {
        let range = Range::from_a1(&af.reference).map_err(|e| e.to_string())?;
        let filter_columns = af
            .filter_columns
            .into_iter()
            .map(|fc| FilterColumn {
                col_id: fc.col_id,
                values: fc
                    .filters
                    .map(|f| f.filters.into_iter().map(|x| x.val).collect())
                    .unwrap_or_default(),
            })
            .collect();
        let sort_state = af.sort_state.map(|s| SortState {
            conditions: s
                .sort_conditions
                .into_iter()
                .filter_map(|c| {
                    let range = Range::from_a1(&c.reference).ok()?;
                    Some(SortCondition {
                        range,
                        descending: c.descending.unwrap_or(0) != 0,
                    })
                })
                .collect(),
        });
        Ok::<_, String>(AutoFilter {
            range,
            filter_columns,
            sort_state,
        })
    });

    let auto_filter = auto_filter.transpose()?;

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

#[derive(Debug, Serialize)]
#[serde(rename = "table")]
struct TableXmlOut<'a> {
    #[serde(rename = "@xmlns")]
    xmlns: &'a str,
    #[serde(rename = "@id")]
    id: u32,
    #[serde(rename = "@name")]
    name: &'a str,
    #[serde(rename = "@displayName")]
    display_name: &'a str,
    #[serde(rename = "@ref")]
    reference: String,
    #[serde(rename = "@headerRowCount", skip_serializing_if = "Option::is_none")]
    header_row_count: Option<u32>,
    #[serde(rename = "@totalsRowCount", skip_serializing_if = "Option::is_none")]
    totals_row_count: Option<u32>,
    #[serde(rename = "autoFilter", skip_serializing_if = "Option::is_none")]
    auto_filter: Option<AutoFilterXmlOut>,
    #[serde(rename = "tableColumns")]
    table_columns: TableColumnsXmlOut<'a>,
    #[serde(rename = "tableStyleInfo", skip_serializing_if = "Option::is_none")]
    style_info: Option<TableStyleInfoXmlOut<'a>>,
}

#[derive(Debug, Serialize)]
struct TableColumnsXmlOut<'a> {
    #[serde(rename = "@count")]
    count: u32,
    #[serde(rename = "tableColumn")]
    columns: Vec<TableColumnXmlOut<'a>>,
}

#[derive(Debug, Serialize)]
struct TableColumnXmlOut<'a> {
    #[serde(rename = "@id")]
    id: u32,
    #[serde(rename = "@name")]
    name: &'a str,
    #[serde(rename = "calculatedColumnFormula", skip_serializing_if = "Option::is_none")]
    calculated_column_formula: Option<&'a str>,
    #[serde(rename = "totalsRowFormula", skip_serializing_if = "Option::is_none")]
    totals_row_formula: Option<&'a str>,
}

#[derive(Debug, Serialize)]
struct TableStyleInfoXmlOut<'a> {
    #[serde(rename = "@name")]
    name: &'a str,
    #[serde(rename = "@showFirstColumn")]
    show_first_column: u8,
    #[serde(rename = "@showLastColumn")]
    show_last_column: u8,
    #[serde(rename = "@showRowStripes")]
    show_row_stripes: u8,
    #[serde(rename = "@showColumnStripes")]
    show_column_stripes: u8,
}

#[derive(Debug, Serialize)]
struct AutoFilterXmlOut {
    #[serde(rename = "@ref")]
    reference: String,
    #[serde(rename = "filterColumn", skip_serializing_if = "Vec::is_empty", default)]
    filter_columns: Vec<FilterColumnXmlOut>,
    #[serde(rename = "sortState", skip_serializing_if = "Option::is_none")]
    sort_state: Option<SortStateXmlOut>,
}

#[derive(Debug, Serialize)]
struct FilterColumnXmlOut {
    #[serde(rename = "@colId")]
    col_id: u32,
    #[serde(rename = "filters", skip_serializing_if = "Option::is_none")]
    filters: Option<FiltersXmlOut>,
}

#[derive(Debug, Serialize)]
struct FiltersXmlOut {
    #[serde(rename = "filter", skip_serializing_if = "Vec::is_empty", default)]
    filters: Vec<FilterXmlOut>,
}

#[derive(Debug, Serialize)]
struct FilterXmlOut {
    #[serde(rename = "@val")]
    val: String,
}

#[derive(Debug, Serialize)]
struct SortStateXmlOut {
    #[serde(rename = "sortCondition", skip_serializing_if = "Vec::is_empty", default)]
    sort_conditions: Vec<SortConditionXmlOut>,
}

#[derive(Debug, Serialize)]
struct SortConditionXmlOut {
    #[serde(rename = "@ref")]
    reference: String,
    #[serde(rename = "@descending", skip_serializing_if = "Option::is_none")]
    descending: Option<u8>,
}

pub fn write_table_xml(table: &Table) -> Result<String, String> {
    let xml = TableXmlOut {
        xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        id: table.id,
        name: &table.name,
        display_name: &table.display_name,
        reference: table.range.to_string(),
        header_row_count: Some(table.header_row_count),
        totals_row_count: Some(table.totals_row_count),
        auto_filter: table.auto_filter.as_ref().map(|af| AutoFilterXmlOut {
            reference: af.range.to_string(),
            filter_columns: af
                .filter_columns
                .iter()
                .map(|fc| FilterColumnXmlOut {
                    col_id: fc.col_id,
                    filters: if fc.values.is_empty() {
                        None
                    } else {
                        Some(FiltersXmlOut {
                            filters: fc
                                .values
                                .iter()
                                .map(|v| FilterXmlOut { val: v.clone() })
                                .collect(),
                        })
                    },
                })
                .collect(),
            sort_state: af.sort_state.as_ref().map(|s| SortStateXmlOut {
                sort_conditions: s
                    .conditions
                    .iter()
                    .map(|c| SortConditionXmlOut {
                        reference: c.range.to_string(),
                        descending: if c.descending { Some(1) } else { None },
                    })
                    .collect(),
            }),
        }),
        table_columns: TableColumnsXmlOut {
            count: table.columns.len() as u32,
            columns: table
                .columns
                .iter()
                .map(|c| TableColumnXmlOut {
                    id: c.id,
                    name: &c.name,
                    calculated_column_formula: c.formula.as_deref(),
                    totals_row_formula: c.totals_formula.as_deref(),
                })
                .collect(),
        },
        style_info: table.style.as_ref().map(|s| TableStyleInfoXmlOut {
            name: &s.name,
            show_first_column: s.show_first_column as u8,
            show_last_column: s.show_last_column as u8,
            show_row_stripes: s.show_row_stripes as u8,
            show_column_stripes: s.show_column_stripes as u8,
        }),
    };

    to_string(&xml).map_err(|e| e.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;
    use formula_model::table::TableArea;

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
}
