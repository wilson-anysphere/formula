use crate::engine::{DaxError, DaxResult, FilterContext, RowContext};
use crate::parser::Expr;
use crate::value::Value;
use std::collections::{HashMap, HashSet};

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Cardinality {
    OneToMany,
    OneToOne,
    ManyToMany,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum CrossFilterDirection {
    Single,
    Both,
}

#[derive(Clone, Debug)]
pub struct Relationship {
    pub name: String,
    pub from_table: String,
    pub from_column: String,
    pub to_table: String,
    pub to_column: String,
    pub cardinality: Cardinality,
    pub cross_filter_direction: CrossFilterDirection,
    pub is_active: bool,
    pub enforce_referential_integrity: bool,
}

#[derive(Clone, Debug)]
pub struct Measure {
    pub name: String,
    pub expression: String,
    pub(crate) parsed: Expr,
}

#[derive(Clone, Debug)]
pub struct CalculatedColumn {
    pub table: String,
    pub name: String,
    pub expression: String,
    pub parsed: Expr,
}

#[derive(Clone, Debug)]
pub struct Table {
    name: String,
    columns: Vec<String>,
    column_index: HashMap<String, usize>,
    rows: Vec<Vec<Value>>,
}

impl Table {
    pub fn new(name: impl Into<String>, columns: Vec<impl Into<String>>) -> Self {
        let name = name.into();
        let columns: Vec<String> = columns.into_iter().map(Into::into).collect();
        let column_index = columns
            .iter()
            .enumerate()
            .map(|(idx, c)| (c.clone(), idx))
            .collect();

        Self {
            name,
            columns,
            column_index,
            rows: Vec::new(),
        }
    }

    pub fn name(&self) -> &str {
        &self.name
    }

    pub fn columns(&self) -> &[String] {
        &self.columns
    }

    pub fn row_count(&self) -> usize {
        self.rows.len()
    }

    pub fn push_row(&mut self, row: Vec<Value>) -> DaxResult<()> {
        if row.len() != self.columns.len() {
            return Err(DaxError::SchemaMismatch {
                table: self.name.clone(),
                expected: self.columns.len(),
                actual: row.len(),
            });
        }

        self.rows.push(row);
        Ok(())
    }

    pub(crate) fn column_idx(&self, column: &str) -> Option<usize> {
        self.column_index.get(column).copied()
    }

    pub fn value(&self, row: usize, column: &str) -> Option<&Value> {
        let idx = self.column_idx(column)?;
        self.rows.get(row)?.get(idx)
    }

    pub(crate) fn value_by_idx(&self, row: usize, idx: usize) -> Option<&Value> {
        self.rows.get(row)?.get(idx)
    }

    pub(crate) fn add_column(
        &mut self,
        name: impl Into<String>,
        values: Vec<Value>,
    ) -> DaxResult<()> {
        let name = name.into();
        if self.column_index.contains_key(&name) {
            return Err(DaxError::DuplicateColumn {
                table: self.name.clone(),
                column: name,
            });
        }
        if values.len() != self.rows.len() {
            return Err(DaxError::ColumnLengthMismatch {
                table: self.name.clone(),
                column: name,
                expected: self.rows.len(),
                actual: values.len(),
            });
        }

        let idx = self.columns.len();
        self.columns.push(name.clone());
        self.column_index.insert(name, idx);
        for (row, value) in self.rows.iter_mut().zip(values) {
            row.push(value);
        }
        Ok(())
    }
}

#[derive(Clone, Debug)]
pub struct DataModel {
    pub(crate) tables: HashMap<String, Table>,
    pub(crate) relationships: Vec<RelationshipInfo>,
    pub(crate) measures: HashMap<String, Measure>,
    pub(crate) calculated_columns: Vec<CalculatedColumn>,
}

#[derive(Clone, Debug)]
pub(crate) struct RelationshipInfo {
    pub(crate) rel: Relationship,
    pub(crate) to_index: HashMap<Value, usize>,
}

impl DataModel {
    pub fn new() -> Self {
        Self {
            tables: HashMap::new(),
            relationships: Vec::new(),
            measures: HashMap::new(),
            calculated_columns: Vec::new(),
        }
    }

    pub fn table(&self, name: &str) -> Option<&Table> {
        self.tables.get(name)
    }

    pub fn table_mut(&mut self, name: &str) -> Option<&mut Table> {
        self.tables.get_mut(name)
    }

    pub fn add_table(&mut self, table: Table) -> DaxResult<()> {
        let name = table.name.clone();
        if self.tables.contains_key(&name) {
            return Err(DaxError::DuplicateTable { table: name });
        }
        self.tables.insert(name, table);
        Ok(())
    }

    pub fn add_relationship(&mut self, relationship: Relationship) -> DaxResult<()> {
        let from_table = self
            .tables
            .get(&relationship.from_table)
            .ok_or_else(|| DaxError::UnknownTable(relationship.from_table.clone()))?;
        let to_table = self
            .tables
            .get(&relationship.to_table)
            .ok_or_else(|| DaxError::UnknownTable(relationship.to_table.clone()))?;

        let from_col = relationship.from_column.clone();
        let to_col = relationship.to_column.clone();

        let from_idx = from_table
            .column_idx(&from_col)
            .ok_or_else(|| DaxError::UnknownColumn {
                table: relationship.from_table.clone(),
                column: from_col.clone(),
            })?;
        let to_idx = to_table
            .column_idx(&to_col)
            .ok_or_else(|| DaxError::UnknownColumn {
                table: relationship.to_table.clone(),
                column: to_col.clone(),
            })?;

        if relationship.cardinality != Cardinality::OneToMany {
            return Err(DaxError::UnsupportedCardinality {
                relationship: relationship.name.clone(),
                cardinality: relationship.cardinality,
            });
        }

        let mut to_index = HashMap::<Value, usize>::new();
        for row in 0..to_table.row_count() {
            let Some(value) = to_table.value_by_idx(row, to_idx) else {
                continue;
            };
            if to_index.insert(value.clone(), row).is_some() {
                return Err(DaxError::NonUniqueKey {
                    table: relationship.to_table.clone(),
                    column: to_col.clone(),
                    value: value.clone(),
                });
            }
        }

        if relationship.enforce_referential_integrity {
            let to_values: HashSet<Value> = to_index.keys().cloned().collect();
            for row in 0..from_table.row_count() {
                let Some(value) = from_table.value_by_idx(row, from_idx) else {
                    continue;
                };
                if value.is_blank() {
                    continue;
                }
                if !to_values.contains(value) {
                    return Err(DaxError::ReferentialIntegrityViolation {
                        relationship: relationship.name.clone(),
                        from_table: relationship.from_table.clone(),
                        from_column: from_col.clone(),
                        to_table: relationship.to_table.clone(),
                        to_column: to_col.clone(),
                        value: value.clone(),
                    });
                }
            }
        }

        self.relationships.push(RelationshipInfo {
            rel: relationship,
            to_index,
        });
        Ok(())
    }

    pub fn add_measure(
        &mut self,
        name: impl Into<String>,
        expression: impl Into<String>,
    ) -> DaxResult<()> {
        let name = name.into();
        if self.measures.contains_key(&name) {
            return Err(DaxError::DuplicateMeasure { measure: name });
        }
        let expression = expression.into();
        let parsed = crate::parser::parse(&expression)?;
        self.measures.insert(
            name.clone(),
            Measure {
                name,
                expression,
                parsed,
            },
        );
        Ok(())
    }

    pub fn add_calculated_column(
        &mut self,
        table: impl Into<String>,
        name: impl Into<String>,
        expression: impl Into<String>,
    ) -> DaxResult<()> {
        let table = table.into();
        let name = name.into();
        let expression = expression.into();

        let parsed = crate::parser::parse(&expression)?;
        let calc = CalculatedColumn {
            table: table.clone(),
            name: name.clone(),
            expression,
            parsed: parsed.clone(),
        };

        let values = {
            let Some(table_ref) = self.tables.get(&table) else {
                return Err(DaxError::UnknownTable(table.clone()));
            };

            let mut results = Vec::with_capacity(table_ref.row_count());
            for row in 0..table_ref.row_count() {
                let mut row_ctx = RowContext::default();
                row_ctx.push(table_ref.name(), row);
                let value = crate::engine::DaxEngine::new().evaluate_expr(
                    self,
                    &parsed,
                    &FilterContext::default(),
                    &row_ctx,
                )?;
                results.push(value);
            }
            results
        };

        let table_mut = self
            .tables
            .get_mut(&table)
            .ok_or_else(|| DaxError::UnknownTable(table.clone()))?;
        table_mut.add_column(name, values)?;

        self.calculated_columns.push(calc);
        Ok(())
    }

    pub fn evaluate_measure(&self, name: &str, filter: &FilterContext) -> DaxResult<Value> {
        let measure = self
            .measures
            .get(Self::normalize_measure_name(name))
            .ok_or_else(|| DaxError::UnknownMeasure(name.to_string()))?;
        crate::engine::DaxEngine::new().evaluate_expr(
            self,
            &measure.parsed,
            filter,
            &RowContext::default(),
        )
    }

    pub(crate) fn measures(&self) -> &HashMap<String, Measure> {
        &self.measures
    }

    pub(crate) fn relationships(&self) -> &[RelationshipInfo] {
        &self.relationships
    }

    pub(crate) fn normalize_measure_name(name: &str) -> &str {
        name.strip_prefix('[')
            .and_then(|n| n.strip_suffix(']'))
            .unwrap_or(name)
            .trim()
    }
}
