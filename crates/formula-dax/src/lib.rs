//! A small (but growing) DAX engine intended to back Formula's Power Pivot / data model features.
//!
//! The core entry points are:
//! - [`DataModel`] to store tables, relationships, measures, and calculated columns
//! - [`DaxEngine`] to evaluate arbitrary DAX expressions
//! - [`pivot`] to produce grouped results (group keys + measures) for pivot tables
//!
//! ## Columnar storage
//! For large models, use [`Table::from_columnar`] to attach a
//! [`formula_columnar::ColumnarTable`] as the storage backend. The engine can use column
//! statistics and dictionary encoding to speed up common aggregations.
//!
//! Calculated columns created via [`DataModel::add_calculated_column`] work for both in-memory and
//! columnar-backed tables. For columnar tables, calculated columns are computed eagerly and
//! materialized by appending a new encoded column to the underlying
//! [`formula_columnar::ColumnarTable`].
//!
//! Columnar tables are treated as immutable snapshots; internally this is implemented via
//! `Arc` copy-on-write and the table may be cloned when the `Arc` is not uniquely owned.
//!
//! ## Quick example
//! ```rust
//! use formula_dax::{pivot, DataModel, FilterContext, GroupByColumn, PivotMeasure, Table, Value};
//!
//! let mut model = DataModel::new();
//! let mut fact = Table::new("Fact", vec!["Category", "Amount"]);
//! fact.push_row(vec![Value::from("A"), Value::from(10.0)]).unwrap();
//! fact.push_row(vec![Value::from("B"), Value::from(5.0)]).unwrap();
//! model.add_table(fact).unwrap();
//! model.add_measure("Total", "SUM(Fact[Amount])").unwrap();
//!
//! let measures = vec![PivotMeasure::new("Total", "[Total]").unwrap()];
//! let group_by = vec![GroupByColumn::new("Fact", "Category")];
//! let result = pivot(&model, "Fact", &group_by, &measures, &FilterContext::empty()).unwrap();
//! assert_eq!(result.rows.len(), 2);
//! ```

mod backend;
mod engine;
mod model;
mod parser;
mod pivot;
mod ident;
#[cfg(feature = "pivot-model")]
mod pivot_adapter;
#[cfg(feature = "pivot-model")]
mod pivot_config;
#[cfg(feature = "pivot-model")]
mod pivot_value;
mod value;

pub use crate::backend::{ColumnarTableBackend, InMemoryTableBackend, TableBackend};
pub use crate::engine::DaxEngine;
pub use crate::model::{Cardinality, CrossFilterDirection, DataModel, Measure, Relationship, Table};
pub use crate::pivot::{
    measures_from_value_fields, pivot, pivot_crosstab, pivot_crosstab_with_options, GroupByColumn,
    PivotCrosstabOptions, PivotMeasure, PivotResult, PivotResultGrid, ValueFieldAggregation,
    ValueFieldSpec,
};
#[cfg(feature = "pivot-model")]
pub use crate::pivot::measures_from_pivot_model_value_fields;
#[cfg(feature = "pivot-model")]
pub use crate::pivot_config::{pivot_crosstab_from_config, pivot_inputs_from_config, PivotInputs};
#[cfg(feature = "pivot-model")]
pub use crate::pivot::{PivotResultGridPivotValues, PivotResultPivotValues};
#[cfg(feature = "pivot-model")]
pub use crate::pivot_value::{dax_value_to_pivot_value, pivot_value_to_dax_value};
#[cfg(feature = "pivot-model")]
pub use crate::pivot_adapter::{build_data_model_pivot_plan, DataModelPivotPlan};
pub use crate::value::Value;

pub use crate::engine::{FilterContext, RowContext};
pub use crate::model::CalculatedColumn;
pub use crate::parser::{BinaryOp, Expr, UnaryOp};

pub use crate::engine::DaxError;
