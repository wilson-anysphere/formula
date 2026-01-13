# DAX engine (`formula-dax`)

This document describes the current implementation of the DAX engine in `crates/formula-dax`.
It is intended for contributors working on:

- Power Pivot / data model features
- DAX parsing and evaluation
- Relationship + filter propagation semantics
- Pivot/group-by execution paths

For overall PivotTable ownership boundaries (model schema vs worksheet pivots vs Data Model pivots)
and end-to-end refresh flows across crates, see
[ADR-0005: PivotTables ownership and data flow across crates](./adr/ADR-0005-pivot-tables-ownership-and-data-flow.md).

The engine is intentionally small and implements a **subset** of DAX. Anything not listed under
“Supported functions” or “Supported syntax” should be assumed **unsupported**.

> Reference implementation: see `crates/formula-dax/src/lib.rs` for a minimal working example.

---

## Core types: `DataModel`, `Table`, and backends

### `DataModel`

`DataModel` is the container for:

- tables (`HashMap<String, Table>`)
- relationships (`Vec<RelationshipInfo>`)
- measures (`HashMap<String, Measure>`)
- calculated columns (`Vec<CalculatedColumn>`)

Key APIs (see `crates/formula-dax/src/model.rs`):

- `DataModel::add_table(Table)`
- `DataModel::add_relationship(Relationship)`
- `DataModel::add_measure(name, expression)`
- `DataModel::add_calculated_column(table, name, expression)`
- `DataModel::evaluate_measure(name, &FilterContext)`
- `DataModel::insert_row(table, row_values)`

The evaluation entry point is typically `DaxEngine::evaluate(...)` (for ad-hoc expressions) or
`DataModel::evaluate_measure(...)` (for named measures).

`DaxEngine` also exposes a helper that is useful when an API takes a [`FilterContext`] but you want to
reuse DAX `CALCULATE`-style filter syntax:

- `DaxEngine::apply_calculate_filters(model, filter, &[filter_arg_strs...]) -> FilterContext`

### `Table` and `TableBackend`

`Table` is a thin wrapper around a `TableBackend` implementation (see `crates/formula-dax/src/backend.rs`).
The backend abstraction exists so the engine can:

- read cell values by row/column
- optionally use backend-specific accelerations (column stats, dictionary scans, group-by)

`TableBackend` supports optional accelerations via methods like:

- `stats_sum`, `stats_min`, `stats_max`, `stats_distinct_count`, `stats_has_blank`
- `dictionary_values`
- `filter_eq`, `filter_in`
- `distinct_values_filtered`
- `group_by_aggregations`

#### In-memory backend (`InMemoryTableBackend`)

Created via `Table::new(name, columns)` and backed by a `Vec<Vec<Value>>`:

- Mutable: supports `push_row`, adding calculated columns, and per-cell mutation (used by calculated columns).
- No accelerations: stats/dictionary/group-by methods return `None`, so evaluation falls back to scans.

This backend is used heavily in tests.

#### Columnar backend (`ColumnarTableBackend`)

Created via `Table::from_columnar(name, formula_columnar::ColumnarTable)`:

- Immutable: `push_row` and adding calculated columns are **not supported**.
- Provides accelerations used by common aggregations and the pivot engine:
  - column statistics (`stats_*`)
  - dictionary value enumeration
  - fast `filter_eq`/`filter_in` scans
  - grouped aggregations via `group_by_aggregations`

The columnar backend maps `formula_columnar::Value` into `formula_dax::Value`:

- `Null` → `Value::Blank`
- `Number` → `Value::Number`
- `Boolean` → `Value::Boolean`
- `String` → `Value::Text`
- `DateTime` / `Currency` / `Percentage` → `Value::Number` (stored as `f64`)

### `Value`

The engine’s scalar type is `formula_dax::Value` (`crates/formula-dax/src/value.rs`):

- `Blank`
- `Number(f64)`
- `Text(Arc<str>)`
- `Boolean(bool)`

There is no distinct datetime type at the DAX layer today; datetime-like values are represented as numbers.

#### Type coercions (current behavior)

The engine implements a small subset of DAX’s coercion rules. These matter most for arithmetic,
comparisons, and text concatenation:

- **Numeric coercion** (used by `+ - * /` and numeric comparisons):
  - `Number(n)` → `n`
  - `Boolean(true/false)` → `1.0` / `0.0`
  - `Blank` → `0.0`
  - `Text(...)` → type error

- **Text coercion** (used by the `&` operator):
  - `Text(s)` → `s`
  - `Number(n)` → `n.to_string()` (Rust formatting; not DAX format strings)
  - `Blank` → `""` (empty string)
  - `Boolean(true/false)` → `"TRUE"` / `"FALSE"`

- **Truthiness** (used by `IF`, `AND`, `OR`, and `&&`/`||`):
  - `Boolean(b)` → `b`
  - `Number(n)` → `n != 0.0`
  - `Blank` → `false`
  - `Text(...)` → type error

- **Comparison rules** (used by `= <> < <= > >=`):
  - Text comparisons are supported for `Text` vs `Text`, with `Blank` treated as the empty string.
  - Numeric comparisons apply numeric coercion (so `Blank` compares like `0`, booleans like `1/0`).
  - Comparing `Text` to a non-text value is a type error.

---

## Relationships

Relationships are registered via `DataModel::add_relationship(Relationship)` (see `model.rs`).

### Direction and indexing

`Relationship` is defined with:

- `from_table` / `from_column`: the fact-side table (often the "many" side) and its key/foreign-key column
- `to_table` / `to_column`: the lookup-side table (often the "one" side) and its key column

The relationship is *oriented*: by default (`CrossFilterDirection::Single`) filters propagate
`to_table → from_table`. This orientation is intended to remain meaningful for
`Cardinality::ManyToMany` relationships as well.

Internally, `DataModel` materializes two indices (`RelationshipInfo`):

- `to_index: HashMap<Value, usize>` mapping **to_table key → to_table row index** (unique-key relationships)
- `from_index: HashMap<Value, Vec<usize>>` mapping **from_table key → from_table row indices**

These indices are built eagerly when the relationship is added, and updated on `DataModel::insert_row(...)`.

### Cardinality

Currently supported cardinalities:

- `Cardinality::OneToMany`
- `Cardinality::OneToOne`

Many-to-many (planned semantics):

- `Cardinality::ManyToMany` is currently rejected by `DataModel::add_relationship` (returns
  `DaxError::UnsupportedCardinality`).
- Intended filter propagation is **distinct-key propagation** (conceptually like
  `TREATAS(VALUES(source[key]), target[key])`).
- `RELATED` is ambiguous when there is more than one match on the `to_table` side (error).
- `RELATEDTABLE` returns the set of matching rows (can be >1 for many-to-many).
- Pivot/group-by: the current pivot/group-by implementation does not expand a base row into multiple
  related rows when traversing a many-to-many relationship, so grouping across many-to-many is
  currently ambiguous/unsupported.

For `OneToOne` relationships, the engine enforces uniqueness on **both** sides when the relationship is
added (`DaxError::NonUniqueKey`), treating `BLANK` as a real key for uniqueness checks.

### Cross-filter direction

`CrossFilterDirection` controls relationship filter propagation:

- `Single`: filters propagate **one → many** only (`to_table` → `from_table`)
- `Both`: filters propagate **both directions** (`to_table` ↔ `from_table`)

Propagation is performed by `resolve_row_sets(...)` in `crates/formula-dax/src/engine.rs` and iterates
until a fixed point is reached.

### Active vs inactive and `USERELATIONSHIP`

Each relationship has `is_active: bool`.

- When `is_active == false`, the relationship is ignored by filter propagation.
- A relationship can be activated inside `CALCULATE(...)` via `USERELATIONSHIP(TableA[Col], TableB[Col])`.

Implementation detail (important for contributors):

- `FilterContext` tracks `active_relationship_overrides: HashSet<usize>` (relationship indices).
- When any override is present for a `(from_table, to_table)` pair, *only* those overridden relationships
  are considered active for that pair.

> Note: `CROSSFILTER(...)` can override direction or disable a relationship, but it does **not** activate
> inactive relationships. Use `USERELATIONSHIP` to activate an inactive relationship.

### Referential integrity

If `Relationship::enforce_referential_integrity == true`:

- On `add_relationship`, all non-blank foreign keys in the from-table must exist in `to_index`.
- On `insert_row`, the inserted row is rejected (and rolled back) if it violates referential integrity.

`BLANK` foreign keys are always allowed.

### Virtual blank row behavior (Tabular “unknown member”)

Tabular models behave as if the one-side table has an extra “unknown/blank” row when the many-side table
contains foreign keys that are:

- `BLANK`, or
- not present in the dimension key column (when referential integrity is not enforced)

`formula-dax` models this row **virtually**:

- The virtual row index is `dimension_table.row_count()` (one past the last physical row).
- Reading any column at that row yields `Value::Blank` (because it’s out-of-bounds in the backend).

Where this matters:

1. **Filter propagation (dimension → fact)**  
   When the virtual blank row is “allowed”, fact rows with unmatched keys stay visible even if the
   dimension is filtered.

2. **`VALUES(Dim[Key])` / `DISTINCT(Dim[Key])` and `DISTINCTCOUNT(Dim[Key])`**  
   These include `BLANK` when the virtual blank row exists and is allowed.

The virtual blank row is considered “allowed” when the filter context does **not** explicitly exclude it:

- Any `row_filters` on the dimension table disable it (row filters do not include the virtual row).
- Any column filter on the dimension table that does not include `BLANK` disables it.

See `blank_row_allowed(...)` and `virtual_blank_row_exists(...)` in `engine.rs`.

### Relationship resolution in DAX functions

Different DAX functions consult relationships in slightly different ways:

- `RELATED(Table[Column])`:
  - Requires a **direct** active relationship from the current row-context table to `Table`
    (`from_table == current_table`, `to_table == Table`, and `Relationship::is_active == true`).
  - Current limitation: it does **not** consult `FilterContext` relationship overrides
    (`USERELATIONSHIP` / `CROSSFILTER`).

- `RELATEDTABLE(Table)`:
  - Requires row context and returns rows from a many-side table related to the current row on the one
    side.
  - Supports multi-hop traversal when there is a **unique** active relationship path (in the reverse
    direction, `OneToMany` at each hop).
  - Consults `FilterContext` relationship overrides when deciding which relationships are active, and
    respects `CROSSFILTER(..., NONE)` disabling relationships.

---

## Measures vs calculated columns

### Measures

Measures are registered with:

```rust
model.add_measure("Total Sales", "SUM(Fact[Amount])")?;
```

They are evaluated in a **filter context** and do not store per-row results.

#### Implicit context transition for measures

When a measure is evaluated inside a **row context**, DAX performs an implicit context transition
(roughly `CALCULATE([Measure])`).

`formula-dax` implements this in `Expr::Measure` evaluation:

- If `row_ctx.current_table().is_some()` and
  `filter.suppress_implicit_measure_context_transition == false`, the engine calls
  `apply_context_transition(...)` before evaluating the measure.

This is why a measure reference like `[Total]` inside `SUMX(Fact, ...)` behaves differently than a
raw aggregation like `Fact[Amount]`.

### Calculated columns

Calculated columns are registered with:

```rust
model.add_calculated_column("Fact", "DoubleAmount", "Fact[Amount] * 2")?;
```

Current implementation details:

- Calculated columns are **materialized** into the table at definition time.
- They are evaluated with:
  - `FilterContext::default()` (effectively “no filter context”)
  - a row context pointing at each row of the table
- They are only supported for **in-memory tables** (`Table::new(...)`).
  - Columnar tables are immutable; use `add_calculated_column_definition(...)` if values are already stored.

On `DataModel::insert_row(...)`, calculated columns for that table are evaluated for the new row and stored
into the in-memory table.

---

## Filter context, row context, and `CALCULATE`

### `FilterContext`

`FilterContext` (`engine.rs`) currently contains:

- `column_filters: HashMap<(table, column), HashSet<Value>>`  
  Allowed values per column.
- `row_filters: HashMap<table, HashSet<usize>>`  
  Allowed row indices per table (usually produced by table expressions like `FILTER(...)`).
- `active_relationship_overrides: HashSet<usize>`  
  Activated relationships (via `USERELATIONSHIP`).
- `cross_filter_overrides: HashMap<usize, RelationshipOverride>`  
  Per-relationship overrides (via `CROSSFILTER`) that can change `CrossFilterDirection` or disable a
  relationship for the duration of evaluation.
- `suppress_implicit_measure_context_transition: bool`  
  Internal flag used to keep `CALCULATE` semantics correct.

Public helper APIs on `FilterContext` that are useful when calling the engine from Rust:

- `FilterContext::with_column_equals(table, column, value)`
- `FilterContext::with_column_in(table, column, values)`
- `FilterContext::set_column_equals(table, column, value)`
- `FilterContext::set_column_in(table, column, values)`
- `FilterContext::clear_column_filter_public(table, column)`

Filters combine with **AND** semantics:

- A row must satisfy all column filters on its table, and
- If a `row_filter` is present for a table, the row must be in that explicit set.

Filter propagation happens in `resolve_row_sets(...)`:

1. Apply explicit row filters and column filters to each table to get an initial “allowed rows” bitmap.
2. Repeatedly propagate filters across relationships until no table changes:
   - Always propagate `to_table (one) → from_table (many)`
   - Additionally propagate `from_table → to_table` for bidirectional relationships.

### `RowContext`

`RowContext` is a stack of `(table, row_index)` pairs.

It is primarily created by iterators (`SUMX`, `FILTER`, …) by pushing a row before evaluating an expression.

`RowContext` can contain **multiple entries for the same table** when iterators nest (e.g. nested `FILTER`
or `SUMX` over the same table). The engine supports this via:

- `EARLIER(Table[Column], [level])` to reference an outer row context for the same table
- `EARLIEST(Table[Column])` to reference the outermost row context for the table

### Context transition

Context transition is implemented by `apply_context_transition(...)`:

- For each table that has a “current row” in `RowContext`,
  the engine adds (or intersects) **equality column filters** for *every column in that row*.

This is used by:

- implicit measure context transition (measure references in row context)
- `CALCULATE` / `CALCULATETABLE` (explicit context transition)

### `CALCULATE(expr, filters...)`

`CALCULATE` is implemented by:

1. Starting from the current `FilterContext`
2. Applying context transition using the current `RowContext`
3. Evaluating filter arguments (order-independent) and applying their effects
4. Evaluating `expr` under the modified filter context

To preserve DAX behavior, the engine sets
`suppress_implicit_measure_context_transition = true` for the inner evaluation: measure references inside
`CALCULATE(...)` should not re-apply context transition and accidentally undo filter modifiers like `ALL(...)`.

#### Supported filter argument shapes

The engine supports the following filter argument forms (see `apply_calculate_filter_args(...)`):

1. `USERELATIONSHIP(TableA[Col], TableB[Col])`  
   Activates a relationship for this calculation.

2. `CROSSFILTER(TableA[Col], TableB[Col], direction)`  
   Overrides relationship filtering for the duration of the evaluation.

   - `direction` is a bare identifier (parsed as `Expr::TableName`), one of:
     - `BOTH`
     - `ONEWAY` (or `SINGLE`)
     - `NONE` (disables the relationship)

3. `ALL(Table)` / `ALL(Table[Column])` and `REMOVEFILTERS(Table)` / `REMOVEFILTERS(Table[Column])`  
   Clears filters on an entire table, or a specific column. (`REMOVEFILTERS` is treated as an alias for the
   `ALL` filter-modifier semantics.)

4. `KEEPFILTERS(innerFilterArg)`  
   Wraps a normal filter argument but changes its semantics from “replace filters” to “intersect filters”.
   Implementation note: `KEEPFILTERS` is supported only inside `CALCULATE` / `CALCULATETABLE`. It affects
   whether the engine clears existing table/column filters before applying the new filter.

5. Column comparisons: `Table[Column] <op> <scalar>` where `<op>` is:
   - `=` (direct value filter)
   - `<>`, `<`, `<=`, `>`, `>=` (implemented by scanning rows to compute the set of allowed values)

   Notes:
   - The RHS is evaluated as a scalar expression.
   - Non-equality comparisons currently build a *value set* by scanning rows under the current filters
     (except the column being filtered).

6. Value-set filters: `VALUES(Table[Column])` or `DISTINCT(Table[Column])`  
   Clears the column filter and replaces it with the set of distinct values visible under the current
   (post-transition) filter context.

7. `TREATAS(VALUES(SourceTable[Col]), TargetTable[Col])` (or `TREATAS(DISTINCT(...), Target...)`)  
   Applies the set of values from one column as a filter on another column.

   Current limitations:
   - The first argument must be `VALUES(column)` or `DISTINCT(column)`
   - The second argument must be a target column reference

8. Table expressions (row filters): any supported table expression (including a bare `TableName`)  
   The table expression is evaluated, and its resulting row set becomes an explicit `row_filter` for that
   table (intersected with any existing row filter).

   Examples:
   - `FILTER(Fact, Fact[Amount] > 0)`
   - `ALLEXCEPT(Dim, Dim[Category])`
   - `CALCULATETABLE(...)`

#### Unsupported filter argument shapes

Notable unsupported patterns:

- Boolean filter expressions that are not a column comparison (e.g. `Fact[Amount] > 0 && ...`)
- `IN` syntax (`Table[Col] IN {...}`) (not parsed)
- Many other filter modifiers (`ALLSELECTED`, `FILTERS`, `ISCROSSFILTERED`, etc.)

---

## Pivot API

The pivot engine produces grouped results (group keys + measures) suitable for pivot tables.

Entry points (see `crates/formula-dax/src/pivot.rs`):

- `formula_dax::pivot(...)` — returns a grouped table (`PivotResult`)
- `formula_dax::pivot_crosstab(...)` / `pivot_crosstab_with_options(...)` — returns a crosstab-shaped 2D
  grid (`PivotResultGrid`)

### `pivot(...)`

```rust
pub fn pivot(
    model: &DataModel,
    base_table: &str,
    group_by: &[GroupByColumn],
    measures: &[PivotMeasure],
    filter: &FilterContext,
) -> DaxResult<PivotResult>
```

### Parameters and grouping rules

- `base_table`: the table that is scanned/grouped (typically a fact table).
- `group_by`: ordered list of grouping columns.
  - Some execution paths require every `GroupByColumn.table == base_table`.
  - When `GroupByColumn.table != base_table`, the pivot engine will try to resolve a **unique active
    relationship path** from `base_table` to the group-by table (a `ManyToOne` chain) and compute the key
    via repeated relationship lookups (similar to `RELATED`, but possibly multi-hop).
    - If no path exists, or there are multiple active paths, pivot returns an evaluation error.
    - Current limitation: the path resolution uses relationship `is_active` flags only; it does not
      consider `USERELATIONSHIP`/`CROSSFILTER` overrides stored in the `FilterContext`.
- `measures`: list of named expressions to evaluate per group.
  - A `PivotMeasure` is *not* required to correspond to a named model measure; `expression` is parsed as DAX.
- `filter`: the initial filter context applied to the pivot query.

`PivotResult` contains:

- `columns`: group-by columns formatted as `"Table[Column]"`, followed by `PivotMeasure.name`
- `rows`: one row per group, sorted lexicographically by group keys (with `BLANK` sorting first)

### Performance paths

`pivot(...)` chooses the fastest applicable strategy:

1. **Columnar group-by + planned measures** (`pivot_columnar_group_by`)  
   Fast path when:
   - all `group_by` columns are on `base_table`, and
   - the backend supports `group_by_aggregations`, and
   - every measure can be “planned” into a small set of aggregations + arithmetic

   Planned expressions support:
   - `SUM`, `AVERAGE`, `MIN`, `MAX`, `DISTINCTCOUNT` over `base_table` columns
   - `COUNTROWS(base_table)`
   - simple arithmetic (`+ - * /`), unary `-`
   - `COALESCE`, `DIVIDE`
   - references to named measures that expand to the above

2. **Columnar groups + per-group measure evaluation** (`pivot_columnar_groups_with_measure_eval`)  
   Fast path when:
   - `group_by` is non-empty and only uses `base_table` columns
   - the backend can produce distinct group keys quickly

   Measures are then evaluated via `DaxEngine` per group. This is still *O(groups × measure_cost)*.

3. **Columnar star-schema group-by + planned measures** (`pivot_columnar_star_schema_group_by`)  
   Fast path when:
   - the base table is columnar and supports `group_by_aggregations`
   - `group_by` may include:
     - base table columns, and
     - one-hop related dimension columns (`base_table (many) -> dim_table (one)`; multi-hop is not supported)
   - measures are plannable, and their required aggregations are composable for rollup:
     - `AVERAGE` is currently not supported by this fast path
     - `DISTINCTCOUNT` is only allowed when it is constant within the final group (i.e. the counted column is
       itself part of the user-specified base-table group keys)

   Internally this groups by base-table columns and foreign keys, then “rolls up” those groups into the
   requested dimension attribute keys using relationship lookups.

4. **Row-scan planned group-by** (`pivot_planned_row_group_by`)  
   Works for non-columnar backends (and supports grouping by related dimension columns via a unique active
   relationship path), but still requires that measures are plannable (same restrictions as path 1).

5. **Fallback row-scan** (`pivot_row_scan`)  
   Always works, but is the slowest: it scans base rows to enumerate groups, then evaluates each measure
   via `DaxEngine` per group.

### `pivot_crosstab(...)`

`pivot_crosstab` is a small helper that calls `pivot` with `row_fields + column_fields`, then reshapes the
grouped result into a 2D grid.

```rust
pub fn pivot_crosstab(
    model: &DataModel,
    base_table: &str,
    row_fields: &[GroupByColumn],
    column_fields: &[GroupByColumn],
    measures: &[PivotMeasure],
    filter: &FilterContext,
) -> DaxResult<PivotResultGrid>
```

- `PivotResultGrid.data[0]` is a header row.
- Each subsequent row contains:
  - the row-axis key values, followed by
  - one value per `(column_key, measure)` combination.
- Missing (row_key, col_key) combinations are filled with `BLANK`.

Header formatting can be customized via `pivot_crosstab_with_options(..., &PivotCrosstabOptions { ... })`.

---

## Supported syntax

Parsing is implemented in `crates/formula-dax/src/parser.rs`.

Supported expression forms:

- Numbers: `1`, `1.5`, `.25` (no exponent syntax)
- Strings: `"hello"` with `""` as an escape for `"`
- Identifiers: `SUM`, `Fact`, `MyTable` (ASCII letters/`_`/`.`; may include digits after the first char)
- Quoted identifiers: `'Table Name'` (escape `'` as `''`)
- Bracket identifiers:
  - measures: `[Total]`
  - column names: `Fact[Amount]` (column part is bracketed)
- Function calls: `NAME(arg1, arg2, ...)` (comma-separated arguments)
- Unary minus: `-expr`
- Binary operators:
  - arithmetic: `+ - * /`
  - text concatenation: `&` (single ampersand)
  - comparisons: `= <> < <= > >=`
  - boolean: `&&` and `||`
- Parentheses for grouping

Unsupported (not parsed):

- `VAR` / `RETURN`
- `{ ... }` table constructors
- `IN` operator
- `;` argument separators (locale variants)

---

## Supported DAX functions

If a function is not listed here, it is currently unsupported and will evaluate to a `DaxError::Eval`
(`"unsupported function ..."` or `"unsupported table function ..."`).

> Note: `TRUE`, `FALSE`, and `BLANK` must be called as functions (`TRUE()`) because the parser does not
> have keyword literals for booleans/null.

### Scalar functions

- `TRUE()`, `FALSE()`, `BLANK()`
- `ISBLANK(x)`
- `IF(condition, then, [else])`
- `SWITCH(expr, value1, result1, ..., [else])`
- `DIVIDE(numerator, denominator, [alternateResult])`
- `COALESCE(arg1, arg2, ...)`
- `NOT(x)`
- `AND(x, y)`
- `OR(x, y)`
- `SUM(Table[Column])`
- `AVERAGE(Table[Column])`
- `MIN(Table[Column])`
- `MAX(Table[Column])`
- `COUNT(Table[Column])` (counts numeric values)
- `COUNTA(Table[Column])` (counts non-blank values)
- `COUNTBLANK(Table[Column])`
- `DISTINCTCOUNT(Table[Column])`
- `DISTINCTCOUNTNOBLANK(Table[Column])`
- `COUNTROWS(tableExpr)`
- `SUMX(tableExpr, valueExpr)`
- `AVERAGEX(tableExpr, valueExpr)`
- `MAXX(tableExpr, valueExpr)`
- `MINX(tableExpr, valueExpr)`
- `COUNTX(tableExpr, valueExpr)`
- `HASONEVALUE(Table[Column])`
- `SELECTEDVALUE(Table[Column], [alternateResult])`
- `LOOKUPVALUE(ResultTable[ResultCol], SearchTable[SearchCol1], SearchValue1, ..., [alternateResult])`  
  (current MVP restriction: all search columns must be in the same table as the result column)
- `CALCULATE(expr, filter1, filter2, ...)`
- `RELATED(Table[Column])` (requires row context)
- `EARLIER(Table[Column], [level])` (requires nested row context)
- `EARLIEST(Table[Column])` (requires row context)

### Table functions

- `FILTER(tableExpr, predicateExpr)`
- `ALL(Table)` and `ALL(Table[Column])`
- `VALUES(Table[Column])`, `VALUES(tableExpr)`
- `DISTINCT(Table[Column])`, `DISTINCT(tableExpr)` (implemented in terms of `VALUES`)
- `ALLEXCEPT(Table, Table[Col1], Table[Col2], ...)`
- `CALCULATETABLE(tableExpr, filter1, filter2, ...)`
- `SUMMARIZE(tableExpr, Table[GroupCol1], Table[GroupCol2], ...)`  
  (limited: currently only grouping columns are supported; group columns may be on the base table or on
  related tables reachable via a unique active relationship path; it returns a row set of the base table)
- `SUMMARIZECOLUMNS(Table[GroupCol1], Table[GroupCol2], ...)`  
  (limited: currently only grouping columns are supported; the engine picks a base table that can reach
  all grouped tables via active relationships; it returns a row set of that base table)
- `RELATEDTABLE(Table)` (requires row context)

### Filter modifiers inside `CALCULATE`

- `USERELATIONSHIP(TableA[Col], TableB[Col])`
- `CROSSFILTER(TableA[Col], TableB[Col], BOTH|ONEWAY|SINGLE|NONE)`
- `TREATAS(VALUES(Source[Col])|DISTINCT(Source[Col]), Target[Col])` (limited)
- `KEEPFILTERS(innerFilterArg)` (supported only as a wrapper inside `CALCULATE` / `CALCULATETABLE`)
- `REMOVEFILTERS(Table|Table[Column])` (alias for `ALL`-style clearing inside `CALCULATE`)

### Notable unsupported functions (non-exhaustive)

The real DAX surface area is large; `formula-dax` currently only implements the functions listed above.
Some common functions that are **not implemented** (and will error if called) include:

- Table shaping: `ADDCOLUMNS`, `SELECTCOLUMNS`, `GENERATE`, `UNION`, `INTERSECT`, `EXCEPT`, `CROSSJOIN`, `TOPN`
- Filter inspection: `ISFILTERED`, `HASONEFILTER`, `ISCROSSFILTERED`, `ALLSELECTED`
- Row context helpers: `RANKX`
- Time intelligence: `CALENDAR`, `DATEADD`, `SAMEPERIODLASTYEAR`, etc.

---

## Examples

### Relationship + `RELATED`

```rust
use formula_dax::{Cardinality, CrossFilterDirection, DataModel, DaxEngine, FilterContext, Relationship, RowContext, Table, Value};

let mut model = DataModel::new();

let mut dim = Table::new("DimCategory", vec!["CategoryId", "Name"]);
dim.push_row(vec![Value::from("A"), Value::from("Alpha")])?;
dim.push_row(vec![Value::from("B"), Value::from("Beta")])?;
model.add_table(dim)?;

let mut fact = Table::new("Fact", vec!["CategoryId", "Amount"]);
fact.push_row(vec![Value::from("A"), Value::from(10.0)])?;
fact.push_row(vec![Value::from("B"), Value::from(5.0)])?;
model.add_table(fact)?;

model.add_relationship(Relationship {
    name: "Fact->DimCategory".to_string(),
    from_table: "Fact".to_string(),
    from_column: "CategoryId".to_string(),
    to_table: "DimCategory".to_string(),
    to_column: "CategoryId".to_string(),
    cardinality: Cardinality::OneToMany,
    cross_filter_direction: CrossFilterDirection::Single,
    is_active: true,
    enforce_referential_integrity: false,
})?;

let engine = DaxEngine::new();
let mut row_ctx = RowContext::default();
row_ctx.push("Fact", 0);
let v = engine.evaluate(&model, "RELATED(DimCategory[Name])", &FilterContext::empty(), &row_ctx)?;
assert_eq!(v, Value::from("Alpha"));
```

### Pivot with a measure

```rust
use formula_dax::{pivot, DataModel, FilterContext, GroupByColumn, PivotMeasure, Table, Value};

let mut model = DataModel::new();
let mut fact = Table::new("Fact", vec!["Category", "Amount"]);
fact.push_row(vec![Value::from("A"), Value::from(10.0)])?;
fact.push_row(vec![Value::from("B"), Value::from(5.0)])?;
model.add_table(fact)?;

model.add_measure("Total", "SUM(Fact[Amount])")?;

let group_by = vec![GroupByColumn::new("Fact", "Category")];
let measures = vec![PivotMeasure::new("Total", "[Total]")?];

let result = pivot(&model, "Fact", &group_by, &measures, &FilterContext::empty())?;
assert_eq!(result.columns, vec!["Fact[Category]".to_string(), "Total".to_string()]);
```

---

## Current limitations and “sharp edges”

This is not an exhaustive list, but the most common contributor-facing constraints:

- **Relationships**
  - `OneToMany` and `OneToOne` are supported; `ManyToMany` is not.
  - Only single-column relationships are supported.
  - Some APIs consult only `Relationship::is_active` and do not apply `FilterContext` relationship
    overrides (notably `RELATED` and pivot group-by path discovery).
- **DAX language coverage**
  - No variables (`VAR`/`RETURN`), no `IN`, no table constructors.
  - Most scalar/table functions are unimplemented (anything not listed above).
- **Types**
  - Only `Blank`, `Number(f64)`, `Boolean`, and `Text` exist at the DAX layer.
- **Calculated columns**
  - `add_calculated_column(...)` only works for in-memory tables.
  - Columnar tables can only register definitions via `add_calculated_column_definition(...)`.
- **Table semantics**
  - `SUMMARIZE(...)` currently returns a row set of the base table, not a materialized grouped table.
  - Table expressions are represented as `(table_name, row_indices)` rather than as independent rowsets.
- **`/` operator**
  - The `/` operator performs raw `f64` division; use `DIVIDE(...)` for DAX-like blank/alternate behavior.
