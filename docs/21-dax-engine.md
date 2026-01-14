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

### Case-insensitive identifiers (`normalize_ident`)

Like Tabular / Power Pivot, `formula-dax` resolves identifiers **case-insensitively**:

- table names (`Orders`)
- column names (`Orders[Amount]`)
- measures (`[Total Sales]`)
- `VAR` identifiers (`VAR Region = ... RETURN REGION`)
- relationship endpoints (e.g. `USERELATIONSHIP(orders[customerid], CUSTOMERS[CUSTOMERID])`)

Internally, schema and filter lookups normalize identifiers via `normalize_ident` (see
`crates/formula-dax/src/model.rs`):

- Identifiers are trimmed (`s.trim()`) before normalization (leading/trailing whitespace is ignored).
- ASCII identifiers are uppercased (`orders` → `ORDERS`)
- non-ASCII identifiers use Unicode-aware uppercasing (e.g. `ß` → `SS`) to approximate Excel/Tabular
  matching

As a consequence, duplicates that collide after normalization are rejected up-front (e.g. two columns
named `Col` and `col`, two measures named `Total` and `[TOTAL]`, or two tables named `Straße` and
`STRASSE`).

### `Table` and `TableBackend`

`Table` is a thin wrapper around a `TableBackend` implementation (see `crates/formula-dax/src/backend.rs`).
The backend abstraction exists so the engine can:

- read cell values by row/column
- optionally use backend-specific accelerations (column stats, dictionary scans, group-by)

`TableBackend` supports optional accelerations via methods like:

- `stats_sum`, `stats_min`, `stats_max`, `stats_non_blank_count`, `stats_distinct_count`, `stats_has_blank`
- `dictionary_values`
- `filter_eq`, `filter_in`
- `distinct_values_filtered`
- `group_by_aggregations`
- `columnar_table` / `hash_join` (when backed by `formula-columnar`)

#### In-memory backend (`InMemoryTableBackend`)

Created via `Table::new(name, columns)` and backed by a `Vec<Vec<Value>>`:

- Mutable: supports `push_row`, adding calculated columns, and per-cell mutation (used by calculated columns).
- No accelerations: stats/dictionary/group-by methods return `None`, so evaluation falls back to scans.

This backend is used heavily in tests.

#### Columnar backend (`ColumnarTableBackend`)

Created via `Table::from_columnar(name, formula_columnar::ColumnarTable)`:

- Immutable: `push_row` is not supported.
- Calculated columns **are supported**: `DataModel::add_calculated_column` computes the expression
  eagerly and materializes the results by appending an encoded column to the underlying
  `formula_columnar::ColumnarTable` (copy-on-write; may clone when shared).
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
- `DateTime` → `Value::Number` (stored as `f64`)
- `Currency(scale)` / `Percentage(scale)` → `Value::Number` (stored as `f64`)  
  The raw integer value is divided by `10^scale` so currency/percentage scaling is preserved.

### `Value`

The engine’s scalar type is `formula_dax::Value` (`crates/formula-dax/src/value.rs`):

- `Blank`
- `Number(ordered_float::OrderedFloat<f64>)`
- `Text(Arc<str>)`
- `Boolean(bool)`

There is no distinct datetime type at the DAX layer today; datetime-like values are represented as numbers.
`ordered_float::OrderedFloat` is used so numeric values can participate in `HashMap` keys and stable
ordering (needed for dictionary scans, relationship indices, and group-by).

#### Type coercions (current behavior)

The engine implements a small subset of DAX’s coercion rules. These matter most for arithmetic,
comparisons, and text concatenation:

- **Numeric coercion** (used by `+ - * /` and numeric comparisons):
  - `Number(n)` → `n.0`
  - `Boolean(true/false)` → `1.0` / `0.0`
  - `Blank` → `0.0`
  - `Text(...)` → type error

- **Text coercion** (used by the `&` operator):
  - `Text(s)` → `s`
  - `Number(n)` → `n.0.to_string()` (Rust formatting; not DAX format strings)
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

- `from_table` / `from_column`: the fact-side table (often the "many" side for 1:*) and its key/foreign-key column
- `to_table` / `to_column`: the lookup-side table (often the "one" side for 1:*) and its key column

The relationship is *oriented*: by default (`CrossFilterDirection::Single`) filters propagate
`to_table → from_table`. This orientation is meaningful for `Cardinality::ManyToMany` relationships
as well (it defines default propagation and how `RELATED` / `RELATEDTABLE` navigate).

Internally, `DataModel` materializes relationship metadata (`RelationshipInfo`), including:

- `from_idx: usize` — resolved column index of `from_column` in `from_table`
- `to_idx: usize` — resolved column index of `to_column` in `to_table`
- `to_index: ToIndex` — lookup structure for the `to_table[to_column]` key:
  - `ToIndex::RowSets { map: HashMap<Value, RowSet>, has_duplicates }`  
    Full mapping from **to_table key → matching to_table row(s)**. `RowSet` is an internal compact
    representation:
    - `RowSet::One(row)` for the common unique-key case (no allocation)
    - `RowSet::Many(Vec<row>)` when a non-blank key matches multiple `to_table` rows (many-to-many)  
      Note: for many-to-many relationships, physical `BLANK` keys do not participate in relationship
      joins and are skipped (they are not stored in the index), to avoid materializing a potentially
      huge, never-used row list.
  - `ToIndex::KeySet { keys: HashSet<Value>, has_duplicates }`  
    Scalable representation for **columnar many-to-many** relationships where the `to_table` side may
    contain a very large number of duplicate keys. Instead of storing `Vec<usize>` row lists per key,
    the engine stores only the **distinct key set** and relies on backend primitives like
    `filter_eq` / `filter_in` to retrieve matching row indices on demand. `BLANK` keys are excluded
    (they do not participate in relationship joins).
    - `KeySet` is currently used when `cardinality == ManyToMany` and the `to_table` is backed by the
      columnar backend (`TableStorage::Columnar`).
- `from_index: Option<HashMap<Value, Vec<usize>>>` mapping **from_table key → from_table row indices**  
  Materialized only for in-memory fact tables. For columnar fact tables it stays `None` and the engine
  relies on backend primitives like `filter_eq` / `filter_in` instead.
- `unmatched_fact_rows: Option<UnmatchedFactRows>` containing **from_table rows whose foreign key is
  `BLANK` or does not exist in `to_index`**. This cache is used to implement Tabular's virtual
  blank/unknown member semantics efficiently, and is especially important when `from_index` is not
  materialized (columnar fact tables). `UnmatchedFactRows` is stored either as:
  - `Sparse(Vec<usize>)` for small sets, or
  - `Dense { bits: Vec<u64>, len, count }` for large sets (bitset representation).

These indices are built eagerly when the relationship is added. For in-memory tables,
`DataModel::insert_row(...)` updates `to_index` and `from_index` incrementally (columnar tables are
immutable and do not support row insertion).

### Relationship join column type validation

When a relationship is added, `DataModel` validates that the join columns have compatible types
(`DaxError::RelationshipJoinColumnTypeMismatch`).

- Columnar tables use the declared column types; `Number`/`DateTime`/`Currency`/`Percentage` are all
  considered “numeric-like” and compatible with each other.
- In-memory tables infer the join type by scanning up to 1,000 rows for the first non-blank value.
  (If a type cannot be inferred, validation is skipped to avoid false positives for empty tables.)

### Cardinality

Supported cardinalities:

- `Cardinality::OneToMany`
- `Cardinality::OneToOne`
- `Cardinality::ManyToMany`

For `OneToMany` relationships, the engine enforces uniqueness on the `to_table` key when the
relationship is added (`DaxError::NonUniqueKey`).

For `OneToOne` relationships, the engine enforces uniqueness on **both** sides when the relationship is
added (`DaxError::NonUniqueKey`), treating `BLANK` as a real key for uniqueness checks.

Many-to-many semantics:

- Filter propagation uses **distinct-key propagation** (conceptually like
  `TREATAS(VALUES(source[key]), target[key])`).
- `RELATED` is ambiguous when there is more than one match on the `to_table` side (error).
- `RELATEDTABLE` returns the set of matching rows (can be >1 for many-to-many).
- Pivot/group-by note: grouping by columns across a many-to-many relationship **expands** a base row
  into multiple related rows. When a key matches multiple rows on the `to_table` side, the engine
  treats the group key as a set of possible related **rows/tuples**. Columns that come from the same
  related table (i.e. share the same relationship path) stay correlated per related row
  (e.g. `(Category, Color)` pairs come from the same related row), while the engine takes a cartesian
  product across **independent** grouping sources (base-table columns and each distinct relationship
  path). This can duplicate measure contributions (a single fact row can contribute to multiple
  groups), so summing group totals can “double count” compared to an ungrouped total. This differs
  from scalar navigation (`RELATED`), which errors on ambiguity.

### Cross-filter direction

`CrossFilterDirection` controls relationship filter propagation:

- `Single`: filters propagate **one → many** only (`to_table` → `from_table`)
- `Both`: filters propagate **both directions** (`to_table` ↔ `from_table`)

Propagation is performed by `resolve_row_sets(...)` in `crates/formula-dax/src/engine.rs` and iterates
until a fixed point is reached.

#### Debugging relationship propagation

Set `FORMULA_DAX_RELATIONSHIP_TRACE=1` to print a one-time summary of relationship propagation
(`resolve_row_sets`) for the first evaluation in the process. This includes table row counts before/after
filtering and the number of propagation iterations/updates.

### Active vs inactive and `USERELATIONSHIP`

Each relationship has `is_active: bool`.

- When `is_active == false`, the relationship is ignored by filter propagation.
- A relationship can be activated inside `CALCULATE(...)` via `USERELATIONSHIP(TableA[Col], TableB[Col])`.

Implementation detail (important for contributors):

- `FilterContext` tracks `active_relationship_overrides: HashSet<usize>` (relationship indices).
- When any override is present for a `(from_table, to_table)` pair, *only* those overridden relationships
  are considered active for that pair.

> Note: `CROSSFILTER(...)` can override direction (including reverse one-way directions) or disable a relationship, but it does **not** activate
> inactive relationships. Use `USERELATIONSHIP` to activate an inactive relationship.

### Referential integrity

If `Relationship::enforce_referential_integrity == true`:

- On `add_relationship`, all non-blank foreign keys in the from-table must exist in `to_index`.
- On `insert_row`, the inserted row is rejected (and rolled back) if it violates referential integrity.

`BLANK` foreign keys are always allowed.

### Virtual blank row behavior (Tabular “unknown member”)

Tabular models behave as if the table on the `to_table` side of a relationship has an extra
“unknown/blank” row when the table on the `from_table` side contains key values that are:

- `BLANK`, or
- not present in the dimension key column (when referential integrity is not enforced)

The engine treats “blank member existence” as a **dynamic, filter-context-dependent** property:
the virtual blank member exists for a `to_table` only when there are **currently visible**
unmatched/blank foreign keys on the `from_table` side under the active relationship set (including
`USERELATIONSHIP` / `CROSSFILTER` overrides). This means that filtering the fact table can make the
blank member appear or disappear even if the model contains unmatched keys overall.

Important nuance: in `formula-dax`, fact-side `BLANK` foreign keys always belong to this
relationship-generated blank member, even if the dimension table contains a *physical* row whose key
is `BLANK`. In other words, `BLANK` is treated as an **unmatchable** relationship join key during
filter propagation and row-context navigation (`RELATED` / `RELATEDTABLE`).

`formula-dax` models this row **virtually**:

- The virtual row index is `dimension_table.row_count()` (one past the last physical row).
- Reading any column at that row yields `Value::Blank` (because it’s out-of-bounds in the backend).

Where this matters:

1. **Filter propagation (`to_table → from_table`)**  
    When the virtual blank row is “allowed”, rows with unmatched keys stay visible even if the
    `to_table` side is filtered.

2. **`VALUES(Dim[Key])` / `DISTINCT(Dim[Key])` and `DISTINCTCOUNT(Dim[Key])`**  
    These include `BLANK` when the virtual blank row exists and is allowed.

3. **Filtering to the blank member**  
   Filtering a dimension attribute to `BLANK` (e.g. `Customers[Region] = BLANK()`) selects the
   relationship-generated blank/unknown member when it is allowed, so measures evaluated under that
   filter include fact rows whose foreign key is `BLANK` or unmatched.

The virtual blank row is considered “allowed” when the filter context does **not** explicitly exclude it:

- Any `row_filters` on the dimension table disable it (row filters do not include the virtual row).
- Any column filter on the dimension table that does not include `BLANK` disables it.

Two additional nuances (important for contributor mental models):

- **Snowflake chains:** the virtual blank member can *cascade* across relationships. Conceptually, the
  blank member of a table belongs to the blank member of its lookup tables, so:
  `Sales (unmatched ProductId)` → `Products (blank member)` → `Categories (blank member)`.
- **Indirect BLANK exclusion:** filtering `BLANK` out of a lookup table can also make downstream blank
  members invisible. The engine computes this with `compute_blank_row_allowed_map(...)`, propagating
  BLANK exclusion from `to_table → from_table` along active relationships. If `CROSSFILTER` forces a
  relationship to propagate **only in the reverse direction** (`ONEWAY_LEFTFILTERSRIGHT` /
  `ONEWAY_RIGHTFILTERSLEFT` in the reverse orientation), filters on `to_table` do not restrict
  `from_table`, so BLANK exclusion does not cascade across that hop.

See `blank_row_allowed(...)`, `compute_blank_row_allowed_map(...)`, and `virtual_blank_row_exists(...)`
in `engine.rs`.

### Relationship resolution in DAX functions

Different DAX functions consult relationships in slightly different ways:

- `RELATED(Table[Column])`:
  - Requires row context.
  - Follows a **unique** active relationship path from the current row-context table to `Table` in the
    `ManyToOne` direction (following relationships in their defined `from_table → to_table` direction).
    If there are multiple active paths, the engine errors.
  - Consults `FilterContext` relationship overrides:
    - `USERELATIONSHIP` activation
    - `CROSSFILTER(..., NONE)` disabling relationships
  - Returns `BLANK` for blank keys or missing matches.
  - Errors if a hop is ambiguous (key matches multiple rows on the one-side).

- `RELATEDTABLE(Table)`:
  - Requires row context and returns rows from a `from_table`-side table related to the current row
    on the `to_table` side.
  - Supports multi-hop traversal when there is a **unique** active relationship path (in the reverse
    direction, `OneToMany` at each hop). If there are multiple active paths, the engine errors.
  - Consults `FilterContext` relationship overrides when deciding which relationships are active, and
    respects `CROSSFILTER(..., NONE)` disabling relationships.

---

## Measures vs calculated columns

### Measures

Measures are registered with:

```rust
model.add_measure("Total Sales", "SUM(Fact[Amount])")?;
```

Measure name normalization notes:

- Measure names are stored/looked up via `normalize_ident` (trimmed + case-insensitive).
- `DataModel::add_measure` / `DataModel::evaluate_measure` normalize names by stripping a single outer
  bracket pair and trimming whitespace (`[Total]` and `Total` refer to the same measure).

They are evaluated in a **filter context** and do not store per-row results.

#### Implicit context transition for measures

When a measure is evaluated inside a **row context**, DAX performs an implicit context transition
(roughly `CALCULATE([Measure])`).

`formula-dax` implements this in `Expr::Measure` evaluation:

- If `!row_ctx.is_empty()` (i.e. there is any row context frame, including virtual row context) and
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
- They are supported for both **in-memory** (`Table::new(...)`) and **columnar-backed** (`Table::from_columnar(...)`)
  tables.
  - For columnar tables, the computed values are encoded and appended as a new column to the underlying
    `formula_columnar::ColumnarTable` (copy-on-write; may clone when shared).
  - Columnar calculated columns currently require a single logical type across all **non-blank** rows:
    number, string, or boolean.
  - When loading persisted models where calculated column values are already stored, use
    `add_calculated_column_definition(...)` to register the metadata without re-evaluating the expression.

On `DataModel::insert_row(...)`, calculated columns for that table are evaluated for the new row and stored
into the in-memory table (note: `insert_row` is not supported for columnar tables).

---

## Filter context, row context, and `CALCULATE`

### `FilterContext`

`FilterContext` (`engine.rs`) currently contains:

- `column_filters: HashMap<(table, column), HashSet<Value>>`  
  Allowed values per column. Keys are normalized via `normalize_ident` (trimmed + case-insensitive).
- `row_filters: HashMap<table, RowFilter>`  
  Allowed physical rows per table (usually produced by table expressions like `FILTER(...)`). The
  internal `RowFilter` representation is:
  - `All` (all physical rows; used to avoid allocating huge sets for some filters, while still
    suppressing the relationship-generated blank member)
  - `Rows(HashSet<usize>)` (sparse row set)
  - `Mask(Arc<BitVec>)` (dense bitmap)
- `active_relationship_overrides: HashSet<usize>`  
  Activated relationships (via `USERELATIONSHIP`).
- `cross_filter_overrides: HashMap<usize, RelationshipOverride>`  
  Per-relationship overrides (via `CROSSFILTER`) that can:
  - disable a relationship (`NONE`)
  - force bidirectional filtering (`BOTH`)
  - force one-way filtering in either the relationship’s default direction (`ONEWAY`/`SINGLE`) or the
    reverse direction (`ONEWAY_LEFTFILTERSRIGHT` / `ONEWAY_RIGHTFILTERSLEFT`)
- `suppress_implicit_measure_context_transition: bool`  
  Internal flag used to keep `CALCULATE` semantics correct.
- `in_scope_columns: HashSet<(table, column)>`  
  Pivot-driven “scope” metadata used to implement `ISINSCOPE`. The pivot engine populates this with
  its axis columns; standalone DAX evaluation does not infer scope, so the set is empty by default.

Note: `column_filters` is currently a **set-based** representation (allowed values), not a predicate/range-based one.
This matters for PivotTables / timelines: expressing a date range like `[start, end]` requires either materializing a
potentially-large allowed set, or extending `FilterContext` to support `>=`/`<=` style predicates.

Public helper APIs on `FilterContext` that are useful when calling the engine from Rust:

- `FilterContext::with_column_equals(table, column, value)`
- `FilterContext::with_column_in(table, column, values)`
- `FilterContext::set_column_equals(table, column, value)`
- `FilterContext::set_column_in(table, column, values)`
- `FilterContext::clear_column_filter_public(table, column)`
- `FilterContext::clear_table_filters_public(table)`

Filters combine with **AND** semantics:

- A row must satisfy all column filters on its table, and
- If a `row_filter` is present for a table, the row must be allowed by that row filter (whether it is
  `All`, a sparse `Rows(...)` set, or a dense bitmap `Mask(...)`).

Filter propagation happens in `resolve_row_sets(...)`:

1. Apply explicit row filters and column filters to each table to get an initial “allowed rows” bitmap.
2. Repeatedly propagate filters across relationships until no table changes:
   - Always propagate `to_table (one) → from_table (many)`
   - Additionally propagate `from_table → to_table` for bidirectional relationships.

### `RowContext`

`RowContext` is a stack of row-context frames.

It is primarily created by iterators (`SUMX`, `FILTER`, …) by pushing a row before evaluating an expression.

The engine supports two kinds of row context frames:

- **Physical rows**: `(table_name, row_index, visible_cols?)`  
  This is the common case when iterating a physical table. `visible_cols` is used for single-column
  table expressions like `VALUES(Table[Column])` where DAX exposes only that column in row context.

- **Virtual rows**: `Vec<((table, column), value)>` bindings  
  Some table functions (notably `SUMMARIZE` / `SUMMARIZECOLUMNS`) return a *virtual table* of grouping
  keys. Iterating those tables pushes a virtual row context containing explicit bindings for the
  grouped columns.

When resolving a column reference in row context (`Table[Column]`), the engine checks for a matching
virtual binding first, then falls back to looking up the value from the most recent physical row for
that table.

`RowContext` can contain **multiple physical entries for the same table** when iterators nest (e.g.
nested `FILTER` or `SUMX` over the same table). The engine supports this via:

- `EARLIER(Table[Column], [level])` to reference an outer row context for the same table
- `EARLIEST(Table[Column])` to reference the outermost row context for the table

Note: `EARLIER`/`EARLIEST` only consult *physical* row contexts for that table; virtual row contexts
do not participate.

### Context transition

Context transition is implemented by `apply_context_transition(...)`:

- For **virtual row contexts** (from virtual table iteration), it adds/intersects equality filters
  for the explicitly bound `(table, column) = value` pairs.
- For **physical row contexts**, it adds/intersects equality filters for the “current row” of each
  physical table (the innermost row context wins when the same table appears multiple times). If the
  row context is restricted to `visible_cols`, it only applies filters for those columns.

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

   - `direction` is a bare identifier (parsed as `Expr::TableName`) or a string literal, one of:
       - `BOTH`
       - `ONEWAY` (or `SINGLE`)
       - `ONEWAY_LEFTFILTERSRIGHT`
       - `ONEWAY_RIGHTFILTERSLEFT`
       - `NONE` (disables the relationship)
   - If multiple `CROSSFILTER(...)` modifiers in the same `CALCULATE(...)` target the **same**
     relationship with different directions, the engine errors (ambiguous).

3. `ALL(Table)` / `ALL(Table[Column])`, `ALLNOBLANKROW(Table)` / `ALLNOBLANKROW(Table[Column])`, and
   `REMOVEFILTERS(Table)` / `REMOVEFILTERS(Table[Column])`  
   Clears filters on an entire table, or a specific column. (`REMOVEFILTERS` is treated as an alias for the
   `ALL` filter-modifier semantics.)

4. `KEEPFILTERS(innerFilterArg)`  
   Wraps a normal filter argument but changes its semantics from “replace filters” to “intersect filters”.
   Implementation note: `KEEPFILTERS` is supported only inside `CALCULATE` / `CALCULATETABLE`. It affects
   whether the engine clears existing table/column filters before applying the new filter.

5. Column comparisons / membership:
   - Column comparisons: `Table[Column] <op> <rhs>` where `<op>` is:
     - `=` (direct value filter)
     - `<>`, `<`, `<=`, `>`, `>=` (implemented by scanning rows to compute the set of allowed values)
   - Membership (`IN`):
     - scalar: `Table[Column] IN tableExpr`, e.g. `Fact[Category] IN { \"A\", \"B\" }` or `Fact[Category] IN VALUES(Dim[Category])`
     - row constructor: `(T[Col1], T[Col2], ...) IN tableExpr`
  
   In `CALCULATE`, row-constructor `IN` filters are applied as a **row filter** (not decomposed into
   independent per-column value filters) to preserve correlation across columns, e.g.:
  
   - `CALCULATE(COUNTROWS(Orders), (Orders[OrderId], Orders[CustomerId]) IN {(100,1), (102,2)})`

   Notes:
   - For `=`/`<>`/`<`/`<=`/`>`/`>=`, the RHS is evaluated as a scalar expression.
   - For scalar `IN` (`T[Col] IN tableExpr`), the RHS must evaluate to a one-column table expression.
   - For row-constructor `IN` (`(T[Col1], T[Col2], ...) IN tableExpr`), the RHS must evaluate to a table
     with the same number of columns. Matching is positional (tuple element 1 vs table column 1, etc.)
     using the normal `=` comparison semantics. In `CALCULATE` filter arguments, the row constructor
     must contain only column references and must reference exactly one table.
   - Non-equality comparisons currently build a *value set* by scanning rows under the current filters
     (except the column being filtered).

6. Boolean filter expressions (row filters): expressions like  
   `Fact[Amount] > 0 && Fact[Amount] < 10` or `NOT(Fact[IsActive])`  
   These are evaluated by scanning candidate rows and building an explicit `row_filter`.

   Current limitations:
   - The expression must reference columns from **exactly one table**.
   - Only `&&`/`||` (binary operators) and `NOT(...)`/`AND(...)`/`OR(...)` (function forms) are recognized as
     boolean filter expressions at the top level.

7. Value-set filters: `VALUES(Table[Column])` or `DISTINCT(Table[Column])`  
   Clears the column filter and replaces it with the set of distinct values visible under the current
   (post-transition) filter context.

8. `TREATAS(VALUES(SourceTable[Col]), TargetTable[Col])` (or `TREATAS(DISTINCT(...), Target...)`)  
   Applies the set of values from one column as a filter on another column.

   Current limitations:
   - The first argument must be `VALUES(column)` or `DISTINCT(column)`
   - The second argument must be a target column reference

9. Table expressions (row filters): any supported **physical** table expression (including a bare
     `TableName`)  
     The table expression is evaluated and must produce a **physical** table result (`TableResult::Physical`,
     `TableResult::PhysicalAll`, or `TableResult::PhysicalMask`) — i.e. a base table plus a set of visible
     rows. How it is interpreted depends on whether the table result is **projected**:
 
     - If `visible_cols == None` (a full-row physical table), its resulting row set becomes an explicit
       `row_filter` for that table (intersected with any existing row filter).
     - If `visible_cols == Some([col_idx])` (a **projected physical table** like `VALUES(Table[Col])`),
       `CALCULATE` treats it as a **column filter** on that one visible column, applying the set of
       visible values (rather than filtering by representative physical rows).
 
     This preserves expected DAX patterns like:
 
     - `CALCULATE([Measure], FILTER(VALUES(T[Col]), ...))`
 
     Current limitation: projected physical table filters support **only one visible column** (otherwise
     the engine errors).

   Examples:
   - `FILTER(Fact, Fact[Amount] > 0)`
   - `ALLEXCEPT(Dim, Dim[Category])`
   - `CALCULATETABLE(...)`

   Current limitation: virtual tables (e.g. `SUMMARIZE(...)` / `SUMMARIZECOLUMNS(...)`) are not
   accepted as `CALCULATE` table filter arguments.

#### Unsupported filter argument shapes

Notable unsupported patterns:

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
  - When `GroupByColumn.table != base_table`, the pivot engine resolves a **unique active relationship
    path** from `base_table` to the group-by table (a `ManyToOne` chain), honoring relationship overrides
    from the `FilterContext` (`USERELATIONSHIP`, `CROSSFILTER`).
    - If no path exists, or there are multiple active paths, pivot returns an evaluation error.
    - Relationship path resolution consults `FilterContext` overrides:
      - `USERELATIONSHIP` activation (and the “override pairs” semantics described above)
      - `CROSSFILTER(..., NONE)` disabling relationships
    - If every hop maps a key to at most one row, the group key is computed via chained lookups (like
      `RELATED`, possibly multi-hop).
    - If any hop maps a key to multiple rows (many-to-many), pivot **expands group keys** by enumerating
      related rows, collecting distinct attribute values per grouping column, and taking the cartesian
      product across grouping columns (measure duplication / double-counting risk).
- `measures`: list of named expressions to evaluate per group.
  - A `PivotMeasure` is *not* required to correspond to a named model measure; `expression` is parsed as DAX.
- `filter`: the initial filter context applied to the pivot query.
  - When evaluating each group, pivot clones this filter, adds equality filters for the group key values, and
    populates `FilterContext.in_scope_columns` with the group-by columns. This is what makes `ISINSCOPE`
    behave like it does in pivot visuals; standalone `DaxEngine::evaluate(...)` calls do not infer scope.

`PivotResult` contains:

- `columns`: group-by columns formatted as `"Table[Column]"`, followed by `PivotMeasure.name`
- `rows`: one row per group, sorted lexicographically by group keys using an Excel-like cross-type
  ordering (`Number` < `Text` < `Boolean` < `BLANK`). Text keys are sorted case-insensitively (Unicode-aware
  uppercasing) with a deterministic case-sensitive tiebreak so the ordering remains total.

### Performance paths

`pivot(...)` chooses the fastest applicable strategy:

1. **Columnar group-by + planned measures** (`pivot_columnar_group_by`)  
    Fast path when:
    - all `group_by` columns are on `base_table`, and
    - the backend supports `group_by_aggregations`, and
    - every measure can be “planned” into a small set of aggregations + arithmetic

    Planned expressions support:
    - Aggregations over `base_table` columns:
      - `SUM`, `AVERAGE`, `MIN`, `MAX`
      - `COUNT`, `COUNTA`, `COUNTBLANK`
      - `DISTINCTCOUNT`
      - `COUNTROWS(base_table)`
    - arithmetic (`+ - * /`), unary `-`
    - text concatenation (`&`)
    - comparisons (`= <> < <= > >=`)
    - boolean ops (`&&`/`||`), plus `NOT`, `AND`, `OR`
    - `IF`, `ISBLANK`, `COALESCE`, `DIVIDE`
    - references to named measures that expand to the above

2. **Columnar groups + per-group measure evaluation** (`pivot_columnar_groups_with_measure_eval`)  
    Fast path when:
    - `group_by` is non-empty and only uses `base_table` columns
    - the backend can produce distinct group keys quickly

    Measures are then evaluated via `DaxEngine` per group. This is still *O(groups × measure_cost)*.

3. **Row-scan many-to-many group expansion** (`pivot_row_scan_many_to_many`)  
    Used when:
    - any relationship hop needed to resolve a group-by column has non-unique keys (many-to-many)

    The engine scans base rows, expands each base row into one or more group keys (cartesian product
    of per-column related values), then evaluates measures per group. This path can be much slower
    than the columnar fast paths and can generate many groups.

4. **Columnar star-schema group-by + planned measures** (`pivot_columnar_star_schema_group_by`)  
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

5. **Row-scan planned group-by** (`pivot_planned_row_group_by`)  
    Works for non-columnar backends (and supports grouping by related dimension columns via a unique active
    relationship path), but still requires that measures are plannable (same restrictions as path 1).

6. **Fallback row-scan** (`pivot_row_scan`)  
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

### Pivot helpers

The pivot module also exposes a few helper APIs that are useful when bridging from Excel-like pivot
configurations:

- `ValueFieldAggregation`, `ValueFieldSpec` — Excel-style value-field settings (Sum/Avg/Count/etc).
- `measures_from_value_fields(base_table, &[ValueFieldSpec { ... }]) -> Vec<PivotMeasure>` — builds DAX
  measures for those value fields (mapping `Count` → `COUNTA`, `CountNumbers` → `COUNT`, etc).

### Debugging / tracing

Set `FORMULA_DAX_PIVOT_TRACE=1` to print which pivot execution path was used (once per path per process).
Current labels include `columnar_group_by`, `columnar_groups_with_measure_eval`,
`columnar_star_schema_group_by`, `planned_row_group_by`, and `row_scan`.

To force the engine to skip the star-schema fast path (useful for debugging), set
`FORMULA_DAX_PIVOT_DISABLE_STAR_SCHEMA=1`.

### Rendering integration

If built with the crate feature `pivot-model`, the pivot results can be converted into
`formula-model` scalar types:

- `PivotResult::into_pivot_values()` / `PivotResultGrid::into_pivot_values()` → `PivotValue` (the
  canonical pivot scalar representation used across IPC/XLSX/model layers)
- `PivotResultGrid::to_pivot_scalars()` → `ScalarValue` (a legacy helper used by some in-memory
  pivot structures)

---

## Supported syntax

Parsing is implemented in `crates/formula-dax/src/parser.rs`.

Supported expression forms:

- Numbers: `1`, `1.5`, `.25`, `1e3`, `1E-3`
- Strings: `"hello"` with `""` as an escape for `"`
- Identifiers: `SUM`, `Fact`, `MyTable` (ASCII letters/`_`/`.`; may include digits after the first char)
- Quoted identifiers: `'Table Name'` (escape `'` as `''`)
- Bracket identifiers:
  - measures: `[Total]`
  - column names: `Fact[Amount]` (column part is bracketed)
  - escape `]` inside bracket identifiers as `]]` (e.g. `Table[Col]]Name]`)
  - note: `[Name]` is ambiguous (measure vs. column-in-row-context). The engine resolves it as a
    measure if one exists, otherwise it resolves it as a column on the current row-context table (or
    a unique virtual binding) if available.
- Function calls: `NAME(arg1, arg2, ...)` (comma- or semicolon-separated arguments)
- Unary minus: `-expr`
- Binary operators:
  - arithmetic: `+ - * /`
  - text concatenation: `&` (single ampersand)
  - comparisons: `= <> < <= > >=`
  - membership:
    - scalar: `expr IN tableExpr` (RHS must evaluate to a one-column table)
    - row constructor: `(expr1, expr2, ...) IN tableExpr` (RHS must have the same number of columns)
  - boolean: `&&` and `||`
- Parentheses for grouping: `(expr)`
- Row constructors / tuples: `(expr1, expr2, ...)` (comma- or semicolon-separated)  
  Parsed as a row constructor (`Expr::Tuple`). Row constructors are supported on the LHS of multi-column
  `IN`, and as the row syntax inside multi-column table constructors (e.g. `{(1,2), (3,4)}`).
  They are not a first-class scalar value: using a row constructor outside of `IN` will error.
- Variables:
  - `VAR Name = <expr> ... RETURN <expr>` (one or more `VAR` bindings)
  - Variables are referenced by bare identifiers (parsed as `Expr::TableName`) and can be **scalar** or
    **table** valued.
  - Variable names are resolved via `normalize_ident` (trimmed + case-insensitive; same rules as tables/columns/measures).
- Table constructors: `{ 1, 2, 3 }` (one column) and `{ (1, 2), (3, 4) }` (multi-column row tuples)  
  Separators may be `,` or `;`. Nested table constructors are not supported, and all rows must have the
  same number of values.
  - One-column constructors evaluate to a **virtual one-column table** with a synthetic column
    named `[Value]` and can be used on the RHS of the `IN` operator (`expr IN { ... }`), as the
    first argument to `CONTAINSROW`, and as a table expression in iterators and table functions.
  - Multi-column constructors evaluate to a **virtual multi-column table** with synthetic columns
    named `[Value1]`, `[Value2]`, ... in order. They can be used as table expressions in iterators, and
    on the RHS of row-constructor `IN` (`(a,b) IN {(1,2), (3,4)}`) / `CONTAINSROW`.
  Table literals can also be bound to a `VAR` and referenced by name.

Unsupported / not yet implemented in the parser:

- Multi-line / query-style DAX (`EVALUATE`, `DEFINE MEASURE`, etc.)

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
- `CONCATENATEX(tableExpr, textExpr, [delimiter], [orderByExpr], [order])`  
  (limited: `order` must be `ASC` or `DESC`; if any `orderByExpr` key evaluates to `Text`, sorting is text-based
  with Excel-like case-insensitive ordering; otherwise sorting is numeric. Mixed text+numeric keys are allowed and
  will sort as text.)
- `HASONEVALUE(Table[Column])`
- `ISINSCOPE(Table[Column])`  
  (pivot-driven: returns `TRUE` when the pivot engine marks the column as “in scope” via
  `FilterContext.in_scope_columns`; standalone DAX evaluation does not infer scope)
- `SELECTEDVALUE(Table[Column], [alternateResult])`
- `LOOKUPVALUE(ResultTable[ResultCol], SearchTable[SearchCol1], SearchValue1, ..., [alternateResult])`  
  (current MVP restriction: all search columns must be in the same table as the result column)
- `CALCULATE(expr, filter1, filter2, ...)`
- `RELATED(Table[Column])` (requires row context)
- `CONTAINSROW(tableExpr, value1, ..., valueN)`  
  (supports multi-column matching when `tableExpr` yields multiple columns; one-column table constructors like
  `{1,2,3}` expose a single implicit column named `Value`, so `CONTAINSROW({1,2,3}, 2)` works as expected)
- `EARLIER(Table[Column], [level])` (requires nested row context)
- `EARLIEST(Table[Column])` (requires row context)

### Table functions

- `FILTER(tableExpr, predicateExpr)`
- `ALL(Table)` and `ALL(Table[Column])`
- `ALLNOBLANKROW(Table)` and `ALLNOBLANKROW(Table[Column])`
- `VALUES(Table[Column])`, `VALUES(tableExpr)`
- `DISTINCT(Table[Column])`, `DISTINCT(tableExpr)` (implemented in terms of `VALUES`)
- `ALLEXCEPT(Table, Table[Col1], Table[Col2], ...)`
- `CALCULATETABLE(tableExpr, filter1, filter2, ...)`
- `SUMMARIZE(tableExpr, Table[GroupCol1], Table[GroupCol2], ...)`  
  (limited: currently only grouping columns are supported; group columns may be on the base table or on
  related tables reachable via a unique active relationship path; grouping across many-to-many hops uses
  the expansion semantics described in the Relationships section; the base table must be physical; it
  returns a **virtual table** containing the grouping columns only)
- `SUMMARIZECOLUMNS(Table[GroupCol1], Table[GroupCol2], ..., [filterArgs...])`  
  (limited: supports leading grouping columns plus optional `CALCULATE`-style filter arguments; the
  engine picks a base table that can reach all grouped tables via active relationships; it returns a
  **virtual table** containing the grouping columns only; grouping across many-to-many hops uses the
  expansion semantics described in the Relationships section; name/expression pairs are accepted but are
  not yet materialized in the returned table representation)
- `RELATEDTABLE(Table)` (requires row context)

### Filter modifiers inside `CALCULATE`

- `USERELATIONSHIP(TableA[Col], TableB[Col])`
- `CROSSFILTER(TableA[Col], TableB[Col], BOTH|ONEWAY|SINGLE|ONEWAY_LEFTFILTERSRIGHT|ONEWAY_RIGHTFILTERSLEFT|NONE)`
- `ALLNOBLANKROW(Table|Table[Column])`
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
  - `OneToMany`, `OneToOne`, and `ManyToMany` are supported (many-to-many uses distinct-key propagation).
  - Only single-column relationships are supported.
  - `RELATED` errors when a relationship key matches multiple rows on the `to_table` side (ambiguous scalar lookup).
  - Grouping by columns across a many-to-many relationship uses expansion semantics (used by pivot, `SUMMARIZE`,
    and `SUMMARIZECOLUMNS`): a base row can contribute to multiple group keys (and combinations), which can
    duplicate measure contributions.
- **DAX language coverage**
  - Variables (`VAR`/`RETURN`) are supported.
  - Table constructors (`{ ... }`) support both one-column and multi-column row tuples (no nesting),
    and evaluate to a virtual table with synthetic columns (`[Value]` for one column, `[Value1]`,
    `[Value2]`, ... for multi-column). These are usable in iterators, as well as membership tests via
    `CONTAINSROW` and `IN` (including tuple `IN` for multi-column row constructors).
  - Most scalar/table functions are unimplemented (anything not listed above).
- **Types**
  - Only `Blank`, `Number(ordered_float::OrderedFloat<f64>)`, `Boolean`, and `Text` exist at the DAX layer.
- **Calculated columns**
  - Calculated columns are supported for both in-memory and columnar tables, but columnar calculated
    columns currently require a single logical type across all non-blank rows (number/string/boolean).
  - `DataModel::insert_row(...)` is not supported for columnar tables (they are immutable).
  - When loading persisted models where calculated column values are already stored, use
    `add_calculated_column_definition(...)` to register metadata without re-evaluating.
- **Table semantics**
  - Table expressions evaluate to either:
    - `TableResult::Physical { table, rows, visible_cols }` (sparse physical row set)
    - `TableResult::PhysicalAll { table, row_count, visible_cols }` (all physical rows; avoids allocating)
    - `TableResult::PhysicalMask { table, mask, visible_cols }` (dense physical row bitmap)
    - `TableResult::Virtual { columns, rows }` (materialized rows with explicit `(table, column)` lineage).
  - `SUMMARIZE`/`SUMMARIZECOLUMNS` return virtual tables of grouping columns only (no computed columns).
  - `SUMMARIZE` currently requires a physical base table argument.
  - `CALCULATE` *table filter* arguments (table expressions used as row filters) must currently evaluate
    to a physical table result; virtual tables are rejected.
- **`/` operator**
  - The `/` operator performs raw `f64` division; use `DIVIDE(...)` for DAX-like blank/alternate behavior.
