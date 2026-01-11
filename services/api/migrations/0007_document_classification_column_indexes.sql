-- Additional indexes to accelerate column selector lookups.
--
-- The initial normalized selector migration adds indexes for:
-- - (document_id, scope, sheet_id)
-- - cell exact lookup
-- - range bounds
--
-- Column selectors are commonly accessed by `column_index`/`column_id`, so add
-- dedicated indexes to avoid scanning all column classifications on a sheet.
CREATE INDEX IF NOT EXISTS document_classifications_column_index_lookup_idx
  ON document_classifications(document_id, scope, sheet_id, table_id, column_index);

CREATE INDEX IF NOT EXISTS document_classifications_column_id_lookup_idx
  ON document_classifications(document_id, scope, sheet_id, table_id, column_id);

