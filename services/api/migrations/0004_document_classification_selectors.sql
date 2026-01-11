-- Add normalized selector columns to document_classifications for efficient lookup.
--
-- These columns duplicate data already present in the selector JSON, but allow:
-- - fast exact-match lookups for cells/columns/sheets
-- - containment queries (cell/range within a classified range)
-- - overlap queries (aggregate classification for a selected range)
--
-- NOTE: We keep backwards compatibility by backfilling existing rows from selector.
ALTER TABLE document_classifications
  ADD COLUMN IF NOT EXISTS scope text,
  ADD COLUMN IF NOT EXISTS sheet_id text,
  ADD COLUMN IF NOT EXISTS table_id text,
  ADD COLUMN IF NOT EXISTS row integer,
  ADD COLUMN IF NOT EXISTS col integer,
  ADD COLUMN IF NOT EXISTS start_row integer,
  ADD COLUMN IF NOT EXISTS start_col integer,
  ADD COLUMN IF NOT EXISTS end_row integer,
  ADD COLUMN IF NOT EXISTS end_col integer,
  ADD COLUMN IF NOT EXISTS column_index integer,
  ADD COLUMN IF NOT EXISTS column_id text;

-- Backfill normalized columns from selector JSON.
UPDATE document_classifications
SET
  scope = selector->>'scope',
  sheet_id = CASE
    WHEN selector->>'scope' IN ('sheet', 'column', 'range', 'cell') THEN selector->>'sheetId'
    ELSE NULL
  END,
  table_id = CASE
    WHEN selector->>'scope' = 'column' THEN selector->>'tableId'
    ELSE NULL
  END,
  row = CASE
    WHEN selector->>'scope' = 'cell' THEN (selector->>'row')::integer
    ELSE NULL
  END,
  col = CASE
    WHEN selector->>'scope' = 'cell' THEN (selector->>'col')::integer
    ELSE NULL
  END,
  start_row = CASE
    WHEN selector->>'scope' = 'range' THEN LEAST(
      (selector -> 'range' -> 'start' ->> 'row')::integer,
      (selector -> 'range' -> 'end' ->> 'row')::integer
    )
    ELSE NULL
  END,
  end_row = CASE
    WHEN selector->>'scope' = 'range' THEN GREATEST(
      (selector -> 'range' -> 'start' ->> 'row')::integer,
      (selector -> 'range' -> 'end' ->> 'row')::integer
    )
    ELSE NULL
  END,
  start_col = CASE
    WHEN selector->>'scope' = 'range' THEN LEAST(
      (selector -> 'range' -> 'start' ->> 'col')::integer,
      (selector -> 'range' -> 'end' ->> 'col')::integer
    )
    ELSE NULL
  END,
  end_col = CASE
    WHEN selector->>'scope' = 'range' THEN GREATEST(
      (selector -> 'range' -> 'start' ->> 'col')::integer,
      (selector -> 'range' -> 'end' ->> 'col')::integer
    )
    ELSE NULL
  END,
  column_index = CASE
    WHEN selector->>'scope' = 'column' AND (selector->>'columnIndex') IS NOT NULL THEN (selector->>'columnIndex')::integer
    ELSE NULL
  END,
  column_id = CASE
    WHEN selector->>'scope' = 'column' THEN selector->>'columnId'
    ELSE NULL
  END
WHERE scope IS NULL;

-- Lookup indexes.
CREATE INDEX IF NOT EXISTS document_classifications_document_scope_sheet_idx
  ON document_classifications(document_id, scope, sheet_id);

CREATE INDEX IF NOT EXISTS document_classifications_cell_lookup_idx
  ON document_classifications(document_id, scope, sheet_id, row, col);

-- Best-effort containment/overlap acceleration for rectangle ranges.
CREATE INDEX IF NOT EXISTS document_classifications_range_bounds_idx
  ON document_classifications(document_id, scope, sheet_id, start_row, start_col, end_row, end_col);
