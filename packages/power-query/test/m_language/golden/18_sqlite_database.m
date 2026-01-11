let
  Source = SQLite.Database("/tmp/test.db", "select * from t"),
  #"Selected Columns" = Table.SelectColumns(Source, {"A"})
in
  #"Selected Columns"

