let
  Source = Sql.Database("server1", "db1", "select * from t"),
  #"Selected Columns" = Table.SelectColumns(Source, {"A"})
in
  #"Selected Columns"

