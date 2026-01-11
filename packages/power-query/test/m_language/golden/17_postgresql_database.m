let
  Source = PostgreSQL.Database("db.example.com:5432", "analytics", "select * from sales"),
  #"Selected Columns" = Table.SelectColumns(Source, {"Region"})
in
  #"Selected Columns"

