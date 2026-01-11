let
  Source = Odbc.Query("Driver={PostgreSQL Unicode};Server=db.example.com;Port=5432;Database=analytics;Uid=alice;", "select * from sales"),
  #"Selected Columns" = Table.SelectColumns(Source, {"Region"})
in
  #"Selected Columns"
