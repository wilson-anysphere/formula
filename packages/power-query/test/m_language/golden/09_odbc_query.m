let
  Source = Odbc.Query("dsn=mydb", "select * from sales"),
  #"Selected Columns" = Table.SelectColumns(Source, {"Region"})
in
  #"Selected Columns"

