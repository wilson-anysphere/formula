let
  Source = Csv.Document(File.Contents("/tmp/sales.csv"), [Delimiter=";", HasHeaders=true]),
  #"Selected Columns" = Table.SelectColumns(Source, {"Region", "Sales"})
in
  #"Selected Columns"

