let
  Source = Excel.CurrentWorkbook(){[Name="Sales"]}[Content],
  #"Filtered Rows" = Table.SelectRows(Source, each [Region] <> "West")
in
  #"Filtered Rows"

