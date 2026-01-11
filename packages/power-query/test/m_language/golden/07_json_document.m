let
  Source = Json.Document(File.Contents("/tmp/data.json"), "items"),
  #"Selected Columns" = Table.SelectColumns(Source, {"id", "name"})
in
  #"Selected Columns"

