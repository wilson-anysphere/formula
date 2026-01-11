let
  Source = Web.Contents("https://example.com/data", [Method="POST", Headers=[Authorization="Bearer token"]]),
  #"Removed Columns" = Table.RemoveColumns(Source, {"Secret"})
in
  #"Removed Columns"

