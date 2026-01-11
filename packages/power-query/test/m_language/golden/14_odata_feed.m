let
  Source = OData.Feed("https://example.com/odata/Products", [Headers=[Authorization="Bearer token"]]),
  #"Selected Columns" = Table.SelectColumns(Source, {"Id", "Name"})
in
  #"Selected Columns"

