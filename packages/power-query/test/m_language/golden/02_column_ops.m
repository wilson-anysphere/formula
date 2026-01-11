let
  Source = Range.FromValues({
    {"Region", "Sales", "Product"},
    {"East", "100", "A"},
    {"West", "200", "B"}
  }),
  #"Changed Type" = Table.TransformColumnTypes(Source, {{"Sales", type number}}),
  #"Added Column" = Table.AddColumn(#"Changed Type", "Double", each [Sales] * 2),
  #"Removed Columns" = Table.RemoveColumns(#"Added Column", {"Product"}),
  #"Selected Columns" = Table.SelectColumns(#"Removed Columns", {"Region", "Sales", "Double"}),
  #"Renamed Columns" = Table.RenameColumns(#"Selected Columns", {{"Double", "Double Sales"}, {"Region", "Area"}})
in
  #"Renamed Columns"

