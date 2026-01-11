let
  Source = Range.FromValues({
    {"Region", "Sales"},
    {"East", 100},
    {"West", 200},
    {"South", 300}
  }),
  #"Added Index" = Table.AddIndexColumn(Source, "Index", 1, 1),
  #"Reordered Columns" = Table.ReorderColumns(#"Added Index", {"Index", "Region", "Sales"}),
  #"Removed Top Rows" = Table.Skip(#"Reordered Columns", 1),
  #"Kept First Rows" = Table.FirstN(#"Removed Top Rows", 1)
in
  #"Kept First Rows"

