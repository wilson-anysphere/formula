let
  Source = Range.FromValues({
    {"Id", "Sales"},
    {1, 10},
    {1, -5},
    {2, 0}
  }),
  #"Added Column" = Table.AddColumn(Source, "Positive", each if [Sales] > 0 then [Sales] else null),
  #"Removed Errors" = Table.RemoveRowsWithErrors(#"Added Column"),
  #"Removed Duplicates" = Table.Distinct(#"Removed Errors", {"Id"}),
  #"Transformed Columns" = Table.TransformColumns(#"Removed Duplicates", {{"Positive", each if _ = null then 0 else _, type number}})
in
  #"Transformed Columns"
