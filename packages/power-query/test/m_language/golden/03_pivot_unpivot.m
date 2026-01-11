let
  Source = Range.FromValues({
    {"Region", "Metric", "Value"},
    {"East", "Sales", 100},
    {"East", "Profit", 20},
    {"West", "Sales", 200},
    {"West", "Profit", 50}
  }),
  #"Pivoted" = Table.Pivot(Source, {"Sales", "Profit"}, "Metric", "Value", List.Sum),
  #"Unpivoted" = Table.Unpivot(#"Pivoted", {"Sales", "Profit"}, "Metric", "Value")
in
  #"Unpivoted"

