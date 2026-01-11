let
  Source = Range.FromValues({
    {"Region", "FullName"},
    {"East", "A,B"},
    {null, "C,D"},
    {"West", "E,F"}
  }),
  #"Filled Down" = Table.FillDown(Source, {"Region"}),
  #"Replaced Value" = Table.ReplaceValue(#"Filled Down", "West", "W", Replacer.ReplaceText, {"Region"}),
  #"Split Column" = Table.SplitColumn(#"Replaced Value", "FullName", ",")
in
  #"Split Column"

