let
  Source = Range.FromValues({
    {"Date", "Region", "Sales"},
    {#date(2024, 1, 1), "East", 100},
    {#date(2024, 1, 2), "East", 150},
    {#date(2024, 1, 3), "West", 200}
  }),
  #"Filtered Rows" = Table.SelectRows(Source, each [Date] >= #date(2024, 1, 2) and [Region] = "East"),
  #"Grouped Rows" = Table.Group(#"Filtered Rows", {"Region"}, {{"Total Sales", each List.Sum([Sales])}}),
  #"Sorted Rows" = Table.Sort(#"Grouped Rows", {{"Total Sales", Order.Descending}})
in
  #"Sorted Rows"

