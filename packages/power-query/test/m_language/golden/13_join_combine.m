let
  Sales = Query.Reference("q_sales"),
  MoreSales = Query.Reference("q_more_sales"),
  #"Appended Queries" = Table.Combine({Sales, MoreSales}),
  Targets = Query.Reference("q_targets"),
  #"Merged Queries" = Table.Join(#"Appended Queries", {"Id"}, Targets, {"Id"}, JoinKind.LeftOuter)
in
  #"Merged Queries"
