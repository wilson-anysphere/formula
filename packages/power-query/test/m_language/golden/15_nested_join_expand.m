let
  Left = Query.Reference("q_left"),
  Right = Query.Reference("q_right"),
  #"Merged Queries" = Table.NestedJoin(Left, {"Id"}, Right, {"Id"}, "Matches", JoinKind.LeftOuter),
  #"Expanded Matches" = Table.ExpandTableColumn(#"Merged Queries", "Matches", {"Target"})
in
  #"Expanded Matches"
