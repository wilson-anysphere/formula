let
  Left = Query.Reference("q_left"),
  Right = Query.Reference("q_right"),
  #"Merged Queries" = Table.Join(Left, {"Id", "Region"}, Right, {"Id", "Region"}, JoinKind.Inner)
in
  #"Merged Queries"
