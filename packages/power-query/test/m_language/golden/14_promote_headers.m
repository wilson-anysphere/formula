let
  Source = Range.FromValues(
    {
      {"Region", "Sales"},
      {"East", 100},
      {"West", 200}
    },
    [HasHeaders = false]
  ),
  #"Promoted Headers" = Table.PromoteHeaders(Source),
  #"Demoted Headers" = Table.DemoteHeaders(#"Promoted Headers")
in
  #"Demoted Headers"

