let
  Source = Range.FromValues({
    {"Name", "Sales", "When"},
    {" alice ", 12.345, "2020-01-01"},
    {"Bob", 67.891, "2020-01-02"}
  }),
  #"Upper Trimmed" = Table.AddColumn(Source, "NameUpper", each Text.Upper(Text.Trim([Name]))),
  #"Name Length" = Table.AddColumn(#"Upper Trimmed", "NameLen", each Text.Length([Name])),
  #"Has A" = Table.AddColumn(#"Name Length", "HasA", each Text.Contains([Name], "a")),
  #"Rounded" = Table.AddColumn(#"Has A", "SalesRounded", each Number.Round([Sales], 1)),
  #"Add Days" = Table.AddColumn(#"Rounded", "WhenPlus", each Date.AddDays(Date.FromText([When]), 1))
in
  #"Add Days"

