let
  Source = SharePoint.Files(
    "https://contoso.sharepoint.com/sites/Finance",
    [
      Recursive = true,
      IncludeContent = false,
      Auth = [Type = "oauth2", ProviderId = "example", Scopes = {"Sites.Read.All"}]
    ]
  ),
  #"Selected Columns" = Table.SelectColumns(Source, {"Name"})
in
  #"Selected Columns"

