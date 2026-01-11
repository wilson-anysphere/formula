let
  Source = OData.Feed(
    "https://example.com/odata/Products",
    [
      Headers = [Authorization = "Bearer token", Accept = "application/json"],
      Auth = [Type = "OAuth2", ProviderId = "example-provider", Scopes = {"scope1", "scope2"}],
      RowsPath = "data.items"
    ]
  )
in
  Source

