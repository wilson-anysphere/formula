Table.SelectColumns(
  Table.SelectRows(
    Range.FromValues({
      {"A", "B"},
      {1, "x"},
      {2, "y"}
    }),
    each [A] > 1
  ),
  {"B"}
)

