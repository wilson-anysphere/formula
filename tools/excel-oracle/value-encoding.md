# Excel Oracle value encoding

Oracle and engine results use a common JSON encoding so comparisons are:

- unambiguous (blank vs 0, errors vs numbers)
- stable across languages/runtimes
- able to represent **spilled arrays**

## Scalar values

Each scalar is an object:

| Type | Shape | Example |
|------|-------|---------|
| blank | `{"t":"blank"}` | empty cell |
| number | `{"t":"n","v":1.5}` | `1.5` |
| string | `{"t":"s","v":"hello"}` | `"hello"` |
| boolean | `{"t":"b","v":true}` | `TRUE` |
| error | `{"t":"e","v":"#DIV/0!"}` | `#DIV/0!` |

`v` is always JSON-native (number/string/bool).

## Spilled arrays

If a formula spills, the result is:

```json
{
  "t": "arr",
  "rows": [
    [ {"t":"n","v":1}, {"t":"n","v":2} ],
    [ {"t":"n","v":3}, {"t":"n","v":4} ]
  ]
}
```

Empty cells inside a spill are represented with `{"t":"blank"}`.

## Display text (optional)

The Excel oracle runner also records the **display text** Excel shows in the grid for debugging.

Comparisons should be done on the typed values above, not on display text (which can vary with locale and formatting).

## Harness-level errors (`#ENGINE!`)

When producing *engine* results (not Excel results), the harness may emit:

```json
{"t":"e","v":"#ENGINE!","detail":"..."}
```

This indicates the engine runner failed to evaluate a case (parse error, unsupported input value encoding, etc). The `detail` field is for debugging and is not emitted by the Excel oracle runner.
