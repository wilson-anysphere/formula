# `@formula/collab-yjs-utils`

Shared helpers for working with Yjs in mixed-module environments (ESM + CJS).

## Why this exists

In some environments we can end up with **multiple `yjs` module instances** loaded at once:

- Node ESM app code importing `yjs`
- a dependency (e.g. `y-websocket`) pulling in the CJS build via `require("yjs")`

When updates are applied through a different module instance, a `Y.Doc` can contain
**foreign constructors** (types that fail `instanceof` checks against the local `yjs`
import). Yjs also performs strict constructor checks when instantiating roots via
`doc.getMap/getArray/getText`, which can throw `"different constructor"`.

This package centralizes the “duck-typing” + root-normalization logic so the rest
of the collab stack doesn’t have to re-implement it.

## Exports

- `getYMap`, `getYArray`, `getYText` – duck-type checks tolerant of foreign constructors.
- `isYAbstractType` – structural `AbstractType` detection without relying solely on `instanceof`.
- `getDocTypeConstructors` – derive Y.Map/Y.Array/Y.Text constructors compatible with a given doc (useful when cloning values into foreign-module docs).
- `cloneYjsValue` – deep-clone a Yjs value into freshly-constructed Yjs types (optionally using constructors from `getDocTypeConstructors`).
- `replaceForeignRootType` – normalize a foreign root into this module’s constructors.
- `getMapRoot`, `getArrayRoot`, `getTextRoot` – safe root access that avoids constructor-mismatch throws.
- `patchForeignItemConstructor` – best-effort prototype patching so foreign Yjs `Item` structs pass `instanceof Y.Item` checks (used by UndoManager integrations).
- `patchForeignAbstractTypeConstructor` – best-effort prototype patching so foreign Yjs types pass `instanceof Y.AbstractType` checks (used by UndoManager integrations).
- `yjsValueToJson` / `cloneYjsValueToJson` – best-effort conversion of nested Yjs values into plain JSON-ish values.

## Notes

- Root replacement is only performed when `doc instanceof Y.Doc` for the **local**
  `yjs` import. If the entire document was created by a foreign Yjs build, mixing
  local types into it is unsafe; in that case the helpers fall back to returning
  the existing foreign types.
- `Y.Text` has an extra caveat: even when the `Y.Text` instance itself is from the
  local module, applying updates via a foreign module instance can leave its internal
  `Item.content` objects (ContentString/ContentFormat/ContentEmbed/ContentType) with
  foreign constructors. Because Yjs uses constructor equality checks for rich-text
  operations, this can make `toString()` / `toDelta()` return empty results. The
  helpers in this package patch content prototypes (by internal `getRef()` ids) in
  `getTextRoot`, `replaceForeignRootType`, and `yjsValueToJson` to keep reads stable.
