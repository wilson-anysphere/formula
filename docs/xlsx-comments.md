# XLSX Comments (Legacy Notes + Threaded Comments)

Excel has *two* comment systems in modern `.xlsx` files:

1. **Legacy “notes”** (historical cell comments)
2. **Modern “threaded comments”** (replies, resolved state, author identities)

This document describes:

- The common OPC part layout + relationship types
- Our parsing approach (what we interpret vs ignore)
- Our **preservation contract** (what is kept byte-for-byte)

---

## Part layout (typical)

The following parts are commonly seen together:

```
xl/
├── worksheets/
│   ├── sheet1.xml
│   └── _rels/sheet1.xml.rels
├── comments1.xml                       # legacy notes payload (text + authors table)
├── threadedComments/
│   └── threadedComments1.xml           # threaded comments payload (roots + replies)
├── persons/
│   └── persons1.xml                    # workbook-level people directory (optional)
├── commentsExt1.xml                    # comment extension metadata (optional; may be unreferenced)
└── drawings/
    └── vmlDrawing1.vml                 # VML shapes for legacy notes (anchors, visibility, etc.)
```

Notes:

- `sheetN.xml` usually contains a `<legacyDrawing r:id="…"/>` element pointing at the VML part.
- The `comments*.xml` and `threadedComments/*.xml` parts are typically referenced only from
  `sheetN.xml.rels` via relationship *types* (not via explicit elements in `sheetN.xml`).
- `commentsExt*.xml` is frequently present in real Excel files but may not have an explicit `.rels`
  entry. Treat it as a comment-adjacent sidecar.

---

## Relationship type URIs (must preserve)

From the fixture `crates/formula-xlsx/tests/fixtures/comments.xlsx`:

### Worksheet relationships (`xl/worksheets/_rels/sheetN.xml.rels`)

- VML drawing (legacy notes shapes):
  - `http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing`
- Legacy comments part (`xl/commentsN.xml`):
  - `http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments`
- Threaded comments part (`xl/threadedComments/threadedCommentsN.xml`):
  - `http://schemas.microsoft.com/office/2017/10/relationships/threadedComment`

### Workbook relationships (`xl/_rels/workbook.xml.rels`)

- Persons directory (`xl/persons/personsN.xml`):
  - `http://schemas.microsoft.com/office/2017/10/relationships/person`

### Content types (non-exhaustive)

Excel commonly advertises these in `[Content_Types].xml`:

- `application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml` (`/xl/comments*.xml`)
- `application/vnd.ms-excel.threadedcomments+xml` (`/xl/threadedComments/threadedComments*.xml`)
- `application/vnd.ms-excel.commentsExt+xml` (`/xl/commentsExt*.xml`)
- `application/vnd.openxmlformats-officedocument.vmlDrawing` (`.vml`)

---

## Parsing strategy

### Legacy notes (`xl/comments*.xml` + VML)

- **Parse**:
  - `comment/@ref` (cell anchor)
  - `comment/@authorId` + `<authors>` table (best-effort author name)
  - visible text by concatenating `<t>` nodes under `<text>`
- **Do not interpret**:
  - VML shape geometry, styling, visibility, margins, etc. (these live in `vmlDrawing*.vml`)

We also parse the VML drawing *only* to recover the cell anchors for note shapes when needed
(see `formula-xlsx::comments::parse_vml_drawing_cells`), but we intentionally avoid modeling the
full VML schema.

### Threaded comments (`xl/threadedComments/threadedComments*.xml`)

- **Parse**:
  - `threadedComment/@ref` (cell)
  - `threadedComment/@id` and `@parentId` (reply threading)
  - `threadedComment/@done` (resolved state)
  - `threadedComment/@personId` / `@author` + optional lookup via `xl/persons/*.xml`
  - visible text (concatenate `<t>` nodes; best-effort)
- **Do not interpret**:
  - unsupported extension payloads, formatting runs beyond visible text, reactions, @mentions, etc.

### Persons (`xl/persons/*.xml`)

`persons*.xml` is treated as an optional directory that maps stable IDs to display names. If it is
missing, we fall back to author attributes stored directly on threaded comment elements.

---

## Preservation strategy

### What we parse into the model

We parse a workbook’s comments into `formula_model::Comment` records:

- comment kind (`Note` vs `Threaded`)
- anchor cell (`CellRef`)
- author id + best-effort display name
- timestamps when available (threaded comments)
- resolved state + replies (threaded comments)
- visible text content

This model is meant to power UI and collaboration features without requiring a full OOXML
implementation of VML / comment extensions.

### What we preserve byte-for-byte

Unless the caller explicitly requests a rewrite, we preserve **all comment-related parts** exactly
as they appeared in the input ZIP:

- `xl/comments*.xml`
- `xl/threadedComments/*`
- `xl/persons/*`
- `xl/drawings/vmlDrawing*.vml`
- `xl/commentsExt*.xml`
- any `*.rels` parts that reference comment artifacts

### What we regenerate (when explicitly editing comments)

When a caller provides an updated comment set, we regenerate only the primary XML payload parts:

- `xl/comments*.xml` (legacy notes)
- `xl/threadedComments/threadedComments*.xml` (threaded comments)

All other comment-related parts remain preserved verbatim. In particular:

- We do **not** attempt to regenerate VML shapes (`vmlDrawing*.vml`).
- We do **not** attempt to normalize or rewrite `commentsExt*.xml`.
- Relationship IDs (`rId*`) and relationship targets are preserved whenever possible.

Implementation entry points:

- Read: `formula_xlsx::comments::extract_comment_parts`
- Write: `formula_xlsx::XlsxPackage::write_comment_parts`

