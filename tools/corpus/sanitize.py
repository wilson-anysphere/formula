from __future__ import annotations

import hashlib
import io
import posixpath
import re
import zipfile
from dataclasses import dataclass
from typing import Iterable
from xml.etree import ElementTree as ET


NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"

_CELLIMAGES_BASENAME_RE = re.compile(r"^cellimages\d*\.xml$")


def qn(ns: str, tag: str) -> str:
    return f"{{{ns}}}{tag}"


@dataclass(frozen=True)
class SanitizeOptions:
    # Cell-level anonymization
    redact_cell_values: bool = True
    hash_strings: bool = False
    hash_salt: str | None = None

    # Privacy controls for workbook metadata & external surfaces
    remove_external_links: bool = True
    remove_secrets: bool = True
    scrub_metadata: bool = True

    # Optional: deterministically rename sheets to Sheet1, Sheet2, ...
    # This is off by default because it requires rewriting references in formulas.
    rename_sheets: bool = False


@dataclass(frozen=True)
class SanitizeSummary:
    removed_parts: list[str]
    rewritten_parts: list[str]


@dataclass(frozen=True)
class LeakScanFinding:
    part_name: str
    kind: str  # plaintext | email | url | aws_key | jwt
    match_sha256: str


@dataclass(frozen=True)
class LeakScanResult:
    findings: list[LeakScanFinding]

    @property
    def ok(self) -> bool:
        return not self.findings


@dataclass(frozen=True)
class TableContext:
    """A rectangular A1 range plus a per-table column rename map.

    Used to rewrite *unqualified* structured references inside table formulas
    (e.g. `[@Amount]`) based on which table the formula cell belongs to.
    """

    start_row: int
    end_row: int
    start_col: int
    end_col: int
    column_map: dict[str, str]


def _hash_text(value: str, *, salt: str) -> str:
    # Use a stable, corpus-level salt so identical strings hash identically across files,
    # but remain resistant to rainbow-table attacks when the salt is private.
    digest = hashlib.sha256((salt + "\0" + value).encode("utf-8")).hexdigest()
    return f"H_{digest[:16]}"


def _require_hash_salt(options: SanitizeOptions) -> str:
    if options.hash_strings and not options.hash_salt:
        raise ValueError("hash_strings requires hash_salt")
    return options.hash_salt or ""


def _sanitize_text(value: str, *, options: SanitizeOptions) -> str:
    if options.hash_strings:
        return _hash_text(value, salt=_require_hash_salt(options))
    return "REDACTED"


def _sanitize_xml_text_elements(root: ET.Element, *, options: SanitizeOptions, local_names: set[str]) -> None:
    for el in root.iter():
        if el.tag.split("}")[-1] not in local_names:
            continue
        if el.text is None:
            continue
        el.text = _sanitize_text(el.text, options=options)


def _sanitize_xml_attributes(
    root: ET.Element, *, options: SanitizeOptions, attr_names: set[str]
) -> None:
    for el in root.iter():
        for k, v in list(el.attrib.items()):
            if k.split("}")[-1] not in attr_names:
                continue
            el.attrib[k] = _sanitize_text(v, options=options)


def _sanitize_shared_strings(xml: bytes, *, options: SanitizeOptions) -> bytes:
    root = ET.fromstring(xml)
    if root.tag != qn(NS_MAIN, "sst"):
        return xml

    _require_hash_salt(options)

    for t in root.iter(qn(NS_MAIN, "t")):
        if t.text is None:
            continue
        if options.redact_cell_values:
            if options.hash_strings:
                t.text = _hash_text(t.text, salt=options.hash_salt or "")
            else:
                t.text = "REDACTED"
        elif options.hash_strings:
            t.text = _hash_text(t.text, salt=options.hash_salt or "")

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _sanitize_inline_string(el: ET.Element, *, options: SanitizeOptions) -> None:
    _require_hash_salt(options)

    for t in el.iter(qn(NS_MAIN, "t")):
        if t.text is None:
            continue
        if options.redact_cell_values:
            if options.hash_strings:
                t.text = _hash_text(t.text, salt=options.hash_salt or "")
            else:
                t.text = "REDACTED"
        elif options.hash_strings:
            t.text = _hash_text(t.text, salt=options.hash_salt or "")


def _sanitize_worksheet(xml: bytes, *, options: SanitizeOptions) -> bytes:
    root = ET.fromstring(xml)
    local_root = root.tag.split("}")[-1]
    if local_root.lower() not in {"worksheet", "dialogsheet", "macrosheet", "chartsheet"}:
        return xml

    for c in root.iter(qn(NS_MAIN, "c")):
        f_el = c.find(qn(NS_MAIN, "f"))
        v_el = c.find(qn(NS_MAIN, "v"))
        is_el = c.find(qn(NS_MAIN, "is"))

        if f_el is not None:
            # Preserve formulas/structure, but optionally drop cached results which can leak data.
            if options.redact_cell_values or options.hash_strings:
                if v_el is not None:
                    c.remove(v_el)
                # Inline strings do not make sense on formula cells, but be defensive.
                if is_el is not None:
                    c.remove(is_el)
            continue

        if is_el is not None:
            _sanitize_inline_string(is_el, options=options)
            continue

        t = c.get("t")
        if t in {"s"}:
            # Shared string value is stored by index; sanitization happens in sharedStrings.xml.
            # No action required here unless caller opted to fully redact values.
            if options.redact_cell_values and v_el is not None and options.hash_strings is False:
                # Keep the shared string index but content will be REDACTED; no need to mutate.
                pass
            continue

        if t in {"str"} and v_el is not None:
            _require_hash_salt(options)
            if options.redact_cell_values:
                if options.hash_strings:
                    v_el.text = _hash_text(v_el.text or "", salt=options.hash_salt or "")
                else:
                    v_el.text = "REDACTED"
            elif options.hash_strings:
                v_el.text = _hash_text(v_el.text or "", salt=options.hash_salt or "")
            continue

        if v_el is None:
            continue

        if options.redact_cell_values:
            # Numeric, boolean, error, etc. Normalize to 0 to keep sheet shape without leaking.
            if t == "e":
                v_el.text = "#N/A"
            else:
                v_el.text = "0"

    # Header/footer text can contain PII (names, phone numbers, emails, etc).
    if options.scrub_metadata or options.hash_strings:
        for hf in root.iter():
            if hf.tag.split("}")[-1] != "headerFooter":
                continue
            for child in hf.iter():
                if child is hf:
                    continue
                if child.text is None:
                    continue
                child.text = _sanitize_text(child.text, options=options)

    # Hyperlink display text / tooltips can leak URLs even when relationship targets are scrubbed.
    if options.remove_external_links or options.scrub_metadata or options.hash_strings:
        for hl in root.iter():
            if hl.tag.split("}")[-1] != "hyperlink":
                continue
            for attr in ("display", "tooltip"):
                if attr in hl.attrib:
                    hl.attrib[attr] = _sanitize_text(hl.attrib[attr], options=options)
            # Be defensive: some writers put external URLs in `location`.
            loc = hl.attrib.get("location")
            if loc and _looks_like_external_url(loc):
                hl.attrib["location"] = _sanitize_text(loc, options=options)

    if options.scrub_metadata or options.hash_strings:
        # Sheet code names can leak business terms (and are exposed to macros, which we drop).
        for el in root.iter():
            if el.tag.split("}")[-1] != "sheetPr":
                continue
            if "codeName" in el.attrib:
                el.attrib["codeName"] = _sanitize_text(el.attrib["codeName"], options=options)

    if options.remove_secrets:
        # Remove sheet-level protection/password hashes and protected range metadata.
        for parent in list(root.iter()):
            for child in list(parent):
                if child.tag.split("}")[-1] in {"sheetProtection", "protectedRanges"}:
                    parent.remove(child)

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _sanitize_core_properties(xml: bytes, *, options: SanitizeOptions) -> bytes:
    root = ET.fromstring(xml)
    # Redact common authoring metadata; leave structure intact.
    for el in root.iter():
        local = el.tag.split("}")[-1]
        if local in {
            "creator",
            "lastModifiedBy",
            "title",
            "subject",
            "description",
            "keywords",
        }:
            if el.text is not None:
                el.text = _sanitize_text(el.text, options=options)
        if local in {"created", "modified"}:
            el.text = "1970-01-01T00:00:00Z"
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _sanitize_app_properties(
    xml: bytes, *, options: SanitizeOptions, sheet_rename_map: dict[str, str] | None = None
) -> bytes:
    root = ET.fromstring(xml)
    for el in root.iter():
        local = el.tag.split("}")[-1]
        if (options.scrub_metadata or options.hash_strings) and local in {"Company", "Manager", "HyperlinkBase"}:
            if el.text is not None:
                el.text = _sanitize_text(el.text, options=options)

        if local == "TitlesOfParts":
            for title in el.iter():
                if title is el:
                    continue
                if title.tag.split("}")[-1] != "lpstr":
                    continue
                if title.text is None:
                    continue
                if sheet_rename_map and title.text in sheet_rename_map:
                    title.text = sheet_rename_map[title.text]
                elif options.scrub_metadata or options.hash_strings:
                    title.text = _sanitize_text(title.text, options=options)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _sanitize_comments(xml: bytes, *, options: SanitizeOptions) -> bytes:
    root = ET.fromstring(xml)
    # Comments are in the spreadsheetml main namespace, but be robust and match by local-name.
    if root.tag.split("}")[-1] != "comments":
        return xml
    if options.scrub_metadata or options.hash_strings:
        _sanitize_xml_text_elements(root, options=options, local_names={"t", "author"})
        # Some comment models store author info as attributes.
        _sanitize_xml_attributes(root, options=options, attr_names={"author", "userId", "displayName"})
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _sanitize_threaded_comments(xml: bytes, *, options: SanitizeOptions) -> bytes:
    root = ET.fromstring(xml)
    if options.scrub_metadata or options.hash_strings:
        # Threaded comments use newer namespaces; sanitize common text/PII fields defensively.
        _sanitize_xml_text_elements(root, options=options, local_names={"t", "text"})
        _sanitize_xml_attributes(root, options=options, attr_names={"displayName", "userId", "author"})
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _sanitize_table(
    xml: bytes,
    *,
    options: SanitizeOptions,
    table_rename_map: dict[str, str] | None = None,
    table_column_rename_map: dict[str, dict[str, str]] | None = None,
    sheet_rename_map: dict[str, str] | None = None,
) -> bytes:
    root = ET.fromstring(xml)
    if root.tag.split("}")[-1] != "table":
        return xml

    if options.scrub_metadata or options.hash_strings:
        # Table name/displayName can leak business terms and are user-visible.
        # Avoid collapsing all tables to the same identifier (Excel requires uniqueness).
        old = root.attrib.get("displayName") or root.attrib.get("name") or ""
        new = table_rename_map.get(old, "") if table_rename_map else ""
        if not new:
            if options.hash_strings:
                new = _hash_text(old, salt=_require_hash_salt(options))
            else:
                # Fall back to a deterministic, non-sensitive name that doesn't require a salt.
                new = "Table1"
        for attr in ("name", "displayName"):
            if attr in root.attrib:
                root.attrib[attr] = new

        # Table column names are duplicated metadata (separate from header cell values) and
        # are displayed in the UI. They can contain sensitive business terms, so scrub them
        # alongside table names. Structured references in formulas must be rewritten using
        # the same mapping (handled separately when sanitizing formulas).
        if table_column_rename_map is not None:
            col_map = table_column_rename_map.get(new, {})
            for col in root.iter():
                if col.tag.split("}")[-1] != "tableColumn":
                    continue
                name = col.attrib.get("name")
                if name and name in col_map:
                    col.attrib["name"] = col_map[name]
                # Totals labels are displayed text too.
                label = col.attrib.get("totalsRowLabel")
                if label:
                    col.attrib["totalsRowLabel"] = _sanitize_text(label, options=options)

            # Table column formulas can include unqualified structured references like `[@Col]`.
            # Rewrite them to match the sanitized column names.
            for el in root.iter():
                if el.tag.split("}")[-1] not in {"calculatedColumnFormula", "totalsRowFormula"}:
                    continue
                if not el.text:
                    continue
                el.text = _sanitize_formula_text(
                    el.text,
                    options=options,
                    sheet_rename_map=sheet_rename_map or {},
                    table_rename_map=table_rename_map or {},
                    table_column_rename_map=table_column_rename_map or {},
                )
                el.text = _rewrite_unqualified_structured_refs(el.text, column_map=col_map)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _sanitize_drawing(xml: bytes, *, options: SanitizeOptions) -> bytes:
    root = ET.fromstring(xml)
    if options.scrub_metadata or options.hash_strings:
        # Text in drawings (text boxes, chart titles, etc.) uses DrawingML <a:t>.
        _sanitize_xml_text_elements(root, options=options, local_names={"t"})
        _sanitize_xml_attributes(root, options=options, attr_names={"name", "descr", "title"})

    # Images can contain PII; remove entire anchors that embed raster content when secret
    # removal is enabled. Removing just `<xdr:pic>` can leave invalid anchors behind.
    if options.remove_secrets:
        for anchor in list(root):
            local = anchor.tag.split("}")[-1]
            if local not in {"twoCellAnchor", "oneCellAnchor", "absoluteAnchor"}:
                continue
            if any(el.tag.split("}")[-1] in {"pic", "blip"} for el in anchor.iter()):
                root.remove(anchor)

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _sanitize_cell_images(xml: bytes, *, options: SanitizeOptions) -> bytes:
    """Best-effort scrubber for `xl/cellImages.xml` (in-cell images).

    The schema for this part is not stable across Excel versions. We treat it as a
    DrawingML-adjacent XML blob and defensively scrub common DrawingML text runs and
    descriptive attributes that can contain user-provided metadata.
    """

    root = ET.fromstring(xml)
    if options.scrub_metadata or options.hash_strings:
        # Cell images store shape metadata using DrawingML too.
        _sanitize_xml_text_elements(root, options=options, local_names={"t"})
        _sanitize_xml_attributes(root, options=options, attr_names={"name", "descr", "title"})
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _is_cellimages_xml_part(part_name: str) -> bool:
    """Return True for any `xl/**/cellimages*.xml` part (case-insensitive)."""

    if not part_name.casefold().startswith("xl/"):
        return False
    return _CELLIMAGES_BASENAME_RE.fullmatch(posixpath.basename(part_name).casefold()) is not None


def _sanitize_vml_drawing(xml: bytes, *, options: SanitizeOptions) -> bytes:
    root = ET.fromstring(xml)
    if options.scrub_metadata or options.hash_strings:
        for el in root.iter():
            if el.text and el.text.strip():
                el.text = _sanitize_text(el.text, options=options)
            if el.tail and el.tail.strip():
                el.tail = _sanitize_text(el.tail, options=options)
        _sanitize_xml_attributes(root, options=options, attr_names={"alt", "title", "href"})

    if options.remove_secrets:
        # VML comments/drawings can embed raster images via `<v:imagedata>` that reference
        # `xl/media/**`. Remove them so we don't leave dangling image references after
        # dropping media parts.
        for parent in list(root.iter()):
            for child in list(parent):
                if child.tag.split("}")[-1] == "imagedata":
                    parent.remove(child)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _sanitize_pivot_cache_definition(
    xml: bytes, *, options: SanitizeOptions, sheet_rename_map: dict[str, str] | None = None
) -> bytes:
    root = ET.fromstring(xml)
    if root.tag.split("}")[-1] != "pivotCacheDefinition":
        return xml

    # Pivot caches can embed worksheet names (via cacheSource/worksheetSource) as well as
    # cached item values (sharedItems). Both are common leakage vectors even when cells are
    # redacted.
    if sheet_rename_map:
        for el in root.iter():
            sheet = el.attrib.get("sheet")
            if sheet and sheet in sheet_rename_map:
                el.attrib["sheet"] = sheet_rename_map[sheet]

    if options.scrub_metadata or options.hash_strings:
        salt = _require_hash_salt(options)
        field_idx = 1
        for el in root.iter():
            if el.tag.split("}")[-1] == "cacheField":
                name = el.attrib.get("name")
                if name:
                    if options.hash_strings:
                        el.attrib["name"] = _hash_text(name, salt=salt)
                    else:
                        el.attrib["name"] = f"Field{field_idx}"
                    field_idx += 1
            caption = el.attrib.get("caption")
            if caption:
                el.attrib["caption"] = _sanitize_text(caption, options=options)

    if options.redact_cell_values or options.hash_strings:
        # Cached unique values are stored under `<sharedItems>` / `<groupItems>` and can
        # leak source data. Drop them so the workbook doesn't contain plaintext caches.
        for parent in list(root.iter()):
            for child in list(parent):
                if child.tag.split("}")[-1] in {"sharedItems", "groupItems"}:
                    parent.remove(child)

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _sanitize_pivot_cache_records(xml: bytes, *, options: SanitizeOptions) -> bytes:
    root = ET.fromstring(xml)
    if root.tag.split("}")[-1] != "pivotCacheRecords":
        return xml

    if options.redact_cell_values or options.hash_strings:
        for child in list(root):
            root.remove(child)
        if "count" in root.attrib:
            root.attrib["count"] = "0"

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _sanitize_chart(
    xml: bytes,
    *,
    options: SanitizeOptions,
    sheet_rename_map: dict[str, str] | None = None,
    table_rename_map: dict[str, str] | None = None,
    table_column_rename_map: dict[str, dict[str, str]] | None = None,
) -> bytes:
    root = ET.fromstring(xml)
    # Remove cached series values; these can leak computed data.
    if options.redact_cell_values or options.hash_strings:
        for parent in list(root.iter()):
            for child in list(parent):
                local = child.tag.split("}")[-1]
                if local in {
                    "numCache",
                    "strCache",
                    "multiLvlStrCache",
                    "numLit",
                    "strLit",
                    "multiLvlStrLit",
                    "dateCache",
                }:
                    parent.remove(child)

    # Chart titles / labels can contain PII as DrawingML text.
    if options.scrub_metadata or options.hash_strings:
        _sanitize_xml_text_elements(root, options=options, local_names={"t"})

    if sheet_rename_map or table_rename_map:
        for el in root.iter():
            if el.tag.split("}")[-1] != "f":
                continue
            if not el.text:
                continue
            el.text = _sanitize_formula_text(
                el.text,
                options=options,
                sheet_rename_map=sheet_rename_map or {},
                table_rename_map=table_rename_map or {},
                table_column_rename_map=table_column_rename_map or {},
            )

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _rels_base_dir(rels_part_name: str) -> str:
    """Return the base directory used to resolve Relationship@Target."""

    rels_no_ext = rels_part_name[:-len(".rels")]
    parts = rels_no_ext.split("/")
    if "_rels" in parts:
        idx = parts.index("_rels")
        source_parts = parts[:idx] + parts[idx + 1 :]
    else:
        source_parts = parts
    source_part = "/".join(source_parts)
    base_dir = posixpath.dirname(source_part)
    return f"{base_dir}/" if base_dir else ""


def _resolve_rel_target(rels_part_name: str, target: str) -> str:
    target = target.split("#", 1)[0]
    if target.startswith("/"):
        return target.lstrip("/")
    base_dir = _rels_base_dir(rels_part_name)
    return posixpath.normpath(posixpath.join(base_dir, target))


def _sanitize_relationships(
    xml: bytes,
    *,
    rels_part_name: str,
    removed_parts: set[str],
    options: SanitizeOptions,
) -> bytes:
    root = ET.fromstring(xml)
    if root.tag != qn(NS_REL, "Relationships"):
        return xml

    removed_lower = {p.lower() for p in removed_parts}

    to_remove: list[ET.Element] = []
    for rel in list(root):
        if rel.tag != qn(NS_REL, "Relationship"):
            continue
        target = rel.attrib.get("Target", "")
        target_mode = rel.attrib.get("TargetMode")

        if options.remove_external_links and (target_mode == "External" or _looks_like_external_url(target)):
            rel.attrib["Target"] = "https://redacted.invalid/"
            rel.attrib["TargetMode"] = "External"
            continue

        if not target or target_mode == "External":
            continue

        resolved = _resolve_rel_target(rels_part_name, target)
        # Be best-effort about part name casing. OOXML part names are case-sensitive in ZIPs,
        # but Excel's writers/readers are not always consistent. If the sanitizer removed a
        # part, drop relationships to any case variant of that part name.
        if resolved in removed_parts or resolved.lower() in removed_lower:
            to_remove.append(rel)

    for rel in to_remove:
        root.remove(rel)

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _sanitize_content_types(xml: bytes, *, removed_parts: set[str]) -> bytes:
    root = ET.fromstring(xml)
    if root.tag != qn(NS_CT, "Types"):
        return xml

    removed_lower = {p.lower() for p in removed_parts}

    to_remove: list[ET.Element] = []
    for child in list(root):
        if child.tag != qn(NS_CT, "Override"):
            continue
        part_name = child.attrib.get("PartName", "")
        if not part_name.startswith("/"):
            continue
        if part_name.lstrip("/") in removed_parts or part_name.lstrip("/").lower() in removed_lower:
            to_remove.append(child)

    for child in to_remove:
        root.remove(child)

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _sanitize_workbook(
    xml: bytes, *, options: SanitizeOptions, sheet_rename_map: dict[str, str] | None = None
) -> bytes:
    root = ET.fromstring(xml)
    # `<externalReferences>` has no consistent prefix usage, so match by local name.
    if options.remove_external_links:
        for child in list(root):
            if child.tag.split("}")[-1] == "externalReferences":
                root.remove(child)

    # Defined names can embed sensitive business terms (both the name and the definition).
    if options.scrub_metadata or options.hash_strings:
        for child in list(root):
            if child.tag.split("}")[-1] == "definedNames":
                root.remove(child)
    elif sheet_rename_map:
        # Sheet rename mode must keep workbook-internal references consistent.
        for el in root.iter():
            if el.tag.split("}")[-1] != "definedName":
                continue
            if el.text is None:
                continue
            el.text = _rewrite_formula_sheet_references(el.text, sheet_rename_map=sheet_rename_map)

    # Workbook-level protection/password hashes and user-sharing metadata are common leak
    # vectors in enterprise spreadsheets (usernames, legacy password hashes).
    if options.remove_secrets or options.scrub_metadata or options.hash_strings:
        for child in list(root):
            local = child.tag.split("}")[-1]
            if options.remove_secrets and local in {"workbookProtection"}:
                root.remove(child)
                continue
            if (options.scrub_metadata or options.hash_strings or options.remove_secrets) and local in {
                "fileSharing"
            }:
                root.remove(child)
                continue
            if (options.scrub_metadata or options.hash_strings) and local == "workbookPr":
                if "codeName" in child.attrib:
                    child.attrib["codeName"] = _sanitize_text(child.attrib["codeName"], options=options)

    if sheet_rename_map:
        sheets = None
        for child in list(root):
            if child.tag.split("}")[-1] == "sheets":
                sheets = child
                break
        if sheets is not None:
            for sheet in sheets:
                if sheet.tag.split("}")[-1] != "sheet":
                    continue
                old = sheet.attrib.get("name")
                if not old:
                    continue
                new = sheet_rename_map.get(old)
                if new:
                    sheet.attrib["name"] = new
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _looks_like_external_url(value: str) -> bool:
    v = value.strip().lower()
    return v.startswith(
        (
            "http://",
            "https://",
            "mailto:",
            "ftp://",
            "ftps://",
            "file:",
            "tel:",
            "smb://",
            "\\\\",  # UNC path
            "//",  # network-path reference
        )
    )


def _rewrite_formula_sheet_references(formula: str, *, sheet_rename_map: dict[str, str]) -> str:
    if not sheet_rename_map:
        return formula

    # Replace quoted sheet references first: 'Old Name'!
    for old, new in sheet_rename_map.items():
        if not old:
            continue
        old_escaped = old.replace("'", "''")
        formula = formula.replace(f"'{old_escaped}'!", f"{new}!")
        formula = formula.replace(f"{old}!", f"{new}!")
    return formula


def _rewrite_formula_table_references(formula: str, *, table_rename_map: dict[str, str]) -> str:
    if not table_rename_map:
        return formula

    out = formula
    for old, new in table_rename_map.items():
        if not old:
            continue
        out = re.sub(rf"(?<![A-Za-z0-9_]){re.escape(old)}(?=\[)", new, out)
    return out


def _rewrite_structured_refs_for_table(
    formula: str, *, table_name: str, column_map: dict[str, str]
) -> str:
    if not table_name or not column_map:
        return formula

    # Rewrite only outside of string literals to avoid accidentally mutating user strings.
    out: list[str] = []
    i = 0
    in_string = False
    while i < len(formula):
        ch = formula[i]

        if in_string:
            out.append(ch)
            if ch == '"':
                # Escaped quote inside string: "" -> literal "
                if i + 1 < len(formula) and formula[i + 1] == '"':
                    out.append('"')
                    i += 2
                    continue
                in_string = False
            i += 1
            continue

        if ch == '"':
            in_string = True
            out.append(ch)
            i += 1
            continue

        if formula.startswith(table_name, i) and i + len(table_name) < len(formula) and formula[i + len(table_name)] == "[":
            # Ensure the table name isn't part of a larger identifier.
            if i > 0 and re.match(r"[A-Za-z0-9_]", formula[i - 1]):
                out.append(ch)
                i += 1
                continue

            out.append(table_name)
            i += len(table_name)

            # Parse the structured reference chunk, balancing nested brackets.
            start = i
            depth = 0
            while i < len(formula):
                ch = formula[i]
                if ch == "[":
                    depth += 1
                elif ch == "]":
                    depth -= 1
                    if depth == 0:
                        i += 1
                        break
                i += 1

            chunk = formula[start:i]
            new_chunk = chunk
            # Replace longer names first to avoid partial matches.
            for old, new in sorted(column_map.items(), key=lambda kv: len(kv[0]), reverse=True):
                if not old:
                    continue
                new_chunk = new_chunk.replace(f"[@{old}]", f"[@{new}]")
                new_chunk = new_chunk.replace(f"[{old}]", f"[{new}]")
            out.append(new_chunk)
            continue

        out.append(ch)
        i += 1

    return "".join(out)


def _split_formula_by_string_literals(formula: str) -> list[tuple[str, bool]]:
    """Return [(segment, is_string_literal)] preserving order."""

    out: list[tuple[str, bool]] = []
    i = 0
    while i < len(formula):
        if formula[i] != '"':
            j = i
            while j < len(formula) and formula[j] != '"':
                j += 1
            out.append((formula[i:j], False))
            i = j
            continue

        # String literal (Excel): "..." with doubled quotes escaping.
        j = i + 1
        while j < len(formula):
            if formula[j] == '"':
                if j + 1 < len(formula) and formula[j + 1] == '"':
                    j += 2
                    continue
                j += 1
                break
            j += 1
        out.append((formula[i:j], True))
        i = j

    return out


def _rewrite_unqualified_structured_refs(formula: str, *, column_map: dict[str, str]) -> str:
    """Rewrite `[@Col]` / `[Col]` references outside of string literals."""

    if not column_map:
        return formula

    replacements = sorted(column_map.items(), key=lambda kv: len(kv[0]), reverse=True)
    parts: list[str] = []
    for seg, is_string in _split_formula_by_string_literals(formula):
        if is_string:
            parts.append(seg)
            continue
        for old, new in replacements:
            if not old:
                continue
            seg = seg.replace(f"[@{old}]", f"[@{new}]")
            seg = seg.replace(f"[{old}]", f"[{new}]")
        parts.append(seg)
    return "".join(parts)


_A1_CELL_RE = re.compile(r"^\$?([A-Za-z]+)\$?(\d+)$")


def _a1_col_to_index(col: str) -> int:
    idx = 0
    for ch in col.upper():
        if not ("A" <= ch <= "Z"):
            break
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx


def _a1_cell_to_row_col(cell: str) -> tuple[int, int] | None:
    m = _A1_CELL_RE.match(cell.strip())
    if not m:
        return None
    col_s, row_s = m.group(1), m.group(2)
    return int(row_s), _a1_col_to_index(col_s)


def _a1_range_to_bounds(ref: str) -> tuple[int, int, int, int] | None:
    ref = ref.strip()
    if not ref:
        return None
    if ":" in ref:
        start, end = ref.split(":", 1)
    else:
        start = end = ref
    start_rc = _a1_cell_to_row_col(start)
    end_rc = _a1_cell_to_row_col(end)
    if not start_rc or not end_rc:
        return None
    r1, c1 = start_rc
    r2, c2 = end_rc
    return min(r1, r2), max(r1, r2), min(c1, c2), max(c1, c2)


def _cell_in_bounds(row: int, col: int, bounds: tuple[int, int, int, int]) -> bool:
    r1, r2, c1, c2 = bounds
    return r1 <= row <= r2 and c1 <= col <= c2


def _bounds_intersect(a: tuple[int, int, int, int], b: tuple[int, int, int, int]) -> bool:
    ar1, ar2, ac1, ac2 = a
    br1, br2, bc1, bc2 = b
    return not (ar2 < br1 or ar1 > br2 or ac2 < bc1 or ac1 > bc2)


def _table_context_for_sqref(sqref: str, *, table_contexts: list[TableContext]) -> TableContext | None:
    if not sqref or not table_contexts:
        return None
    for token in sqref.split():
        bounds = _a1_range_to_bounds(token)
        if not bounds:
            continue
        for ctx in table_contexts:
            if _bounds_intersect(
                bounds, (ctx.start_row, ctx.end_row, ctx.start_col, ctx.end_col)
            ):
                return ctx
    return None


def _rewrite_formula_table_column_references(
    formula: str, *, table_column_rename_map: dict[str, dict[str, str]]
) -> str:
    if not table_column_rename_map:
        return formula

    out = formula
    for table_name, column_map in table_column_rename_map.items():
        out = _rewrite_structured_refs_for_table(out, table_name=table_name, column_map=column_map)
    return out


def _sanitize_formula_text(
    formula: str,
    *,
    options: SanitizeOptions,
    sheet_rename_map: dict[str, str],
    table_rename_map: dict[str, str],
    table_column_rename_map: dict[str, dict[str, str]],
) -> str:
    out = formula
    if sheet_rename_map:
        out = _rewrite_formula_sheet_references(out, sheet_rename_map=sheet_rename_map)
    if table_rename_map:
        out = _rewrite_formula_table_references(out, table_rename_map=table_rename_map)
    if table_column_rename_map:
        out = _rewrite_formula_table_column_references(out, table_column_rename_map=table_column_rename_map)

    # If we're hashing, also hash string literals inside formulas; those can leak PII.
    if options.hash_strings:
        salt = _require_hash_salt(options)
        # Excel formula string literals: "..." with doubled quotes for escaping.
        def _repl(match: re.Match[str]) -> str:
            raw = match.group(0)
            inner = raw[1:-1].replace('""', '"')
            hashed = _hash_text(inner, salt=salt)
            return f'"{hashed}"'

        out = re.sub(r'"(?:[^"]|"")*"', _repl, out)
    elif options.redact_cell_values:
        # Redact formula string literals too so PII doesn't survive via formulas.
        out = re.sub(r'"(?:[^"]|"")*"', '"REDACTED"', out)
    return out


def _sanitize_formula_cells_in_worksheet(
    root: ET.Element,
    *,
    options: SanitizeOptions,
    sheet_rename_map: dict[str, str],
    table_rename_map: dict[str, str],
    table_column_rename_map: dict[str, dict[str, str]],
    table_contexts: list[TableContext] | None = None,
) -> None:
    for c in root.iter(qn(NS_MAIN, "c")):
        f_el = c.find(qn(NS_MAIN, "f"))
        if f_el is None or not f_el.text:
            continue
        f_el.text = _sanitize_formula_text(
            f_el.text,
            options=options,
            sheet_rename_map=sheet_rename_map,
            table_rename_map=table_rename_map,
            table_column_rename_map=table_column_rename_map,
        )
        if table_contexts:
            addr = c.attrib.get("r")
            if addr:
                rc = _a1_cell_to_row_col(addr)
                if rc:
                    row, col = rc
                    for ctx in table_contexts:
                        if _cell_in_bounds(row, col, (ctx.start_row, ctx.end_row, ctx.start_col, ctx.end_col)):
                            f_el.text = _rewrite_unqualified_structured_refs(
                                f_el.text or "", column_map=ctx.column_map
                            )
                            break

    if sheet_rename_map:
        for hl in root.iter():
            if hl.tag.split("}")[-1] != "hyperlink":
                continue
            loc = hl.attrib.get("location")
            if not loc or _looks_like_external_url(loc):
                continue
            hl.attrib["location"] = _rewrite_formula_sheet_references(loc, sheet_rename_map=sheet_rename_map)

    # Data validation rules can embed plaintext lists and user-visible prompts/errors.
    if options.redact_cell_values or options.hash_strings or options.scrub_metadata:
        for dv in root.iter():
            if dv.tag.split("}")[-1] != "dataValidation":
                continue

            if options.scrub_metadata or options.hash_strings:
                for attr in ("prompt", "error", "promptTitle", "errorTitle"):
                    if attr in dv.attrib and dv.attrib[attr]:
                        dv.attrib[attr] = _sanitize_text(dv.attrib[attr], options=options)

            ctx = None
            if table_contexts:
                ctx = _table_context_for_sqref(dv.attrib.get("sqref", ""), table_contexts=table_contexts)

            for child in dv:
                local = child.tag.split("}")[-1]
                if local not in {"formula1", "formula2"}:
                    continue
                if not child.text:
                    continue
                child.text = _sanitize_formula_text(
                    child.text,
                    options=options,
                    sheet_rename_map=sheet_rename_map,
                    table_rename_map=table_rename_map,
                    table_column_rename_map=table_column_rename_map,
                )
                if ctx:
                    child.text = _rewrite_unqualified_structured_refs(child.text, column_map=ctx.column_map)

    # Conditional formatting formulas can contain string literals and structured references.
    if options.redact_cell_values or options.hash_strings:
        for cf in root.iter():
            if cf.tag.split("}")[-1] != "conditionalFormatting":
                continue
            ctx = None
            if table_contexts:
                ctx = _table_context_for_sqref(cf.attrib.get("sqref", ""), table_contexts=table_contexts)
            for formula_el in cf.iter():
                if formula_el.tag.split("}")[-1] != "formula":
                    continue
                if not formula_el.text:
                    continue
                formula_el.text = _sanitize_formula_text(
                    formula_el.text,
                    options=options,
                    sheet_rename_map=sheet_rename_map,
                    table_rename_map=table_rename_map,
                    table_column_rename_map=table_column_rename_map,
                )
                if ctx:
                    formula_el.text = _rewrite_unqualified_structured_refs(
                        formula_el.text, column_map=ctx.column_map
                    )

    # AutoFilter criteria can cache plaintext filter values even after cells are redacted.
    if options.redact_cell_values or options.hash_strings:
        for af in root.iter():
            if af.tag.split("}")[-1] != "autoFilter":
                continue
            for child in af.iter():
                if child is af:
                    continue
                if "val" in child.attrib and child.attrib["val"]:
                    child.attrib["val"] = _sanitize_text(child.attrib["val"], options=options)


def scan_xlsx_bytes_for_leaks(
    data: bytes,
    *,
    plaintext_strings: Iterable[str] | None = None,
    scan_patterns: bool = True,
) -> LeakScanResult:
    """Scan an XLSX zip blob for common PII/secret leaks.

    This is intentionally heuristic and should be used as a safety net, not a substitute
    for sanitizing known parts.

    Note: Findings intentionally avoid returning the matched plaintext (only a SHA256),
    so callers can fail CI without leaking secrets into logs.
    """

    findings: list[LeakScanFinding] = []

    # Keep URL matching fairly strict to avoid false positives on unrelated strings.
    email_re = re.compile(r"\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b", re.IGNORECASE)
    url_re = re.compile(r"\b(?:https?|file|ftp|ftps|smb)://[^\s\"'<>]+", re.IGNORECASE)
    # UNC paths start with `\\server\share`; treat these as high-risk external URLs too.
    unc_re = re.compile(r"\\\\[^\s\"'<>]+")
    aws_key_re = re.compile(r"\b(?:AKIA|ASIA)[0-9A-Z]{16}\b")
    jwt_re = re.compile(r"\beyJ[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{10,}\b")

    url_allowlist = {
        "schemas.openxmlformats.org",
        "schemas.microsoft.com",
        "www.w3.org",
        "purl.org",
        "redacted.invalid",
    }

    def _sha(text: str) -> str:
        return hashlib.sha256(text.encode("utf-8")).hexdigest()

    def _record(part: str, kind: str, match: str) -> None:
        findings.append(LeakScanFinding(part_name=part, kind=kind, match_sha256=_sha(match)))

    def _xml_escape(s: str) -> str:
        return (
            s.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;")
            .replace("'", "&apos;")
        )

    needles: list[str] = []
    if plaintext_strings:
        for s in plaintext_strings:
            if not s:
                continue
            if len(s) < 4:
                # Avoid a flood of false positives on short substrings.
                continue
            needles.append(s)
            needles.append(_xml_escape(s))

    input_buf = io.BytesIO(data)
    with zipfile.ZipFile(input_buf, "r") as zin:
        for info in zin.infolist():
            if info.is_dir():
                continue
            part = info.filename
            raw = zin.read(part)
            try:
                text = raw.decode("utf-8")
            except UnicodeDecodeError:
                text = raw.decode("utf-8", errors="ignore")

            for s in needles:
                if s in text:
                    _record(part, "plaintext", s)

            if not scan_patterns:
                continue

            for m in email_re.findall(text):
                _record(part, "email", m)

            for m in aws_key_re.findall(text):
                _record(part, "aws_key", m)

            for m in jwt_re.findall(text):
                _record(part, "jwt", m)

            for m in url_re.findall(text):
                # Filter out standard OOXML namespace URLs.
                try:
                    from urllib.parse import urlparse

                    parsed = urlparse(m)
                    host = (parsed.hostname or "").lower()
                except Exception:
                    host = ""
                if host in url_allowlist or any(host.endswith("." + d) for d in url_allowlist):
                    continue
                _record(part, "url", m)

            for m in unc_re.findall(text):
                _record(part, "url", m)

    return LeakScanResult(findings=findings)


def sanitize_xlsx_bytes(data: bytes, *, options: SanitizeOptions) -> tuple[bytes, SanitizeSummary]:
    """Return a sanitized XLSX zip blob.

    The sanitizer is intentionally conservative:
    - it never writes plaintext cell values into the output unless explicitly configured
    - it removes common secret-bearing parts (connections/customXml) by default
    - it scrubs external relationship targets so links/URLs don't leak
    """

    input_buf = io.BytesIO(data)
    with zipfile.ZipFile(input_buf, "r") as zin:
        def _normalize_zip_entry_name(name: str) -> str:
            # ZIP entry names in valid XLSX/XLSM packages should not start with `/`, but some
            # producers emit them. Be tolerant by normalizing entry names for all downstream
            # bookkeeping and rewrites.
            return name.replace("\\", "/").lstrip("/")

        # Normalize ZIP entry names (trim leading `/`, normalize `\\`) but keep a mapping to the
        # original archive entry name for reads. This avoids failures when workbooks store core
        # parts like `/xl/workbook.xml` instead of `xl/workbook.xml`.
        #
        # If the archive contains duplicate entries that collapse to the same normalized part name
        # (e.g. both `xl/workbook.xml` and `/xl/workbook.xml`), prefer the canonical non-slashed
        # entry name when present.
        names: list[str] = []
        original_by_name: dict[str, str] = {}
        for info in zin.infolist():
            if info.is_dir():
                continue
            raw = info.filename
            name = _normalize_zip_entry_name(raw)
            if name not in original_by_name:
                names.append(name)
                original_by_name[name] = raw
            else:
                prev = original_by_name[name]
                if prev.startswith("/") and not raw.startswith("/"):
                    original_by_name[name] = raw

        removed_parts: set[str] = set()
        if options.remove_external_links:
            removed_parts |= {n for n in names if n.startswith("xl/externalLinks/")}
        if options.remove_secrets:
            removed_parts |= {n for n in names if n == "xl/connections.xml"}
            removed_parts |= {n for n in names if n.startswith("xl/queryTables/")}
            removed_parts |= {n for n in names if n.startswith("customXml/")}
            removed_parts |= {n for n in names if n.startswith("xl/customXml/")}
            removed_parts |= {n for n in names if n.startswith("xl/model/")}
            removed_parts |= {n for n in names if n.startswith("xl/printerSettings/")}
            removed_parts |= {n for n in names if n.startswith("xl/revisions/")}
            removed_parts |= {n for n in names if n.startswith("xl/webExtensions/")}
            removed_parts |= {n for n in names if n.startswith("xl/controls/")}
            removed_parts |= {n for n in names if n.startswith("xl/ctrlProps/")}
            removed_parts |= {n for n in names if n.startswith("xl/media/")}
            removed_parts |= {n for n in names if n.startswith("xl/embeddings/")}
            removed_parts |= {n for n in names if n == "xl/vbaProject.bin"}
            removed_parts |= {n for n in names if n == "xl/vbaProjectSignature.bin"}
            removed_parts |= {n for n in names if n.startswith("xl/activeX/")}
            removed_parts |= {n for n in names if n.startswith("customUI/")}
            removed_parts |= {n for n in names if n.startswith("docProps/thumbnail")}

            # Excel in-cell images parts (cellImages*.xml) can reference raster parts in
            # `xl/media/**`. When secrets are removed we drop media, so remove cellImages*.xml
            # parts too to avoid leaving dangling `r:embed` references.
            cellimages_parts = {n for n in names if _is_cellimages_xml_part(n)}
            removed_parts |= cellimages_parts
            for part in cellimages_parts:
                rels_lower = posixpath.join(
                    posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels"
                ).casefold()
                removed_parts |= {n for n in names if n.casefold() == rels_lower}

        if options.scrub_metadata or options.remove_secrets:
            removed_parts |= {n for n in names if n == "docProps/custom.xml"}

        sheet_rename_map: dict[str, str] = {}
        if options.rename_sheets:
            try:
                wb_part = original_by_name.get("xl/workbook.xml")
                if wb_part is None:
                    raise KeyError("missing xl/workbook.xml")
                wb_root = ET.fromstring(zin.read(wb_part))
                sheets = None
                for child in wb_root:
                    if child.tag.split("}")[-1] == "sheets":
                        sheets = child
                        break
                if sheets is not None:
                    idx = 1
                    for sheet in sheets:
                        if sheet.tag.split("}")[-1] != "sheet":
                            continue
                        name = sheet.attrib.get("name")
                        if not name:
                            continue
                        sheet_rename_map[name] = f"Sheet{idx}"
                        idx += 1
            except Exception:
                sheet_rename_map = {}

        table_rename_map: dict[str, str] = {}
        table_column_rename_map: dict[str, dict[str, str]] = {}
        table_part_context: dict[str, TableContext] = {}
        if options.scrub_metadata or options.hash_strings:
            table_parts = sorted([n for n in names if n.startswith("xl/tables/") and n.endswith(".xml")])
            idx = 1
            for part in table_parts:
                try:
                    t_root = ET.fromstring(zin.read(original_by_name[part]))
                except Exception:
                    continue
                if t_root.tag.split("}")[-1] != "table":
                    continue
                old = t_root.attrib.get("displayName") or t_root.attrib.get("name")
                if not old:
                    continue
                if options.hash_strings:
                    new_table = _hash_text(old, salt=_require_hash_salt(options))
                else:
                    new_table = f"Table{idx}"
                table_rename_map[old] = new_table

                # Table column names are duplicated metadata and should be scrubbed too.
                col_map: dict[str, str] = {}
                col_idx = 1
                for col in t_root.iter():
                    if col.tag.split("}")[-1] != "tableColumn":
                        continue
                    col_name = col.attrib.get("name")
                    if not col_name:
                        continue
                    if options.hash_strings:
                        new_col = _hash_text(col_name, salt=_require_hash_salt(options))
                    else:
                        new_col = f"Column{col_idx}"
                    col_map[col_name] = new_col
                    col_idx += 1
                if col_map:
                    table_column_rename_map[new_table] = col_map
                    ref = t_root.attrib.get("ref") or ""
                    bounds = _a1_range_to_bounds(ref)
                    if bounds:
                        r1, r2, c1, c2 = bounds
                        table_part_context[part] = TableContext(
                            start_row=r1,
                            end_row=r2,
                            start_col=c1,
                            end_col=c2,
                            column_map=col_map,
                        )

                idx += 1

        rewritten: list[str] = []

        output_buf = io.BytesIO()
        with zipfile.ZipFile(output_buf, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for name in names:
                if name in removed_parts:
                    continue

                raw = zin.read(original_by_name[name])
                new = raw

                try:
                    if name == "[Content_Types].xml":
                        new = _sanitize_content_types(raw, removed_parts=removed_parts)
                        rewritten.append(name)
                    elif name.endswith(".rels"):
                        new = _sanitize_relationships(
                            raw,
                            rels_part_name=name,
                            removed_parts=removed_parts,
                            options=options,
                        )
                        rewritten.append(name)
                    elif name == "xl/workbook.xml":
                        new = _sanitize_workbook(
                            raw, options=options, sheet_rename_map=sheet_rename_map or None
                        )
                        if new != raw:
                            rewritten.append(name)
                    elif (
                        (
                            name.startswith("xl/worksheets/")
                            or name.startswith("xl/macrosheets/")
                            or name.startswith("xl/dialogsheets/")
                            or name.startswith("xl/chartsheets/")
                        )
                        and name.endswith(".xml")
                        and (
                            options.redact_cell_values
                            or options.hash_strings
                            or options.scrub_metadata
                            or options.remove_external_links
                            or options.remove_secrets
                            or options.rename_sheets
                        )
                    ):
                        new = _sanitize_worksheet(raw, options=options)
                        if (
                            options.redact_cell_values
                            or options.hash_strings
                            or sheet_rename_map
                            or table_rename_map
                            or table_column_rename_map
                        ):
                            # Rewrite formula sheet/table references and scrub string literals.
                            ws_root = ET.fromstring(new)
                            sheet_table_contexts: list[TableContext] | None = None
                            if table_part_context:
                                rels_name = posixpath.join(
                                    posixpath.dirname(name),
                                    "_rels",
                                    posixpath.basename(name) + ".rels",
                                )
                                if rels_name in names and rels_name not in removed_parts:
                                    try:
                                        rels_root = ET.fromstring(zin.read(original_by_name[rels_name]))
                                        table_rids: dict[str, str] = {}
                                        for rel in list(rels_root):
                                            if rel.tag != qn(NS_REL, "Relationship"):
                                                continue
                                            if not rel.attrib.get("Type", "").endswith("/table"):
                                                continue
                                            rid = rel.attrib.get("Id")
                                            target = rel.attrib.get("Target")
                                            if not rid or not target:
                                                continue
                                            table_rids[rid] = _resolve_rel_target(rels_name, target)
                                        if table_rids:
                                            sheet_table_contexts = []
                                            for tp in ws_root.iter():
                                                if tp.tag.split("}")[-1] != "tablePart":
                                                    continue
                                                rid = None
                                                for k, v in tp.attrib.items():
                                                    if k.split("}")[-1] == "id":
                                                        rid = v
                                                        break
                                                if not rid:
                                                    continue
                                                table_part = table_rids.get(rid)
                                                if not table_part:
                                                    continue
                                                ctx = table_part_context.get(table_part)
                                                if ctx:
                                                    sheet_table_contexts.append(ctx)
                                    except Exception:
                                        sheet_table_contexts = None
                            _sanitize_formula_cells_in_worksheet(
                                ws_root,
                                options=options,
                                sheet_rename_map=sheet_rename_map,
                                table_rename_map=table_rename_map,
                                table_column_rename_map=table_column_rename_map,
                                table_contexts=sheet_table_contexts,
                            )
                            new = ET.tostring(ws_root, encoding="utf-8", xml_declaration=True)
                        rewritten.append(name)
                    elif name == "xl/sharedStrings.xml" and (options.redact_cell_values or options.hash_strings):
                        new = _sanitize_shared_strings(raw, options=options)
                        rewritten.append(name)
                    elif name == "docProps/core.xml" and (options.scrub_metadata or options.hash_strings):
                        new = _sanitize_core_properties(raw, options=options)
                        rewritten.append(name)
                    elif name == "docProps/app.xml" and (
                        options.scrub_metadata or options.hash_strings or options.rename_sheets
                    ):
                        new = _sanitize_app_properties(
                            raw, options=options, sheet_rename_map=sheet_rename_map or None
                        )
                        rewritten.append(name)
                    elif name.startswith("xl/comments") and name.endswith(".xml") and (
                        options.scrub_metadata or options.hash_strings
                    ):
                        new = _sanitize_comments(raw, options=options)
                        rewritten.append(name)
                    elif (
                        name.startswith("xl/threadedComments/") or name.startswith("xl/persons/")
                    ) and name.endswith(".xml") and (options.scrub_metadata or options.hash_strings):
                        new = _sanitize_threaded_comments(raw, options=options)
                        rewritten.append(name)
                    elif name.startswith("xl/pivotCache/") and name.endswith(".xml") and (
                        options.redact_cell_values or options.hash_strings or options.scrub_metadata or options.rename_sheets
                    ):
                        if "pivotCacheDefinition" in posixpath.basename(name):
                            new = _sanitize_pivot_cache_definition(
                                raw,
                                options=options,
                                sheet_rename_map=sheet_rename_map or None,
                            )
                        elif "pivotCacheRecords" in posixpath.basename(name):
                            new = _sanitize_pivot_cache_records(raw, options=options)
                        rewritten.append(name)
                    elif name.startswith("xl/tables/") and name.endswith(".xml") and (
                        options.scrub_metadata or options.hash_strings
                    ):
                        new = _sanitize_table(
                            raw,
                            options=options,
                            table_rename_map=table_rename_map,
                            table_column_rename_map=table_column_rename_map,
                            sheet_rename_map=sheet_rename_map or None,
                        )
                        rewritten.append(name)
                    elif _is_cellimages_xml_part(name) and (options.scrub_metadata or options.hash_strings):
                        new = _sanitize_cell_images(raw, options=options)
                        rewritten.append(name)
                    elif name.startswith("xl/drawings/") and name.endswith(".xml") and (
                        options.scrub_metadata or options.hash_strings or options.remove_secrets
                    ):
                        new = _sanitize_drawing(raw, options=options)
                        rewritten.append(name)
                    elif name.startswith("xl/drawings/") and name.endswith(".vml") and (
                        options.scrub_metadata or options.hash_strings or options.remove_secrets
                    ):
                        new = _sanitize_vml_drawing(raw, options=options)
                        rewritten.append(name)
                    elif name.startswith("xl/charts/") and name.endswith(".xml") and (
                        options.redact_cell_values
                        or options.hash_strings
                        or options.scrub_metadata
                        or options.rename_sheets
                    ):
                        new = _sanitize_chart(
                            raw,
                            options=options,
                            sheet_rename_map=sheet_rename_map or None,
                            table_rename_map=table_rename_map or None,
                            table_column_rename_map=table_column_rename_map or None,
                        )
                        rewritten.append(name)
                except ET.ParseError:
                    # If a part isn't well-formed XML, leave it untouched (we still might remove it above).
                    new = raw

                # Deterministic ZIP output: avoid embedding current timestamps in sanitized
                # workbooks (which both adds noise to git diffs and can leak ingest time).
                info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
                info.compress_type = zipfile.ZIP_DEFLATED
                info.create_system = 0
                info.external_attr = 0
                zout.writestr(info, new)

    removed_list = sorted(removed_parts)
    rewritten_list = sorted(set(rewritten))
    return output_buf.getvalue(), SanitizeSummary(removed_parts=removed_list, rewritten_parts=rewritten_list)
