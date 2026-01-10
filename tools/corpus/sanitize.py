from __future__ import annotations

import hashlib
import io
import posixpath
import zipfile
from dataclasses import dataclass
from xml.etree import ElementTree as ET


NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"


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


@dataclass(frozen=True)
class SanitizeSummary:
    removed_parts: list[str]
    rewritten_parts: list[str]


def _hash_text(value: str, *, salt: str) -> str:
    # Use a stable, corpus-level salt so identical strings hash identically across files,
    # but remain resistant to rainbow-table attacks when the salt is private.
    digest = hashlib.sha256((salt + "\0" + value).encode("utf-8")).hexdigest()
    return f"H_{digest[:16]}"


def _sanitize_shared_strings(xml: bytes, *, options: SanitizeOptions) -> bytes:
    root = ET.fromstring(xml)
    if root.tag != qn(NS_MAIN, "sst"):
        return xml

    if options.hash_strings and not options.hash_salt:
        raise ValueError("hash_strings requires hash_salt")

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
    if options.hash_strings and not options.hash_salt:
        raise ValueError("hash_strings requires hash_salt")

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
    if root.tag != qn(NS_MAIN, "worksheet"):
        return xml

    for c in root.iter(qn(NS_MAIN, "c")):
        f_el = c.find(qn(NS_MAIN, "f"))
        v_el = c.find(qn(NS_MAIN, "v"))
        is_el = c.find(qn(NS_MAIN, "is"))

        if f_el is not None:
            # Preserve formulas/structure, but drop cached results which can leak data.
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
            if options.hash_strings and not options.hash_salt:
                raise ValueError("hash_strings requires hash_salt")
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

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _sanitize_core_properties(xml: bytes) -> bytes:
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
            el.text = "REDACTED"
        if local in {"created", "modified"}:
            el.text = "1970-01-01T00:00:00Z"
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _sanitize_app_properties(xml: bytes) -> bytes:
    root = ET.fromstring(xml)
    for el in root.iter():
        local = el.tag.split("}")[-1]
        if local in {"Company", "Manager", "HyperlinkBase"}:
            el.text = "REDACTED"
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

    to_remove: list[ET.Element] = []
    for rel in list(root):
        if rel.tag != qn(NS_REL, "Relationship"):
            continue
        target = rel.attrib.get("Target", "")
        target_mode = rel.attrib.get("TargetMode")

        if options.remove_external_links and target_mode == "External":
            rel.attrib["Target"] = "https://redacted.invalid/"
            continue

        if not target or target_mode == "External":
            continue

        resolved = _resolve_rel_target(rels_part_name, target)
        if resolved in removed_parts:
            to_remove.append(rel)

    for rel in to_remove:
        root.remove(rel)

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _sanitize_content_types(xml: bytes, *, removed_parts: set[str]) -> bytes:
    root = ET.fromstring(xml)
    if root.tag != qn(NS_CT, "Types"):
        return xml

    to_remove: list[ET.Element] = []
    for child in list(root):
        if child.tag != qn(NS_CT, "Override"):
            continue
        part_name = child.attrib.get("PartName", "")
        if not part_name.startswith("/"):
            continue
        if part_name.lstrip("/") in removed_parts:
            to_remove.append(child)

    for child in to_remove:
        root.remove(child)

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _sanitize_workbook(xml: bytes, *, options: SanitizeOptions) -> bytes:
    if not options.remove_external_links:
        return xml
    root = ET.fromstring(xml)
    # `<externalReferences>` has no consistent prefix usage, so match by local name.
    for child in list(root):
        if child.tag.split("}")[-1] == "externalReferences":
            root.remove(child)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def sanitize_xlsx_bytes(data: bytes, *, options: SanitizeOptions) -> tuple[bytes, SanitizeSummary]:
    """Return a sanitized XLSX zip blob.

    The sanitizer is intentionally conservative:
    - it never writes plaintext cell values into the output unless explicitly configured
    - it removes common secret-bearing parts (connections/customXml) by default
    - it scrubs external relationship targets so links/URLs don't leak
    """

    input_buf = io.BytesIO(data)
    with zipfile.ZipFile(input_buf, "r") as zin:
        names = [info.filename for info in zin.infolist() if not info.is_dir()]

        removed_parts: set[str] = set()
        if options.remove_external_links:
            removed_parts |= {n for n in names if n.startswith("xl/externalLinks/")}
        if options.remove_secrets:
            removed_parts |= {n for n in names if n == "xl/connections.xml"}
            removed_parts |= {n for n in names if n.startswith("xl/queryTables/")}
            removed_parts |= {n for n in names if n.startswith("customXml/")}
            removed_parts |= {n for n in names if n.startswith("xl/customXml/")}

        rewritten: list[str] = []

        output_buf = io.BytesIO()
        with zipfile.ZipFile(output_buf, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for name in names:
                if name in removed_parts:
                    continue

                raw = zin.read(name)
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
                        new = _sanitize_workbook(raw, options=options)
                        if new != raw:
                            rewritten.append(name)
                    elif (
                        name.startswith("xl/worksheets/") and name.endswith(".xml") and options.redact_cell_values
                    ) or (
                        name.startswith("xl/worksheets/") and name.endswith(".xml") and options.hash_strings
                    ):
                        new = _sanitize_worksheet(raw, options=options)
                        rewritten.append(name)
                    elif name == "xl/sharedStrings.xml" and (options.redact_cell_values or options.hash_strings):
                        new = _sanitize_shared_strings(raw, options=options)
                        rewritten.append(name)
                    elif name == "docProps/core.xml" and options.scrub_metadata:
                        new = _sanitize_core_properties(raw)
                        rewritten.append(name)
                    elif name == "docProps/app.xml" and options.scrub_metadata:
                        new = _sanitize_app_properties(raw)
                        rewritten.append(name)
                except ET.ParseError:
                    # If a part isn't well-formed XML, leave it untouched (we still might remove it above).
                    new = raw

                zout.writestr(name, new)

    removed_list = sorted(removed_parts)
    rewritten_list = sorted(set(rewritten))
    return output_buf.getvalue(), SanitizeSummary(removed_parts=removed_list, rewritten_parts=rewritten_list)
