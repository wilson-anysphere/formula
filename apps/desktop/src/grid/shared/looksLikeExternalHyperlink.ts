export function looksLikeExternalHyperlink(text: string): boolean {
  const trimmed = text.trim();
  if (!trimmed) return false;
  // Avoid interpreting arbitrary "foo:bar" values as URLs; require either a
  // scheme separator (`://`) or a `mailto:` prefix.
  if (/^mailto:/i.test(trimmed)) return true;
  return /^[a-zA-Z][a-zA-Z0-9+.-]*:\/\//.test(trimmed);
}

