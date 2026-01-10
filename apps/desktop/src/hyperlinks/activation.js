/**
 * Excel-like activation gesture: Ctrl/Cmd + click.
 *
 * @param {{ button?: number, metaKey?: boolean, ctrlKey?: boolean } | null | undefined} event
 * @returns {boolean}
 */
export function isHyperlinkActivation(event) {
  if (!event) return false;
  const buttonOk = event.button == null || event.button === 0;
  const modifier = Boolean(event.metaKey || event.ctrlKey);
  return buttonOk && modifier;
}

