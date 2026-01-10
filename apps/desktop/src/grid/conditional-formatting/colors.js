export function argbToCss(argb) {
  if (!argb || typeof argb !== "string" || argb.length !== 8) return undefined;
  const a = parseInt(argb.slice(0, 2), 16);
  const r = parseInt(argb.slice(2, 4), 16);
  const g = parseInt(argb.slice(4, 6), 16);
  const b = parseInt(argb.slice(6, 8), 16);
  const alpha = (a / 255).toFixed(3).replace(/0+$/, "").replace(/\.$/, "");
  return `rgba(${r},${g},${b},${alpha})`;
}
