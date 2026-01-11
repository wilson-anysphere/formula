/**
 * @param {import("@formula/extension-api").CellValue[][] | undefined | null} values
 */
function sumValues(values) {
  let sum = 0;
  if (!Array.isArray(values)) return sum;
  for (const row of values) {
    if (!Array.isArray(row)) continue;
    for (const value of row) {
      if (typeof value === "number" && Number.isFinite(value)) {
        sum += value;
      }
    }
  }
  return sum;
}

module.exports = { sumValues };

