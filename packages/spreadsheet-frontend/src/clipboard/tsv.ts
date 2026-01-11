export function serializeGridToTsv(grid: string[][]): string {
  return grid.map((row) => row.join("\t")).join("\n");
}

export function parseTsvToGrid(tsv: string): string[][] {
  const normalized = tsv.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  const lines = normalized.split("\n");

  // Drop the final empty record when the clipboard payload ends with a newline.
  if (lines.length > 1 && lines.at(-1) === "") lines.pop();

  return lines.map((line) => line.split("\t"));
}

