import { splitSheetQualifier } from "../../../../packages/search/index.js";
import { parseA1Range, type RangeAddress } from "../spreadsheet/a1.js";

/**
 * Parses an A1 range reference that may be prefixed with a sheet qualifier.
 *
 * Formula highlighting/tokenization treats `Sheet1!A1` (and quoted sheet names like
 * `'My Sheet'!A1:B2`) as a single reference token. Most range-preview logic only
 * needs the A1 coordinates, so we strip the sheet prefix and parse the remaining
 * A1 range.
 */
export function parseSheetQualifiedA1Range(text: string): RangeAddress | null {
  const { ref } = splitSheetQualifier(text);
  return parseA1Range(ref);
}

