import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

/**
 * Returns the ribbon schema *content* (button/menu item definitions) as a single string.
 *
 * The ribbon schema was historically defined inline in `src/ribbon/ribbonSchema.ts`, but is
 * now split across per-tab modules under `src/ribbon/schema/*.ts`. Many Node-based
 * regression tests validate command ids / test ids by grepping the schema source; this
 * helper keeps those tests resilient across that refactor.
 *
 * @param {string | string[] | undefined} tabFiles Optional subset of schema modules to read
 * (ex: `"homeTab.ts"`). If omitted, reads all `*.ts` files in `src/ribbon/schema/`.
 */
export function readRibbonSchemaSource(tabFiles) {
  const schemaDir = path.join(__dirname, "..", "src", "ribbon", "schema");
  if (fs.existsSync(schemaDir)) {
    const files = tabFiles
      ? Array.isArray(tabFiles)
        ? tabFiles
        : [tabFiles]
      : fs
          .readdirSync(schemaDir, { withFileTypes: true })
          .filter((entry) => entry.isFile() && entry.name.endsWith(".ts"))
          .map((entry) => entry.name)
          .sort();
    return files.map((file) => fs.readFileSync(path.join(schemaDir, file), "utf8")).join("\n");
  }

  // Back-compat: older versions kept all tab definitions in ribbonSchema.ts.
  const schemaPath = path.join(__dirname, "..", "src", "ribbon", "ribbonSchema.ts");
  return fs.readFileSync(schemaPath, "utf8");
}
