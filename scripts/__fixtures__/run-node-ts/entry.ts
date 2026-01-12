import { valueFromBar } from "../resolve-ts-imports/foo.ts";
import { valueFromBarExtensionless } from "../resolve-ts-imports/foo-extensionless.ts";
import { valueFromDirImport } from "../resolve-ts-imports/foo-dir-import.ts";

// Prints a stable sentinel value used by `scripts/run-node-ts.test.js`.
console.log([valueFromBar(), valueFromBarExtensionless(), valueFromDirImport()].join(","));

