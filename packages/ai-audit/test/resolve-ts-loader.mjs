import { readFile } from "node:fs/promises";

import ts from "typescript";

export async function resolve(specifier, context, defaultResolve) {
  if (specifier.endsWith(".js") && (specifier.startsWith("./") || specifier.startsWith("../") || specifier.startsWith("/"))) {
    try {
      return await defaultResolve(specifier, context, defaultResolve);
    } catch {
      return defaultResolve(specifier.slice(0, -3) + ".ts", context, defaultResolve);
    }
  }

  return defaultResolve(specifier, context, defaultResolve);
}

export async function load(url, context, defaultLoad) {
  if (url.endsWith(".ts")) {
    const source = await readFile(new URL(url), "utf8");
    const result = ts.transpileModule(source, {
      compilerOptions: {
        module: ts.ModuleKind.ESNext,
        target: ts.ScriptTarget.ES2022
      }
    });

    return {
      format: "module",
      source: result.outputText,
      shortCircuit: true
    };
  }

  return defaultLoad(url, context, defaultLoad);
}
