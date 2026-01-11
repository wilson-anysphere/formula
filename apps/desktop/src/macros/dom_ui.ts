import { DefaultMacroSecurityController } from "./security";
import type { MacroBackend, MacroCellUpdate } from "./types";
import { MacroRunner } from "./runner";

export interface MacroRunnerRenderOptions {
  onApplyUpdates?: (updates: MacroCellUpdate[]) => void | Promise<void>;
}

/**
 * Minimal "macro runner" UI using DOM APIs (framework-agnostic).
 *
 * Production builds can wrap this in React or another UI layer; this file exists
 * to encode the core UX requirements in executable code:
 * - Select a macro
 * - Run it
 * - Show output and errors
 * - Prompt for macro enablement (and permissions) when needed
 */
export async function renderMacroRunner(
  container: HTMLElement,
  backend: MacroBackend,
  workbookId: string,
  opts: MacroRunnerRenderOptions = {}
): Promise<void> {
  const security = new DefaultMacroSecurityController();
  const runner = new MacroRunner(backend, security);
  const macros = await runner.list(workbookId);

  container.innerHTML = "";

  const header = document.createElement("div");
  header.textContent = "Macros";

  const select = document.createElement("select");
  for (const macro of macros) {
    const opt = document.createElement("option");
    opt.value = macro.id;
    opt.textContent = macro.name;
    select.appendChild(opt);
  }

  const runButton = document.createElement("button");
  runButton.textContent = "Run";

  const output = document.createElement("pre");
  output.style.whiteSpace = "pre-wrap";

  runButton.onclick = async () => {
    output.textContent = "";
    runButton.disabled = true;
    try {
      const macroId = select.value;
      const result = await runner.run({ workbookId, macroId, timeoutMs: 5_000 });
      if (result.output.length) {
        output.textContent += result.output.join("\n") + "\n";
      }
      if (result.updates && result.updates.length) {
        if (opts.onApplyUpdates) {
          try {
            await opts.onApplyUpdates(result.updates);
            output.textContent += `Applied ${result.updates.length} updates.\n`;
          } catch (err) {
            output.textContent += `Error applying updates: ${String(err)}\n`;
          }
        } else {
          output.textContent += `Macro returned ${result.updates.length} updates (not applied).\n`;
        }
      }
      if (!result.ok) {
        output.textContent += `Error: ${result.error?.message ?? "Unknown error"}\n`;
      }
    } catch (err) {
      output.textContent += `Error: ${String(err)}\n`;
    } finally {
      runButton.disabled = false;
    }
  };

  container.appendChild(header);
  container.appendChild(select);
  container.appendChild(runButton);
  container.appendChild(output);
}
