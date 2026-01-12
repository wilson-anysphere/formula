import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("AI chat React panels avoid inline styles (use ai-chat.css classes)", () => {
  const panelDir = path.join(__dirname, "..", "src", "panels", "ai-chat");
  const cssPath = path.join(__dirname, "..", "src", "styles", "ai-chat.css");
  const mainPath = path.join(__dirname, "..", "src", "main.ts");

  const sources = {
    container: fs.readFileSync(path.join(panelDir, "AIChatPanelContainer.tsx"), "utf8"),
    panel: fs.readFileSync(path.join(panelDir, "AIChatPanel.tsx"), "utf8"),
    modal: fs.readFileSync(path.join(panelDir, "ApprovalModal.tsx"), "utf8"),
  };

  for (const [name, content] of Object.entries(sources)) {
    assert.equal(
      /\bstyle\s*=/.test(content),
      false,
      `${name} should not use inline styles; use src/styles/ai-chat.css classes instead`,
    );
  }

  // Sanity-check that the shared CSS classes are present in the JSX.
  for (const className of ["ai-chat-runtime", "ai-chat-panel", "ai-chat-approval-modal"]) {
    assert.ok(
      sources.container.includes(className) || sources.panel.includes(className) || sources.modal.includes(className),
      `Expected AI chat components to reference the ${className} CSS class`,
    );
  }

  assert.equal(fs.existsSync(cssPath), true, "Expected apps/desktop/src/styles/ai-chat.css to exist");
  const css = fs.readFileSync(cssPath, "utf8");
  for (const selector of [".ai-chat-runtime", ".ai-chat-panel", ".ai-chat-approval-modal"]) {
    assert.ok(css.includes(selector), `Expected ai-chat.css to define ${selector}`);
  }

  const mainSrc = fs.readFileSync(mainPath, "utf8");
  assert.match(
    mainSrc,
    /import\s+["'][^"']*styles\/ai-chat\.css["']/,
    "apps/desktop/src/main.ts should import src/styles/ai-chat.css so the AI chat UI is styled in production builds",
  );
});

