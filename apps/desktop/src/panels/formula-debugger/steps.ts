import type { DebugStep, TraceNode } from "./types.ts";

export function buildStepTree(formula: string, trace: TraceNode, id: string = "0"): DebugStep {
  const children = (trace.children ?? []).map((child, idx) =>
    buildStepTree(formula, child, `${id}.${idx}`),
  );
  const text = formula.slice(trace.span.start, trace.span.end);
  return {
    id,
    text,
    span: trace.span,
    value: trace.value,
    reference: trace.reference,
    kind: trace.kind,
    children,
  };
}

export class DebuggerState {
  collapsed = new Set<string>();

  toggle(stepId: string): void {
    if (this.collapsed.has(stepId)) {
      this.collapsed.delete(stepId);
    } else {
      this.collapsed.add(stepId);
    }
  }

  isCollapsed(stepId: string): boolean {
    return this.collapsed.has(stepId);
  }
}

export function flattenVisibleSteps(root: DebugStep, collapsed: Set<string>): DebugStep[] {
  const out: DebugStep[] = [];
  function walk(step: DebugStep) {
    out.push(step);
    if (collapsed.has(step.id)) return;
    for (const child of step.children) walk(child);
  }
  walk(root);
  return out;
}

