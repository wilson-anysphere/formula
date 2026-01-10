import React, { useMemo, useState } from "react";

import { t } from "../../i18n/index.js";

import type {
  ObjectiveKind,
  Relation,
  SolveMethod,
  SolverConfig,
  SolverConstraintSpec,
  SolverVariableSpec,
} from "./types";

type Props = {
  initial?: Partial<SolverConfig>;
  onCancel: () => void;
  onRun: (config: SolverConfig) => void;
};

const DEFAULT_VARIABLE: SolverVariableSpec = {
  ref: "Sheet1!A1",
  type: "continuous",
};

const DEFAULT_CONSTRAINT: SolverConstraintSpec = {
  ref: "Sheet1!B1",
  relation: "<=",
  rhs: 0,
};

export function SolverDialog({ initial, onCancel, onRun }: Props) {
  const [method, setMethod] = useState<SolveMethod>(initial?.method ?? "grg");
  const [objectiveRef, setObjectiveRef] = useState<string>(
    initial?.objectiveRef ?? "Sheet1!C1",
  );
  const [objectiveKind, setObjectiveKind] = useState<ObjectiveKind>(
    initial?.objectiveKind ?? "minimize",
  );
  const [targetValue, setTargetValue] = useState<number>(
    initial?.targetValue ?? 0,
  );
  const [targetTolerance, setTargetTolerance] = useState<number>(
    initial?.targetTolerance ?? 1e-6,
  );

  const [variables, setVariables] = useState<SolverVariableSpec[]>(
    initial?.variables?.length ? initial.variables : [DEFAULT_VARIABLE],
  );
  const [constraints, setConstraints] = useState<SolverConstraintSpec[]>(
    initial?.constraints?.length ? initial.constraints : [],
  );

  const canTarget = objectiveKind === "target";

  const config: SolverConfig = useMemo(
    () => ({
      method,
      objectiveRef,
      objectiveKind,
      targetValue: canTarget ? targetValue : undefined,
      targetTolerance: canTarget ? targetTolerance : undefined,
      variables,
      constraints,
    }),
    [
      method,
      objectiveRef,
      objectiveKind,
      canTarget,
      targetValue,
      targetTolerance,
      variables,
      constraints,
    ],
  );

  return (
    <div className="solver-dialog">
      <h2>{t("panels.solver.title")}</h2>

      <section>
        <label>
          {t("solver.dialog.solvingMethod")}
          <select
            value={method}
            onChange={(e) => setMethod(e.target.value as SolveMethod)}
          >
            <option value="simplex">{t("solver.dialog.method.simplex")}</option>
            <option value="grg">{t("solver.dialog.method.grg")}</option>
            <option value="evolutionary">{t("solver.dialog.method.evolutionary")}</option>
          </select>
        </label>
      </section>

      <section>
        <h3>{t("solver.dialog.objective.title")}</h3>
        <label>
          {t("solver.dialog.objective.setObjective")}
          <input
            value={objectiveRef}
            onChange={(e) => setObjectiveRef(e.target.value)}
          />
        </label>

        <label>
          {t("solver.dialog.objective.to")}
          <select
            value={objectiveKind}
            onChange={(e) => setObjectiveKind(e.target.value as ObjectiveKind)}
          >
            <option value="maximize">{t("solver.dialog.objective.max")}</option>
            <option value="minimize">{t("solver.dialog.objective.min")}</option>
            <option value="target">{t("solver.dialog.objective.valueOf")}</option>
          </select>
        </label>

        {canTarget && (
          <div style={{ display: "flex", gap: 12 }}>
            <label>
              {t("solver.dialog.objective.value")}
              <input
                type="number"
                value={targetValue}
                onChange={(e) => setTargetValue(Number(e.target.value))}
              />
            </label>
            <label>
              {t("solver.dialog.objective.tolerance")}
              <input
                type="number"
                value={targetTolerance}
                onChange={(e) => setTargetTolerance(Number(e.target.value))}
              />
            </label>
          </div>
        )}
      </section>

      <section>
        <h3>{t("solver.dialog.variables.title")}</h3>

        <button
          type="button"
          onClick={() => setVariables((v) => [...v, DEFAULT_VARIABLE])}
        >
          {t("solver.dialog.variables.addVariable")}
        </button>

        {variables.map((v, idx) => (
          <div key={idx} style={{ display: "grid", gap: 8, gridTemplateColumns: "2fr 1fr 1fr 1fr" }}>
            <input
              value={v.ref}
              onChange={(e) =>
                setVariables((vars) =>
                  vars.map((vv, i) => (i === idx ? { ...vv, ref: e.target.value } : vv)),
                )
              }
            />
            <select
              value={v.type}
              onChange={(e) =>
                setVariables((vars) =>
                  vars.map((vv, i) =>
                    i === idx ? { ...vv, type: e.target.value as SolverVariableSpec["type"] } : vv,
                  ),
                )
              }
            >
              <option value="continuous">{t("solver.dialog.variables.type.continuous")}</option>
              <option value="integer">{t("solver.dialog.variables.type.integer")}</option>
              <option value="binary">{t("solver.dialog.variables.type.binary")}</option>
            </select>
            <input
              type="number"
              placeholder={t("solver.dialog.variables.lower")}
              value={v.lower ?? ""}
              onChange={(e) =>
                setVariables((vars) =>
                  vars.map((vv, i) =>
                    i === idx
                      ? { ...vv, lower: e.target.value === "" ? undefined : Number(e.target.value) }
                      : vv,
                  ),
                )
              }
            />
            <input
              type="number"
              placeholder={t("solver.dialog.variables.upper")}
              value={v.upper ?? ""}
              onChange={(e) =>
                setVariables((vars) =>
                  vars.map((vv, i) =>
                    i === idx
                      ? { ...vv, upper: e.target.value === "" ? undefined : Number(e.target.value) }
                      : vv,
                  ),
                )
              }
            />
          </div>
        ))}
      </section>

      <section>
        <h3>{t("solver.dialog.constraints.title")}</h3>

        <button
          type="button"
          onClick={() => setConstraints((c) => [...c, DEFAULT_CONSTRAINT])}
        >
          {t("solver.dialog.constraints.addConstraint")}
        </button>

        {constraints.map((c, idx) => (
          <div key={idx} style={{ display: "grid", gap: 8, gridTemplateColumns: "2fr 1fr 1fr" }}>
            <input
              value={c.ref}
              onChange={(e) =>
                setConstraints((cons) =>
                  cons.map((cc, i) => (i === idx ? { ...cc, ref: e.target.value } : cc)),
                )
              }
            />
            <select
              value={c.relation}
              onChange={(e) =>
                setConstraints((cons) =>
                  cons.map((cc, i) =>
                    i === idx ? { ...cc, relation: e.target.value as Relation } : cc,
                  ),
                )
              }
            >
              <option value="<=">&le;</option>
              <option value=">=">&ge;</option>
              <option value="=">=</option>
            </select>
            <input
              type="number"
              value={c.rhs}
              onChange={(e) =>
                setConstraints((cons) =>
                  cons.map((cc, i) =>
                    i === idx ? { ...cc, rhs: Number(e.target.value) } : cc,
                  ),
                )
              }
            />
          </div>
        ))}
      </section>

      <footer style={{ display: "flex", gap: 12, justifyContent: "flex-end" }}>
        <button type="button" onClick={onCancel}>
          {t("solver.dialog.cancel")}
        </button>
        <button type="button" onClick={() => onRun(config)}>
          {t("solver.dialog.solve")}
        </button>
      </footer>
    </div>
  );
}
