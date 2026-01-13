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
      <h2 className="solver-dialog__title">{t("panels.solver.title")}</h2>

      <section className="solver-dialog__section">
        <label className="solver-dialog__field">
          <span className="solver-dialog__label">{t("solver.dialog.solvingMethod")}</span>
          <select
            className="solver-dialog__select"
            value={method}
            onChange={(e) => setMethod(e.target.value as SolveMethod)}
          >
            <option value="simplex">{t("solver.dialog.method.simplex")}</option>
            <option value="grg">{t("solver.dialog.method.grg")}</option>
            <option value="evolutionary">{t("solver.dialog.method.evolutionary")}</option>
          </select>
        </label>
      </section>

      <section className="solver-dialog__section">
        <h3 className="solver-dialog__section-title">{t("solver.dialog.objective.title")}</h3>
        <label className="solver-dialog__field">
          <span className="solver-dialog__label">{t("solver.dialog.objective.setObjective")}</span>
          <input
            className="solver-dialog__input solver-dialog__input--mono"
            value={objectiveRef}
            onChange={(e) => setObjectiveRef(e.target.value)}
          />
        </label>

        <label className="solver-dialog__field">
          <span className="solver-dialog__label">{t("solver.dialog.objective.to")}</span>
          <select
            className="solver-dialog__select"
            value={objectiveKind}
            onChange={(e) => setObjectiveKind(e.target.value as ObjectiveKind)}
          >
            <option value="maximize">{t("solver.dialog.objective.max")}</option>
            <option value="minimize">{t("solver.dialog.objective.min")}</option>
            <option value="target">{t("solver.dialog.objective.valueOf")}</option>
          </select>
        </label>

        {canTarget && (
          <div className="solver-dialog__row">
            <label className="solver-dialog__field">
              <span className="solver-dialog__label">{t("solver.dialog.objective.value")}</span>
              <input
                className="solver-dialog__input"
                type="number"
                value={targetValue}
                onChange={(e) => setTargetValue(Number(e.target.value))}
              />
            </label>
            <label className="solver-dialog__field">
              <span className="solver-dialog__label">{t("solver.dialog.objective.tolerance")}</span>
              <input
                className="solver-dialog__input"
                type="number"
                value={targetTolerance}
                onChange={(e) => setTargetTolerance(Number(e.target.value))}
              />
            </label>
          </div>
        )}
      </section>

      <section className="solver-dialog__section">
        <h3 className="solver-dialog__section-title">{t("solver.dialog.variables.title")}</h3>

        <button
          type="button"
          className="solver__button"
          onClick={() => setVariables((v) => [...v, DEFAULT_VARIABLE])}
        >
          {t("solver.dialog.variables.addVariable")}
        </button>

        {variables.map((v, idx) => (
          <div key={idx} className="solver-dialog__variable-row">
            <input
              className="solver-dialog__input solver-dialog__input--mono"
              value={v.ref}
              onChange={(e) =>
                setVariables((vars) =>
                  vars.map((vv, i) => (i === idx ? { ...vv, ref: e.target.value } : vv)),
                )
              }
            />
            <select
              className="solver-dialog__select"
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
              className="solver-dialog__input"
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
              className="solver-dialog__input"
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

      <section className="solver-dialog__section">
        <h3 className="solver-dialog__section-title">{t("solver.dialog.constraints.title")}</h3>

        <button
          type="button"
          className="solver__button"
          onClick={() => setConstraints((c) => [...c, DEFAULT_CONSTRAINT])}
        >
          {t("solver.dialog.constraints.addConstraint")}
        </button>

        {constraints.map((c, idx) => (
          <div key={idx} className="solver-dialog__constraint-row">
            <input
              className="solver-dialog__input solver-dialog__input--mono"
              value={c.ref}
              onChange={(e) =>
                setConstraints((cons) =>
                  cons.map((cc, i) => (i === idx ? { ...cc, ref: e.target.value } : cc)),
                )
              }
            />
            <select
              className="solver-dialog__select"
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
              className="solver-dialog__input"
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

      <footer className="solver-dialog__footer">
        <button type="button" className="solver__button" onClick={onCancel}>
          {t("solver.dialog.cancel")}
        </button>
        <button type="button" className="solver__button solver__button--primary" onClick={() => onRun(config)}>
          {t("solver.dialog.solve")}
        </button>
      </footer>
    </div>
  );
}
