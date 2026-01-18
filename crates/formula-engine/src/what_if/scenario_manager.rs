use std::collections::HashMap;
use std::time::{SystemTime, UNIX_EPOCH};

use crate::what_if::{CellRef, CellValue, WhatIfError, WhatIfModel};
use serde::{Deserialize, Serialize};

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash, Serialize, Deserialize)]
#[serde(transparent)]
pub struct ScenarioId(u64);

impl ScenarioId {
    pub fn new(id: u64) -> Self {
        Self(id)
    }

    pub fn as_u64(&self) -> u64 {
        self.0
    }
}

impl From<u64> for ScenarioId {
    fn from(value: u64) -> Self {
        ScenarioId::new(value)
    }
}

#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Scenario {
    pub id: ScenarioId,
    pub name: String,
    pub changing_cells: Vec<CellRef>,
    pub values: HashMap<CellRef, CellValue>,
    pub created_ms: u64,
    pub created_by: String,
    pub comment: Option<String>,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SummaryReport {
    pub changing_cells: Vec<CellRef>,
    pub result_cells: Vec<CellRef>,
    /// scenario_name -> (cell -> value)
    pub results: HashMap<String, HashMap<CellRef, CellValue>>,
}

#[derive(Debug, Default, Clone)]
pub struct ScenarioManager {
    scenarios: HashMap<ScenarioId, Scenario>,
    current_scenario: Option<ScenarioId>,
    base_values: HashMap<CellRef, CellValue>,
    next_id: u64,
}

impl ScenarioManager {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn scenarios(&self) -> impl Iterator<Item = &Scenario> {
        self.scenarios.values()
    }

    pub fn get(&self, id: ScenarioId) -> Option<&Scenario> {
        self.scenarios.get(&id)
    }

    pub fn create_scenario(
        &mut self,
        name: impl Into<String>,
        changing_cells: Vec<CellRef>,
        values: Vec<CellValue>,
        created_by: impl Into<String>,
        comment: Option<String>,
    ) -> Result<ScenarioId, WhatIfError<&'static str>> {
        if changing_cells.len() != values.len() {
            return Err(WhatIfError::InvalidParams(
                "changing_cells and values must have equal length",
            ));
        }

        let mut value_map: HashMap<CellRef, CellValue> = HashMap::new();
        if value_map.try_reserve(changing_cells.len()).is_err() {
            debug_assert!(false, "allocation failed (scenario values)");
            return Err(WhatIfError::NumericalFailure("allocation failed"));
        }
        for (cell, value) in changing_cells.iter().cloned().zip(values.into_iter()) {
            value_map.insert(cell, value);
        }

        let id = ScenarioId(self.next_id);
        self.next_id = self.next_id.wrapping_add(1);

        if self.scenarios.try_reserve(1).is_err() {
            debug_assert!(false, "allocation failed (scenarios)");
            return Err(WhatIfError::NumericalFailure("allocation failed"));
        }
        self.scenarios.insert(
            id,
            Scenario {
                id,
                name: name.into(),
                changing_cells,
                values: value_map,
                created_ms: now_ms(),
                created_by: created_by.into(),
                comment,
            },
        );

        Ok(id)
    }

    pub fn delete_scenario(&mut self, id: ScenarioId) -> bool {
        if self.current_scenario == Some(id) {
            self.current_scenario = None;
        }
        self.scenarios.remove(&id).is_some()
    }

    pub fn current_scenario(&self) -> Option<ScenarioId> {
        self.current_scenario
    }

    pub fn clear_base_values(&mut self) {
        self.base_values.clear();
    }

    pub fn base_values(&self) -> &HashMap<CellRef, CellValue> {
        &self.base_values
    }

    pub fn restore_base<M: WhatIfModel>(
        &mut self,
        model: &mut M,
    ) -> Result<(), WhatIfError<M::Error>> {
        if self.base_values.is_empty() {
            return Ok(());
        }

        for (cell, value) in &self.base_values {
            model.set_cell_value(cell, value.clone())?;
        }
        model.recalculate()?;

        self.current_scenario = None;
        Ok(())
    }

    pub fn apply_scenario<M: WhatIfModel>(
        &mut self,
        model: &mut M,
        id: ScenarioId,
    ) -> Result<(), WhatIfError<M::Error>> {
        let scenario = self
            .scenarios
            .get(&id)
            .ok_or_else(|| WhatIfError::InvalidParams("scenario not found"))?;

        // Capture base values for any changing cells we haven't seen yet.
        //
        // Scenarios may have different changing cell sets, so the base snapshot
        // needs to be the union of inputs across all applied scenarios.
        if self
            .base_values
            .try_reserve(scenario.changing_cells.len())
            .is_err()
        {
            debug_assert!(false, "allocation failed (base scenario values)");
            return Err(WhatIfError::NumericalFailure("allocation failed"));
        }
        for cell in &scenario.changing_cells {
            if !self.base_values.contains_key(cell) {
                let base = model.get_cell_value(cell)?;
                self.base_values.insert(cell.clone(), base);
            }
        }

        for (cell, value) in &scenario.values {
            model.set_cell_value(cell, value.clone())?;
        }
        model.recalculate()?;

        self.current_scenario = Some(id);
        Ok(())
    }

    pub fn generate_summary_report<M: WhatIfModel>(
        &mut self,
        model: &mut M,
        result_cells: Vec<CellRef>,
        scenario_ids: Vec<ScenarioId>,
    ) -> Result<SummaryReport, WhatIfError<M::Error>> {
        self.restore_base(model)?;

        let mut results: HashMap<String, HashMap<CellRef, CellValue>> = HashMap::new();
        if results
            .try_reserve(scenario_ids.len().saturating_add(1))
            .is_err()
        {
            debug_assert!(false, "allocation failed (scenario summary results)");
            return Err(WhatIfError::NumericalFailure("allocation failed"));
        }

        // Base case.
        let mut base_row: HashMap<CellRef, CellValue> = HashMap::new();
        if base_row.try_reserve(result_cells.len()).is_err() {
            debug_assert!(false, "allocation failed (scenario summary base row)");
            return Err(WhatIfError::NumericalFailure("allocation failed"));
        }
        for cell in &result_cells {
            base_row.insert(cell.clone(), model.get_cell_value(cell)?);
        }
        results.insert("Base".to_string(), base_row);

        // Each scenario.
        for id in &scenario_ids {
            self.apply_scenario(model, *id)?;
            let scenario = self
                .scenarios
                .get(id)
                .ok_or_else(|| WhatIfError::InvalidParams("scenario not found"))?;

            let mut row: HashMap<CellRef, CellValue> = HashMap::new();
            if row.try_reserve(result_cells.len()).is_err() {
                debug_assert!(false, "allocation failed (scenario summary row)");
                return Err(WhatIfError::NumericalFailure("allocation failed"));
            }
            for cell in &result_cells {
                row.insert(cell.clone(), model.get_cell_value(cell)?);
            }
            results.insert(scenario.name.clone(), row);
        }

        self.restore_base(model)?;

        let changing_cells = scenario_ids
            .first()
            .and_then(|id| self.scenarios.get(id))
            .map(|s| s.changing_cells.clone())
            .unwrap_or_default();

        Ok(SummaryReport {
            changing_cells,
            result_cells,
            results,
        })
    }
}

fn now_ms() -> u64 {
    SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|d| d.as_millis() as u64)
        .unwrap_or(0)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::what_if::{CellValue, InMemoryModel, WhatIfModel};

    #[test]
    fn restore_base_covers_union_of_changing_cells_across_scenarios() {
        let mut model = InMemoryModel::new()
            .with_value("A1", CellValue::Number(10.0))
            .with_value("B1", CellValue::Number(20.0))
            .with_value("C1", CellValue::Number(30.0));

        let mut manager = ScenarioManager::new();

        let s1 = manager
            .create_scenario(
                "S1",
                vec![CellRef::from("A1"), CellRef::from("B1")],
                vec![CellValue::Number(1.0), CellValue::Number(2.0)],
                "tester",
                None,
            )
            .unwrap();

        let s2 = manager
            .create_scenario(
                "S2",
                vec![CellRef::from("B1"), CellRef::from("C1")],
                vec![CellValue::Number(200.0), CellValue::Number(300.0)],
                "tester",
                None,
            )
            .unwrap();

        manager.apply_scenario(&mut model, s1).unwrap();
        manager.apply_scenario(&mut model, s2).unwrap();

        // Scenario 2 doesn't touch A1, so it retains scenario 1's value.
        assert_eq!(
            model.get_cell_value(&CellRef::from("A1")).unwrap(),
            CellValue::Number(1.0)
        );
        assert_eq!(
            model.get_cell_value(&CellRef::from("B1")).unwrap(),
            CellValue::Number(200.0)
        );
        assert_eq!(
            model.get_cell_value(&CellRef::from("C1")).unwrap(),
            CellValue::Number(300.0)
        );

        manager.restore_base(&mut model).unwrap();

        assert_eq!(
            model.get_cell_value(&CellRef::from("A1")).unwrap(),
            CellValue::Number(10.0)
        );
        assert_eq!(
            model.get_cell_value(&CellRef::from("B1")).unwrap(),
            CellValue::Number(20.0)
        );
        assert_eq!(
            model.get_cell_value(&CellRef::from("C1")).unwrap(),
            CellValue::Number(30.0)
        );
    }
}
