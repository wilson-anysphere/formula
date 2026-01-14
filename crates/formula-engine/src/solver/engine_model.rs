use super::{SolverError, SolverModel};
use crate::coercion::number::parse_number_coercion;
use crate::coercion::ValueLocaleConfig;
use crate::{Engine, Value};

#[derive(Clone, Debug)]
struct CellAddress {
    sheet: String,
    addr: String,
}

impl CellAddress {
    fn parse(input: &str, default_sheet: &str) -> Result<Self, SolverError> {
        let trimmed = input.trim();
        if trimmed.is_empty() {
            return Err(SolverError::new("cell reference cannot be empty"));
        }

        let (sheet_part, addr_part) = match trimmed.rsplit_once('!') {
            Some((sheet, addr)) => (sheet.trim(), addr.trim()),
            None => (default_sheet, trimmed),
        };

        if addr_part.is_empty() {
            return Err(SolverError::new(format!(
                "invalid cell reference '{input}': missing address"
            )));
        }

        let sheet = unquote_sheet_name(sheet_part);
        if sheet.is_empty() {
            return Err(SolverError::new(format!(
                "invalid cell reference '{input}': missing sheet name"
            )));
        }

        Ok(Self {
            sheet,
            addr: addr_part.to_string(),
        })
    }
}

fn unquote_sheet_name(name: &str) -> String {
    let name = name.trim();
    if name.starts_with('\'') && name.ends_with('\'') && name.len() >= 2 {
        return name[1..name.len() - 1].replace("''", "'");
    }
    name.to_string()
}

/// Adapter that exposes a [`crate::Engine`] workbook as a [`SolverModel`].
///
/// The solver iteratively overwrites the provided decision-variable cells,
/// triggers engine recalculation, then reads objective and constraint cell values.
///
/// Cell references may be provided as `Sheet1!A1`, `A1` (default sheet), or
/// `\'My Sheet\'!A1`.
pub struct EngineSolverModel<'a> {
    engine: &'a mut Engine,
    vars: Vec<CellAddress>,
    objective_cell: CellAddress,
    constraint_cells: Vec<CellAddress>,
    cached_vars: Vec<f64>,
    cached_objective: f64,
    cached_constraints: Vec<f64>,
}

impl<'a> EngineSolverModel<'a> {
    pub fn new(
        engine: &'a mut Engine,
        default_sheet: impl Into<String>,
        objective_cell: &str,
        vars: Vec<&str>,
        constraint_cells: Vec<&str>,
    ) -> Result<Self, SolverError> {
        let default_sheet = default_sheet.into();
        let objective_cell = CellAddress::parse(objective_cell, &default_sheet)?;
        let vars = vars
            .into_iter()
            .map(|s| CellAddress::parse(s, &default_sheet))
            .collect::<Result<Vec<_>, _>>()?;
        let constraint_cells = constraint_cells
            .into_iter()
            .map(|s| CellAddress::parse(s, &default_sheet))
            .collect::<Result<Vec<_>, _>>()?;

        let cached_vars = vec![0.0; vars.len()];
        let cached_constraints = vec![0.0; constraint_cells.len()];

        let mut model = Self {
            engine,
            vars,
            objective_cell,
            constraint_cells,
            cached_vars,
            cached_objective: f64::NAN,
            cached_constraints,
        };

        model.refresh_cache()?;
        Ok(model)
    }

    fn refresh_cache(&mut self) -> Result<(), SolverError> {
        self.cached_objective = self.read_cell_number_or_nan(&self.objective_cell);
        for (idx, cell) in self.constraint_cells.iter().enumerate() {
            self.cached_constraints[idx] = self.read_cell_number_or_nan(cell);
        }
        for (idx, cell) in self.vars.iter().enumerate() {
            self.cached_vars[idx] = self
                .read_cell_number_strict(cell)
                .map_err(|msg| SolverError::new(msg))?;
        }
        Ok(())
    }

    fn read_cell_number_or_nan(&self, cell: &CellAddress) -> f64 {
        let locale = self.engine.value_locale();
        match coerce_value_to_number(&self.engine.get_cell_value(&cell.sheet, &cell.addr), locale) {
            Some(n) => n,
            None => f64::NAN,
        }
    }

    fn read_cell_number_strict(&self, cell: &CellAddress) -> Result<f64, String> {
        let value = self.engine.get_cell_value(&cell.sheet, &cell.addr);
        let locale = self.engine.value_locale();
        coerce_value_to_number(&value, locale).ok_or_else(|| {
            format!(
                "cell {}!{} is not numeric (value: {value})",
                cell.sheet, cell.addr
            )
        })
    }
}

impl SolverModel for EngineSolverModel<'_> {
    fn num_vars(&self) -> usize {
        self.vars.len()
    }

    fn num_constraints(&self) -> usize {
        self.constraint_cells.len()
    }

    fn get_vars(&self, out: &mut [f64]) {
        out.copy_from_slice(&self.cached_vars);
    }

    fn set_vars(&mut self, vars: &[f64]) -> Result<(), SolverError> {
        if vars.len() != self.vars.len() {
            return Err(SolverError::new("wrong variable vector length"));
        }
        for (val, cell) in vars.iter().zip(self.vars.iter()) {
            self.engine
                .set_cell_value(&cell.sheet, &cell.addr, *val)
                .map_err(|e| SolverError::new(e.to_string()))?;
        }
        self.cached_vars.copy_from_slice(vars);
        Ok(())
    }

    fn recalc(&mut self) -> Result<(), SolverError> {
        self.engine.recalculate();
        self.refresh_cache()
    }

    fn objective(&self) -> f64 {
        self.cached_objective
    }

    fn constraints(&self, out: &mut [f64]) {
        out.copy_from_slice(&self.cached_constraints);
    }
}

fn coerce_text_to_number(text: &str, value_locale: ValueLocaleConfig) -> Option<f64> {
    parse_number_coercion(
        text,
        value_locale.separators.decimal_sep,
        Some(value_locale.separators.thousands_sep),
    )
    .ok()
}

fn coerce_value_to_number(value: &Value, value_locale: ValueLocaleConfig) -> Option<f64> {
    match value {
        Value::Number(n) => n.is_finite().then_some(*n),
        Value::Bool(b) => Some(if *b { 1.0 } else { 0.0 }),
        Value::Blank => Some(0.0),
        Value::Text(s) => coerce_text_to_number(s, value_locale),
        // Solver only supports numeric values (variables/constraints/objectives are `f64`).
        //
        // For rich scalar values (e.g. Entity/Record) we follow the same behavior as text: attempt
        // to parse the *display string* as a number using the engine's value-locale rules.
        //
        // This is intentionally more permissive than `str::parse::<f64>()`:
        // - accepts thousands separators (locale-aware),
        // - accepts common currency symbols, accounting parentheses, and percent signs,
        // - rejects `NaN`/`Inf` textual inputs (Excel-compatible).
        //
        // If the display string isn't a valid number, treat the value as non-numeric.
        //
        // NOTE: Arrays are explicitly treated as non-numeric even though their `Display`
        // implementation shows the top-left element. Solver decision variables must be
        // scalar cells.
        Value::Array(_) => None,
        other => coerce_text_to_number(&other.to_string(), value_locale),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn coerce_value_to_number_returns_none_for_non_numeric_values() {
        let value = Value::Array(crate::value::Array::new(1, 1, vec![Value::Number(1.0)]));
        assert_eq!(
            coerce_value_to_number(&value, ValueLocaleConfig::en_us()),
            None
        );
    }
}
