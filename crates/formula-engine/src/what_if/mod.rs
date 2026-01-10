//! What‑If analysis tools (Goal Seek, Scenario Manager, Monte Carlo simulation).

mod types;

pub mod goal_seek;
pub mod monte_carlo;
pub mod scenario_manager;

pub use types::{CellRef, CellValue, InMemoryModel, WhatIfError, WhatIfModel};
