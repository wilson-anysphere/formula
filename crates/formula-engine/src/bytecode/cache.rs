use super::ast::Expr;
use super::compiler::{normalized_key, Compiler};
use super::Program;
#[cfg(all(feature = "parallel", not(target_arch = "wasm32")))]
use dashmap::DashMap;
#[cfg(not(all(feature = "parallel", not(target_arch = "wasm32"))))]
use std::collections::HashMap;
use std::sync::Arc;
#[cfg(not(all(feature = "parallel", not(target_arch = "wasm32"))))]
use std::sync::Mutex;

/// Cache compiled bytecode programs keyed by normalized formula.
///
/// The key is derived from the normalized AST where relative references are expressed
/// as offsets (so the same pattern filled across a range hits the same cache entry).
#[derive(Default)]
pub struct BytecodeCache {
    #[cfg(all(feature = "parallel", not(target_arch = "wasm32")))]
    programs: DashMap<Arc<str>, Arc<Program>>,
    #[cfg(not(all(feature = "parallel", not(target_arch = "wasm32"))))]
    programs: Mutex<HashMap<Arc<str>, Arc<Program>>>,
}

impl BytecodeCache {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn program_count(&self) -> usize {
        #[cfg(all(feature = "parallel", not(target_arch = "wasm32")))]
        {
            self.programs.len()
        }
        #[cfg(not(all(feature = "parallel", not(target_arch = "wasm32"))))]
        {
            self.programs
                .lock()
                .unwrap_or_else(|poison| poison.into_inner())
                .len()
        }
    }

    pub fn get_or_compile(&self, expr: &Expr) -> Arc<Program> {
        let key = normalized_key(expr);
        let key_for_compile = key.clone();

        #[cfg(all(feature = "parallel", not(target_arch = "wasm32")))]
        {
            self.programs
                .entry(key)
                .or_insert_with(|| Arc::new(Compiler::compile(key_for_compile, expr)))
                .clone()
        }

        #[cfg(not(all(feature = "parallel", not(target_arch = "wasm32"))))]
        {
            let mut programs = self
                .programs
                .lock()
                .unwrap_or_else(|poison| poison.into_inner());
            if let Some(program) = programs.get(&key) {
                return program.clone();
            }
            let program = Arc::new(Compiler::compile(key_for_compile, expr));
            programs.insert(key, program.clone());
            program
        }
    }
}
