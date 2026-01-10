use super::ast::Expr;
use super::compiler::{normalized_key, Compiler};
use super::Program;
use dashmap::DashMap;
use std::sync::Arc;

/// Cache compiled bytecode programs keyed by normalized formula.
///
/// The key is derived from the normalized AST where relative references are expressed
/// as offsets (so the same pattern filled across a range hits the same cache entry).
#[derive(Default)]
pub struct BytecodeCache {
    programs: DashMap<Arc<str>, Arc<Program>>,
}

impl BytecodeCache {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn program_count(&self) -> usize {
        self.programs.len()
    }

    pub fn get_or_compile(&self, expr: &Expr) -> Arc<Program> {
        let key = normalized_key(expr);
        let key_for_compile = key.clone();
        self.programs
            .entry(key)
            .or_insert_with(|| Arc::new(Compiler::compile(key_for_compile, expr)))
            .clone()
    }
}
