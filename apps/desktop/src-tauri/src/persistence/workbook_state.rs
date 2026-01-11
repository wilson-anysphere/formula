use formula_storage::{AutoSaveManager, MemoryManager, MemoryManagerConfig, Storage};
use std::collections::HashMap;
use std::path::PathBuf;
use std::sync::Arc;
use uuid::Uuid;

#[derive(Clone, Debug)]
pub enum WorkbookPersistenceLocation {
    InMemory,
    OnDisk(PathBuf),
}

pub struct PersistentWorkbookState {
    pub location: WorkbookPersistenceLocation,
    pub storage: Storage,
    pub memory: MemoryManager,
    pub autosave: Option<Arc<AutoSaveManager>>,
    pub workbook_id: Uuid,
    pub sheet_map: HashMap<String, Uuid>,
}

impl PersistentWorkbookState {
    pub fn sheet_uuid(&self, sheet_id: &str) -> Option<Uuid> {
        self.sheet_map.get(sheet_id).copied()
    }

    pub async fn flush_autosave(&self) -> formula_storage::storage::Result<()> {
        if let Some(manager) = self.autosave.as_ref() {
            manager.flush().await
        } else {
            Ok(())
        }
    }

    pub fn autosave_save_count(&self) -> usize {
        self.autosave.as_ref().map(|a| a.save_count()).unwrap_or(0)
    }
}

pub fn open_storage(location: &WorkbookPersistenceLocation) -> formula_storage::storage::Result<Storage> {
    match location {
        WorkbookPersistenceLocation::InMemory => Storage::open_in_memory(),
        WorkbookPersistenceLocation::OnDisk(path) => Storage::open_path(path),
    }
}

pub fn open_memory_manager(storage: Storage) -> MemoryManager {
    MemoryManager::new(storage, MemoryManagerConfig::default())
}

