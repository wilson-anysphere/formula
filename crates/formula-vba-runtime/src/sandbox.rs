use std::time::Duration;

/// Sandboxed capabilities that VBA code can request.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Permission {
    FileSystemRead,
    FileSystemWrite,
    Network,
}

/// A host-provided permission checker (Task 32 integration point).
pub trait PermissionChecker: Send + Sync {
    fn has_permission(&self, permission: Permission) -> bool;
}

/// Default sandbox policy for running VBA.
#[derive(Debug, Clone)]
pub struct VbaSandboxPolicy {
    pub allow_filesystem_read: bool,
    pub allow_filesystem_write: bool,
    pub allow_network: bool,
    pub max_execution_time: Duration,
    /// Hard cap on the number of interpreter "steps" (roughly statement/expression
    /// evaluation units) to avoid pathological infinite loops even when the wall
    /// clock isn't advancing (e.g. busy-loop).
    pub max_steps: u64,
}

impl Default for VbaSandboxPolicy {
    fn default() -> Self {
        Self {
            allow_filesystem_read: false,
            allow_filesystem_write: false,
            allow_network: false,
            max_execution_time: Duration::from_millis(250),
            max_steps: 100_000,
        }
    }
}

impl VbaSandboxPolicy {
    pub fn can(&self, permission: Permission, checker: Option<&dyn PermissionChecker>) -> bool {
        let allowed_by_policy = match permission {
            Permission::FileSystemRead => self.allow_filesystem_read,
            Permission::FileSystemWrite => self.allow_filesystem_write,
            Permission::Network => self.allow_network,
        };

        if !allowed_by_policy {
            return false;
        }

        if let Some(checker) = checker {
            checker.has_permission(permission)
        } else {
            true
        }
    }
}
