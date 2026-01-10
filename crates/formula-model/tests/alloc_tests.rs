use std::alloc::{GlobalAlloc, Layout, System};
use std::sync::atomic::{AtomicUsize, Ordering};

use formula_model::{CellRef, CellValue, Worksheet};

/// A simple global allocator wrapper that counts allocated/deallocated bytes.
///
/// This is used to validate that sparse storage memory is proportional to the
/// number of stored (non-empty) cells, not the magnitude of their coordinates.
struct CountingAllocator;

static BYTES_ALLOCATED: AtomicUsize = AtomicUsize::new(0);
static BYTES_DEALLOCATED: AtomicUsize = AtomicUsize::new(0);

#[global_allocator]
static GLOBAL: CountingAllocator = CountingAllocator;

unsafe impl GlobalAlloc for CountingAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        BYTES_ALLOCATED.fetch_add(layout.size(), Ordering::SeqCst);
        System.alloc(layout)
    }

    unsafe fn dealloc(&self, ptr: *mut u8, layout: Layout) {
        BYTES_DEALLOCATED.fetch_add(layout.size(), Ordering::SeqCst);
        System.dealloc(ptr, layout)
    }

    unsafe fn realloc(&self, ptr: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        BYTES_DEALLOCATED.fetch_add(layout.size(), Ordering::SeqCst);
        BYTES_ALLOCATED.fetch_add(new_size, Ordering::SeqCst);
        System.realloc(ptr, layout, new_size)
    }

    unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
        BYTES_ALLOCATED.fetch_add(layout.size(), Ordering::SeqCst);
        System.alloc_zeroed(layout)
    }
}

fn live_bytes() -> usize {
    BYTES_ALLOCATED
        .load(Ordering::SeqCst)
        .saturating_sub(BYTES_DEALLOCATED.load(Ordering::SeqCst))
}

#[test]
fn sparse_storage_does_not_allocate_based_on_coordinate_magnitude() {
    let before = live_bytes();

    let mut sheet = Worksheet::new(1, "Sheet1");
    sheet.set_value(CellRef::new(0, 0), CellValue::Number(1.0));
    sheet.set_value(CellRef::new(1_000_000, 16_383), CellValue::Number(2.0));

    let after = live_bytes();
    let delta = after.saturating_sub(before);

    // This threshold is intentionally generous to avoid platform/allocator
    // variance, while still catching accidental dense pre-allocation.
    assert!(
        delta < 1_000_000,
        "expected sparse storage (<1MB for 2 cells), got {delta} bytes"
    );
}
