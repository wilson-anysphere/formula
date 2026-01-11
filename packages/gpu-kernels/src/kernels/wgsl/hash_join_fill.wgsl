const WG_SIZE: u32 = 256u;
const EMPTY_U32: u32 = 0xffff_ffffu;

struct Params {
  n_left: u32,
  table_size: u32,
  _pad0: u32,
  _pad1: u32,
}

@group(0) @binding(0) var<storage, read> left_keys: array<u32>;
@group(0) @binding(1) var<storage, read> table_keys: array<atomic<u32>>;
@group(0) @binding(2) var<storage, read> table_heads: array<atomic<u32>>;
@group(0) @binding(3) var<storage, read> next_ptr: array<u32>;
@group(0) @binding(4) var<storage, read> counts: array<u32>;
@group(0) @binding(5) var<storage, read> offsets: array<u32>;
@group(0) @binding(6) var<storage, read_write> out_left_index: array<u32>;
@group(0) @binding(7) var<storage, read_write> out_right_index: array<u32>;
@group(0) @binding(8) var<uniform> params: Params;

fn hash(key: u32) -> u32 {
  return (key * 2654435761u) & (params.table_size - 1u);
}

fn find_slot(key: u32) -> u32 {
  if (key == EMPTY_U32) {
    return params.table_size;
  }

  var slot = hash(key);
  loop {
    let existing = atomicLoad(&table_keys[slot]);
    if (existing == key) {
      return slot;
    }
    if (existing == EMPTY_U32) {
      return EMPTY_U32;
    }
    slot = (slot + 1u) & (params.table_size - 1u);
  }
  return EMPTY_U32;
}

@compute @workgroup_size(WG_SIZE)
fn main(@builtin(global_invocation_id) gid: vec3<u32>, @builtin(num_workgroups) nwg: vec3<u32>) {
  let stride = nwg.x * WG_SIZE;
  let i = gid.x + gid.y * stride;
  if (i >= params.n_left) {
    return;
  }

  let count = counts[i];
  if (count == 0u) {
    return;
  }

  let key = left_keys[i];
  let slot = find_slot(key);
  if (slot == EMPTY_U32) {
    return;
  }

  let base = offsets[i];
  var ptr = atomicLoad(&table_heads[slot]);
  var t: u32 = 0u;
  loop {
    if (ptr == EMPTY_U32 || t >= count) {
      break;
    }
    out_left_index[base + t] = i;
    out_right_index[base + t] = ptr;
    ptr = next_ptr[ptr];
    t = t + 1u;
  }
}

