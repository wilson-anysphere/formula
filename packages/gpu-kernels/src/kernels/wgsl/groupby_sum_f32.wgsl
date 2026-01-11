alias Scalar = f32;

const WG_SIZE: u32 = 256u;
const EMPTY_KEY: u32 = 0xffff_ffffu;

struct Params {
  n: u32,
  table_size: u32,
  _pad0: u32,
  _pad1: u32,
}

@group(0) @binding(0) var<storage, read> in_keys: array<u32>;
@group(0) @binding(1) var<storage, read> in_vals: array<Scalar>;
@group(0) @binding(2) var<storage, read_write> table_keys: array<atomic<u32>>;
@group(0) @binding(3) var<storage, read_write> table_counts: array<atomic<u32>>;
// f32 stored as u32 bits (with CAS-based atomic add).
@group(0) @binding(4) var<storage, read_write> table_sums: array<atomic<u32>>;
@group(0) @binding(5) var<uniform> params: Params;

fn hash(key: u32) -> u32 {
  return (key * 2654435761u) & (params.table_size - 1u);
}

fn find_or_insert_slot(key: u32) -> u32 {
  if (key == EMPTY_KEY) {
    return params.table_size;
  }

  var slot = hash(key);
  loop {
    let existing = atomicLoad(&table_keys[slot]);
    if (existing == key) {
      return slot;
    }
    if (existing == EMPTY_KEY) {
      let res = atomicCompareExchangeWeak(&table_keys[slot], EMPTY_KEY, key);
      if (res.exchanged || res.old_value == key) {
        return slot;
      }
    }
    slot = (slot + 1u) & (params.table_size - 1u);
  }
  return slot;
}

fn atomic_add_f32(addr: ptr<storage, atomic<u32>, read_write>, value: Scalar) {
  loop {
    let old_bits = atomicLoad(addr);
    let old = bitcast<Scalar>(old_bits);
    let new = old + value;
    let new_bits = bitcast<u32>(new);
    let res = atomicCompareExchangeWeak(addr, old_bits, new_bits);
    if (res.exchanged) {
      break;
    }
  }
}

@compute @workgroup_size(WG_SIZE)
fn main(@builtin(global_invocation_id) gid: vec3<u32>, @builtin(num_workgroups) nwg: vec3<u32>) {
  let stride = nwg.x * WG_SIZE;
  let i = gid.x + gid.y * stride;
  if (i >= params.n) {
    return;
  }

  let key = in_keys[i];
  let v = in_vals[i];
  let slot = find_or_insert_slot(key);
  atomicAdd(&table_counts[slot], 1u);
  atomic_add_f32(&table_sums[slot], v);
}

