const WG_SIZE: u32 = 256u;
const EMPTY_U32: u32 = 0xffff_ffffu;

struct Params {
  n: u32,
  _pad0: u32,
  _pad1: u32,
  _pad2: u32,
}

@group(0) @binding(0) var<storage, read_write> table_keys: array<atomic<u32>>;
@group(0) @binding(1) var<storage, read_write> table_heads: array<atomic<u32>>;
@group(0) @binding(2) var<uniform> params: Params;

@compute @workgroup_size(WG_SIZE)
fn main(@builtin(global_invocation_id) gid: vec3<u32>, @builtin(num_workgroups) nwg: vec3<u32>) {
  let stride = nwg.x * WG_SIZE;
  let i = gid.x + gid.y * stride;
  if (i >= params.n) {
    return;
  }

  atomicStore(&table_keys[i], EMPTY_U32);
  atomicStore(&table_heads[i], EMPTY_U32);
}

