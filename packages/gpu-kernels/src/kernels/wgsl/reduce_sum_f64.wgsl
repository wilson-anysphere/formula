enable f64;

alias Scalar = f64;

const WG_SIZE: u32 = 256u;

struct Params {
  length: u32,
  _pad0: u32,
  _pad1: u32,
  _pad2: u32,
}

@group(0) @binding(0) var<storage, read> input: array<Scalar>;
@group(0) @binding(1) var<storage, read_write> output: array<Scalar>;
@group(0) @binding(2) var<uniform> params: Params;

var<workgroup> shared: array<Scalar, WG_SIZE>;

@compute @workgroup_size(WG_SIZE)
fn main(
  @builtin(local_invocation_id) lid: vec3<u32>,
  @builtin(workgroup_id) wid: vec3<u32>,
  @builtin(num_workgroups) nwg: vec3<u32>
) {
  let local = lid.x;
  let wg_index = wid.x + wid.y * nwg.x;
  let base = wg_index * WG_SIZE * 2u;

  var sum: Scalar = 0.0;
  let idx1 = base + local;
  let idx2 = base + local + WG_SIZE;
  if (idx1 < params.length) {
    sum = sum + input[idx1];
  }
  if (idx2 < params.length) {
    sum = sum + input[idx2];
  }

  shared[local] = sum;
  workgroupBarrier();

  var stride = WG_SIZE / 2u;
  while (stride > 0u) {
    if (local < stride) {
      shared[local] = shared[local] + shared[local + stride];
    }
    workgroupBarrier();
    stride = stride / 2u;
  }

  if (local == 0u) {
    output[wg_index] = shared[0];
  }
}
