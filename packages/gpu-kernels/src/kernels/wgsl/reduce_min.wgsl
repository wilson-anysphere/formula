alias Scalar = f32;

const WG_SIZE: u32 = 256u;

fn nan_propagating_min(a: Scalar, b: Scalar) -> Scalar {
  if (isNan(a) || isNan(b)) {
    // Propagate NaN deterministically.
    return a + b;
  }
  if (a < b) {
    return a;
  }
  return b;
}

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
  @builtin(workgroup_id) wid: vec3<u32>
) {
  let local = lid.x;
  let base = wid.x * WG_SIZE * 2u;

  var acc: Scalar = Scalar(0.0);
  var has: bool = false;

  let idx1 = base + local;
  let idx2 = base + local + WG_SIZE;
  if (idx1 < params.length) {
    acc = input[idx1];
    has = true;
  }
  if (idx2 < params.length) {
    let v = input[idx2];
    acc = select(v, nan_propagating_min(acc, v), has);
    has = true;
  }

  shared[local] = select(Scalar(1.0 / 0.0), acc, has);
  workgroupBarrier();

  var stride = WG_SIZE / 2u;
  while (stride > 0u) {
    if (local < stride) {
      shared[local] = nan_propagating_min(shared[local], shared[local + stride]);
    }
    workgroupBarrier();
    stride = stride / 2u;
  }

  if (local == 0u) {
    output[wid.x] = shared[0];
  }
}
