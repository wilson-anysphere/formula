alias Scalar = f32;

const WG_SIZE: u32 = 256u;

struct Params {
  n: u32,
  j: u32,
  k: u32,
  _pad0: u32,
}

@group(0) @binding(0) var<storage, read_write> data: array<Scalar>;
@group(0) @binding(1) var<uniform> params: Params;

@compute @workgroup_size(WG_SIZE)
fn main(@builtin(global_invocation_id) gid: vec3<u32>) {
  let i = gid.x;
  if (i >= params.n) {
    return;
  }

  let ixj = i ^ params.j;
  if (ixj <= i || ixj >= params.n) {
    return;
  }

  let ascending = (i & params.k) == 0u;
  let a = data[i];
  let b = data[ixj];

  if (ascending) {
    if (a > b) {
      data[i] = b;
      data[ixj] = a;
    }
  } else {
    if (a < b) {
      data[i] = b;
      data[ixj] = a;
    }
  }
}

