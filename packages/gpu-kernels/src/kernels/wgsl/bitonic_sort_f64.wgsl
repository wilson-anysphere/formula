enable f64;

alias Scalar = f64;

const WG_SIZE: u32 = 256u;

struct Params {
  n: u32,
  j: u32,
  k: u32,
  _pad0: u32,
}

@group(0) @binding(0) var<storage, read_write> data: array<Scalar>;
@group(0) @binding(1) var<uniform> params: Params;

fn cmp(a: Scalar, b: Scalar) -> i32 {
  let a_nan = isNan(a);
  let b_nan = isNan(b);
  if (a_nan && b_nan) {
    return 0;
  }
  if (a_nan) {
    // Match TypedArray sorting: NaN compares greater than any number.
    return 1;
  }
  if (b_nan) {
    return -1;
  }
  // Match TypedArray sorting for signed zero: -0 compares less than +0.
  if (a == 0.0 && b == 0.0) {
    let a_neg_zero = (1.0 / a) < 0.0;
    let b_neg_zero = (1.0 / b) < 0.0;
    if (a_neg_zero && !b_neg_zero) {
      return -1;
    }
    if (!a_neg_zero && b_neg_zero) {
      return 1;
    }
    return 0;
  }
  if (a < b) {
    return -1;
  }
  if (a > b) {
    return 1;
  }
  return 0;
}

@compute @workgroup_size(WG_SIZE)
fn main(@builtin(global_invocation_id) gid: vec3<u32>, @builtin(num_workgroups) nwg: vec3<u32>) {
  let stride = nwg.x * WG_SIZE;
  let i = gid.x + gid.y * stride;
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
    if (cmp(a, b) > 0) {
      data[i] = b;
      data[ixj] = a;
    }
  } else {
    if (cmp(a, b) < 0) {
      data[i] = b;
      data[ixj] = a;
    }
  }
}
