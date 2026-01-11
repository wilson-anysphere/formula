enable f64;

alias Scalar = f64;

const TILE: u32 = 8u;

struct Params {
  a_rows: u32,
  a_cols: u32,
  b_cols: u32,
  _pad0: u32,
}

@group(0) @binding(0) var<storage, read> a: array<Scalar>;
@group(0) @binding(1) var<storage, read> b: array<Scalar>;
@group(0) @binding(2) var<storage, read_write> out: array<Scalar>;
@group(0) @binding(3) var<uniform> params: Params;

@compute @workgroup_size(TILE, TILE, 1)
fn main(@builtin(global_invocation_id) gid: vec3<u32>) {
  let col = gid.x;
  let row = gid.y;

  if (row >= params.a_rows || col >= params.b_cols) {
    return;
  }

  var acc: Scalar = 0.0;
  for (var k: u32 = 0u; k < params.a_cols; k = k + 1u) {
    acc = acc + (a[row * params.a_cols + k] * b[k * params.b_cols + col]);
  }

  out[row * params.b_cols + col] = acc;
}

