alias Scalar = f32;

const WG_SIZE: u32 = 256u;

struct Params {
  min: Scalar,
  max: Scalar,
  inv_bin_width: Scalar,
  length: u32,
  bin_count: u32,
  _pad0: u32,
  _pad1: u32,
  _pad2: u32,
}

@group(0) @binding(0) var<storage, read> input: array<Scalar>;
@group(0) @binding(1) var<storage, read_write> bins: array<atomic<u32>>;
@group(0) @binding(2) var<uniform> params: Params;

@compute @workgroup_size(WG_SIZE)
fn main(@builtin(global_invocation_id) gid: vec3<u32>, @builtin(num_workgroups) nwg: vec3<u32>) {
  let stride = nwg.x * WG_SIZE;
  let i = gid.x + gid.y * stride;
  if (i >= params.length) {
    return;
  }

  let v = input[i];
  if (isNan(v)) {
    return;
  }
  var bin_i: i32;
  if (v <= params.min) {
    bin_i = 0;
  } else if (v >= params.max) {
    bin_i = i32(params.bin_count) - 1;
  } else {
    let scaled = (v - params.min) * params.inv_bin_width;
    // Guard extreme range configurations like min=-Infinity where
    // (v - min) * inv_bin_width can become NaN (Infinity * 0).
    if (isNan(scaled)) {
      bin_i = 0;
    } else {
      bin_i = i32(scaled);
    }
  }
  if (bin_i < 0) {
    bin_i = 0;
  }
  if (bin_i >= i32(params.bin_count)) {
    bin_i = i32(params.bin_count) - 1;
  }

  atomicAdd(&bins[u32(bin_i)], 1u);
}
