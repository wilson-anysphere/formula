/**
 * @param {ArrayLike<number>} vector
 * @returns {Float32Array}
 */
export function toFloat32Array(vector) {
  if (vector instanceof Float32Array) return vector;
  return Float32Array.from(vector);
}

/**
 * @param {ArrayLike<number>} vector
 * @returns {Float32Array}
 */
export function normalizeL2(vector) {
  const v = toFloat32Array(vector);
  let sumSq = 0;
  for (let i = 0; i < v.length; i += 1) sumSq += v[i] * v[i];
  const norm = Math.sqrt(sumSq) || 1;
  const out = new Float32Array(v.length);
  for (let i = 0; i < v.length; i += 1) out[i] = v[i] / norm;
  return out;
}

/**
 * Cosine similarity for L2-normalized vectors.
 * @param {ArrayLike<number>} a
 * @param {ArrayLike<number>} b
 */
export function cosineSimilarity(a, b) {
  if (a.length !== b.length) {
    throw new Error(
      `cosineSimilarity dimension mismatch: ${a.length} vs ${b.length}`
    );
  }
  let dot = 0;
  for (let i = 0; i < a.length; i += 1) dot += a[i] * b[i];
  return dot;
}
