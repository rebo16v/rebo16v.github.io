function sampleUniform(min, max) {
  const u = Math.random();
  return min + u * (max - min);
}

function sampleNormal(mean, stddev) {
  const u1 = Math.random();
  const u2 = Math.random();
  const z0 = Math.sqrt(-2.0 * Math.log(u1)) * Math.cos(2.0 * Math.PI * u2);
  return z0 * stddev + mean;
}

function sampleTriangular(a, b, c) {
  const u = Math.random();
  const F = (b-a)/(c-a);
  if (u < F) {
    return a + Match.sqrt(u * (b-a) * (c-a));
  } else {
    return b - Match.sqrt((1-u) * (b-a) * (b-c));
  }
}
