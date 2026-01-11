import test from "node:test";
import assert from "node:assert/strict";

import semverCjs from "../shared/semver-range/index.js";
import sharedSemver from "../shared/semver.js";
import semverEsm, {
  compareSemver,
  isValidSemver,
  maxSemver,
  parseSemver,
  satisfies,
} from "../shared/semver-range/index.mjs";

test("semver-range: CJS + ESM entrypoints agree", () => {
  assert.equal(typeof semverCjs.compareSemver, "function");
  assert.equal(typeof semverCjs.satisfies, "function");

  assert.equal(semverCjs.compareSemver, semverEsm.compareSemver);
  assert.equal(semverCjs.satisfies, semverEsm.satisfies);

  // Backwards compatibility: marketplace imports historically used shared/semver.js.
  assert.equal(sharedSemver.compareSemver, semverEsm.compareSemver);
  assert.equal(sharedSemver.maxSemver, semverEsm.maxSemver);
});

test("semver-range: parseSemver supports prerelease/build metadata", () => {
  assert.deepEqual(parseSemver("1.2.3-alpha.1+build.9"), {
    major: 1,
    minor: 2,
    patch: 3,
    prerelease: ["alpha", "1"],
    build: ["build", "9"],
    raw: "1.2.3-alpha.1+build.9",
  });
});

test("semver-range: compareSemver handles prerelease precedence", () => {
  assert.equal(compareSemver("1.0.0-alpha", "1.0.0"), -1);
  assert.equal(compareSemver("1.0.0-alpha", "1.0.0-alpha"), 0);
  assert.equal(compareSemver("1.0.0-alpha.1", "1.0.0-alpha"), 1);
  assert.equal(compareSemver("1.0.0-alpha.1", "1.0.0-alpha.beta"), -1);
  assert.equal(compareSemver("1.0.0+build.1", "1.0.0+build.2"), 0);

  assert.throws(() => compareSemver("not-a-version", "1.0.0"), /Invalid semver compare/);
});

test("semver-range: satisfies supports wildcard, exact, comparisons, and compound AND ranges", () => {
  assert.equal(satisfies("1.2.3", "*"), true);
  assert.equal(satisfies("1.2.3", "1.2.3"), true);
  assert.equal(satisfies("1.2.3", "1.2.4"), false);

  assert.equal(satisfies("1.2.3", ">=1.0.0"), true);
  assert.equal(satisfies("1.0.0", ">=1.0.0"), true);
  assert.equal(satisfies("0.9.9", ">=1.0.0"), false);
  assert.equal(satisfies("1.0.1", ">1.0.0"), true);
  assert.equal(satisfies("1.0.0", ">1.0.0"), false);
  assert.equal(satisfies("1.0.0", "<=1.0.0"), true);
  assert.equal(satisfies("0.9.0", "<1.0.0"), true);

  assert.equal(satisfies("1.5.0", ">=1.0.0 <2.0.0"), true);
  assert.equal(satisfies("2.0.0", ">=1.0.0 <2.0.0"), false);
});

test("semver-range: satisfies caret/tilde semantics match common semver expectations", () => {
  assert.equal(satisfies("1.2.3", "^1.2.3"), true);
  assert.equal(satisfies("1.9.0", "^1.2.3"), true);
  assert.equal(satisfies("2.0.0", "^1.2.3"), false);

  // 0-major caret ranges are more restrictive.
  assert.equal(satisfies("0.2.3", "^0.2.3"), true);
  assert.equal(satisfies("0.2.9", "^0.2.3"), true);
  assert.equal(satisfies("0.3.0", "^0.2.3"), false);
  assert.equal(satisfies("0.0.4", "^0.0.3"), false);

  assert.equal(satisfies("1.2.3", "~1.2.3"), true);
  assert.equal(satisfies("1.2.9", "~1.2.3"), true);
  assert.equal(satisfies("1.3.0", "~1.2.3"), false);
});

test("semver-range: prereleases participate in comparisons", () => {
  assert.equal(satisfies("1.0.0-alpha", ">=1.0.0"), false);
  assert.equal(satisfies("1.0.0-alpha", "<1.0.0"), true);
});

test("semver-range: supports OR (||) ranges", () => {
  const range = ">=1.0.0 <2.0.0 || >=3.0.0";
  assert.equal(satisfies("1.5.0", range), true);
  assert.equal(satisfies("2.5.0", range), false);
  assert.equal(satisfies("3.0.0", range), true);
});

test("semver-range: invalid versions/ranges are rejected", () => {
  assert.equal(isValidSemver("not-a-version"), false);
  assert.equal(satisfies("not-a-version", "*"), false);
  assert.equal(satisfies("1.0.0", ""), false);
  assert.equal(satisfies("1.0.0", "not-a-range"), false);
});

test("semver-range: maxSemver respects prerelease ordering and matches compareSemver", () => {
  const versions = ["1.0.0-alpha", "1.0.0-alpha.1", "1.0.0"];
  assert.equal(maxSemver(versions), "1.0.0");

  const sorted = versions.slice().sort(compareSemver);
  assert.equal(maxSemver(versions), sorted[sorted.length - 1]);

  assert.equal(maxSemver(["not-a-version"]), null);
});
