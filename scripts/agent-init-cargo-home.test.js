import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { mkdtempSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, resolve } from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');

const hasBash = (() => {
  if (process.platform === 'win32') return false;
  const probe = spawnSync('bash', ['-lc', 'exit 0'], { stdio: 'ignore' });
  return probe.status === 0;
})();

const hasSh = (() => {
  if (process.platform === 'win32') return false;
  const probe = spawnSync('sh', ['-c', 'exit 0'], { stdio: 'ignore' });
  return probe.status === 0;
})();

function runBash(command) {
  const proc = spawnSync('bash', ['-lc', command], {
    encoding: 'utf8',
    cwd: repoRoot,
  });
  if (proc.error) throw proc.error;
  assert.equal(proc.status, 0, proc.stderr);
  return proc.stdout.trim();
}

function runSh(command) {
  const proc = spawnSync('sh', ['-c', command], {
    encoding: 'utf8',
    cwd: repoRoot,
  });
  if (proc.error) throw proc.error;
  assert.equal(proc.status, 0, proc.stderr);
  return { stdout: proc.stdout.trim(), stderr: proc.stderr.trim() };
}

test('agent-init warns when executed instead of sourced', { skip: !hasBash }, () => {
  const proc = spawnSync(
    'bash',
    [
      '-lc',
      [
        // Prevent agent-init from spawning Xvfb during this test.
        'export DISPLAY=:99',
        'bash scripts/agent-init.sh >/dev/null',
      ].join(' && '),
    ],
    { encoding: 'utf8', cwd: repoRoot },
  );
  if (proc.error) throw proc.error;
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(
    proc.stderr,
    /warning: scripts\/agent-init\.sh is meant to be sourced/,
    `expected a warning on stderr; got: ${proc.stderr}`,
  );
});

test('agent-init does not leak setup_display helper function (bash)', { skip: !hasBash }, () => {
  const out = runBash(
    [
      // Prevent agent-init from spawning Xvfb during this test.
      'export DISPLAY=:99',
      'source scripts/agent-init.sh >/dev/null',
      'type setup_display >/dev/null 2>&1 && echo leak || echo ok',
    ].join(' && '),
  );
  assert.equal(out, 'ok');
});

test('agent-init does not leak REPO_ROOT helper variable (bash)', { skip: !hasBash }, () => {
  const out = runBash(
    [
      'unset REPO_ROOT',
      // Prevent agent-init from spawning Xvfb during this test.
      'export DISPLAY=:99',
      'source scripts/agent-init.sh >/dev/null',
      // `${VAR+x}` expands to 'x' when VAR is set (even if empty), and empty when unset.
      'if [ -z "${REPO_ROOT+x}" ]; then echo ok; else echo leak; fi',
    ].join(' && '),
  );
  assert.equal(out, 'ok');
});

test('agent-init preserves existing NODE_OPTIONS flags while adding heap cap (bash)', { skip: !hasBash }, () => {
  const out = runBash(
    [
      'export NODE_OPTIONS="--trace-warnings"',
      // Prevent agent-init from spawning Xvfb during this test.
      'export DISPLAY=:99',
      'source scripts/agent-init.sh >/dev/null',
      'printf "%s" "$NODE_OPTIONS"',
    ].join(' && '),
  );
  assert.equal(out, '--max-old-space-size=3072 --trace-warnings');
});

test('agent-init does not duplicate NODE_OPTIONS heap cap when already set (bash)', { skip: !hasBash }, () => {
  const out = runBash(
    [
      'export NODE_OPTIONS="--max-old-space-size=4096 --trace-warnings"',
      // Prevent agent-init from spawning Xvfb during this test.
      'export DISPLAY=:99',
      'source scripts/agent-init.sh >/dev/null',
      'printf "%s" "$NODE_OPTIONS"',
    ].join(' && '),
  );
  assert.equal(out, '--max-old-space-size=4096 --trace-warnings');
});

test('agent-init defaults CARGO_HOME to a repo-local directory', { skip: !hasBash }, () => {
  const cargoHome = runBash(
    [
      'unset CARGO_HOME',
      // Prevent agent-init from spawning Xvfb during this test.
      'export DISPLAY=:99',
      'source scripts/agent-init.sh >/dev/null',
      'printf "%s" "$CARGO_HOME"',
    ].join(' && '),
  );

  assert.equal(cargoHome, resolve(repoRoot, 'target', 'cargo-home'));
});

test('agent-init treats $HOME/.cargo as default and overrides it', { skip: !hasBash }, () => {
  const cargoHome = runBash(
    [
      'unset CI',
      'export CARGO_HOME="$HOME/.cargo/"',
      'export DISPLAY=:99',
      'source scripts/agent-init.sh >/dev/null',
      'printf "%s" "$CARGO_HOME"',
    ].join(' && '),
  );

  assert.equal(cargoHome, resolve(repoRoot, 'target', 'cargo-home'));
});

test('agent-init can preserve $HOME/.cargo when FORMULA_ALLOW_GLOBAL_CARGO_HOME=1', { skip: !hasBash }, () => {
  const out = runBash(
    [
      'unset CI',
      'export CARGO_HOME="$HOME/.cargo"',
      'export FORMULA_ALLOW_GLOBAL_CARGO_HOME=1',
      'export DISPLAY=:99',
      'source scripts/agent-init.sh >/dev/null',
      'printf "%s\\n%s" "$HOME" "$CARGO_HOME"',
    ].join(' && '),
  );

  const [home, cargoHome] = out.split('\n');
  assert.equal(cargoHome, resolve(home, '.cargo'));
});

test('agent-init preserves $HOME/.cargo when running in CI', { skip: !hasBash }, () => {
  const out = runBash(
    [
      'export CI=1',
      'unset FORMULA_ALLOW_GLOBAL_CARGO_HOME',
      'export CARGO_HOME="$HOME/.cargo"',
      'export DISPLAY=:99',
      'source scripts/agent-init.sh >/dev/null',
      'printf "%s\\n%s" "$HOME" "$CARGO_HOME"',
    ].join(' && '),
  );

  const [home, cargoHome] = out.split('\n');
  assert.equal(cargoHome, resolve(home, '.cargo'));
});

test('agent-init preserves an existing CARGO_HOME override', { skip: !hasBash }, () => {
  const override = mkdtempSync(resolve(tmpdir(), 'formula-cargo-home-'));
  const cargoHome = runBash(
    [
      `export CARGO_HOME="${override}"`,
      'export DISPLAY=:99',
      'source scripts/agent-init.sh >/dev/null',
      'printf "%s" "$CARGO_HOME"',
    ].join(' && '),
  );

  assert.equal(cargoHome, override);
});

test('agent-init exports CARGO_HOME when set without export (bash)', { skip: !hasBash }, () => {
  const override = mkdtempSync(resolve(tmpdir(), 'formula-cargo-home-unexported-'));
  const out = runBash(
    [
      `CARGO_HOME="${override}"`,
      'export DISPLAY=:99',
      'source scripts/agent-init.sh >/dev/null',
      'env | grep "^CARGO_HOME="',
    ].join(' && '),
  );

  assert.equal(out, `CARGO_HOME=${override}`);
});

test('agent-init exports CARGO_BUILD_JOBS when set without export (bash)', { skip: !hasBash }, () => {
  const out = runBash(
    [
      'unset FORMULA_CARGO_JOBS',
      'unset CARGO_BUILD_JOBS',
      // Set but do not export (common interactive usage).
      'CARGO_BUILD_JOBS=7',
      'export DISPLAY=:99',
      'source scripts/agent-init.sh >/dev/null',
      'env | grep "^CARGO_BUILD_JOBS="',
    ].join(' && '),
  );

  assert.equal(out, 'CARGO_BUILD_JOBS=7');
});

test('agent-init exports MAKEFLAGS when set without export (bash)', { skip: !hasBash }, () => {
  const out = runBash(
    [
      'unset MAKEFLAGS',
      // Set but do not export (common interactive usage).
      'MAKEFLAGS="-j9 -l2"',
      'export DISPLAY=:99',
      'source scripts/agent-init.sh >/dev/null',
      'env | grep "^MAKEFLAGS="',
    ].join(' && '),
  );

  assert.equal(out, 'MAKEFLAGS=-j9 -l2');
});

test('agent-init exports RAYON_NUM_THREADS when set without export (bash)', { skip: !hasBash }, () => {
  const out = runBash(
    [
      'unset RAYON_NUM_THREADS',
      // Set but do not export (common interactive usage).
      'RAYON_NUM_THREADS=11',
      'export DISPLAY=:99',
      'source scripts/agent-init.sh >/dev/null',
      'env | grep "^RAYON_NUM_THREADS="',
    ].join(' && '),
  );

  assert.equal(out, 'RAYON_NUM_THREADS=11');
});

test('agent-init prepends CARGO_HOME/bin to PATH', { skip: !hasBash }, () => {
  const out = runBash(
    [
      'unset CARGO_HOME',
      // Stabilize PATH so the assertion isn't affected by login shell defaults.
      'export PATH="/usr/bin:/bin"',
      'export DISPLAY=:99',
      'source scripts/agent-init.sh >/dev/null',
      'printf "%s\\n%s" "$CARGO_HOME" "$PATH"',
    ].join(' && '),
  );

  const [cargoHome, pathValue] = out.split('\n');
  assert.equal(cargoHome, resolve(repoRoot, 'target', 'cargo-home'));
  assert.ok(pathValue.split(':')[0] === resolve(repoRoot, 'target', 'cargo-home', 'bin'));
});

test('agent-init derives CARGO_BUILD_JOBS from FORMULA_CARGO_JOBS', { skip: !hasBash }, () => {
  const out = runBash(
    [
      'unset CARGO_HOME',
      'unset CARGO_BUILD_JOBS',
      'unset MAKEFLAGS',
      'unset RAYON_NUM_THREADS',
      'export FORMULA_CARGO_JOBS=7',
      // Prevent agent-init from spawning Xvfb during this test.
      'export DISPLAY=:99',
      'source scripts/agent-init.sh >/dev/null',
      'printf "%s\\n%s\\n%s" "$CARGO_BUILD_JOBS" "$MAKEFLAGS" "$RAYON_NUM_THREADS"',
    ].join(' && '),
  );

  const [jobs, makeflags, rayon] = out.split('\n');
  assert.equal(jobs, '7');
  assert.equal(makeflags, '-j7');
  assert.equal(rayon, '7');
});

test(
  'agent-init updates MAKEFLAGS and RAYON_NUM_THREADS when re-sourced after changing FORMULA_CARGO_JOBS',
  { skip: !hasBash },
  () => {
    const out = runBash(
      [
        'unset MAKEFLAGS',
        'unset RAYON_NUM_THREADS',
        'unset CARGO_BUILD_JOBS',
        'export DISPLAY=:99',
        'export FORMULA_CARGO_JOBS=7',
        'source scripts/agent-init.sh >/dev/null',
        'first=$(printf "%s,%s,%s" "$CARGO_BUILD_JOBS" "$MAKEFLAGS" "$RAYON_NUM_THREADS")',
        'export FORMULA_CARGO_JOBS=3',
        'source scripts/agent-init.sh >/dev/null',
        'second=$(printf "%s,%s,%s" "$CARGO_BUILD_JOBS" "$MAKEFLAGS" "$RAYON_NUM_THREADS")',
        'printf "%s\\n%s" "$first" "$second"',
      ].join(' && '),
    );

    const [first, second] = out.split('\n');
    assert.equal(first, '7,-j7,7');
    assert.equal(second, '3,-j3,3');
  },
);

test('agent-init exports FORMULA_CARGO_JOBS when set without export', { skip: !hasBash }, () => {
  const out = runBash(
    [
      'unset FORMULA_CARGO_JOBS',
      // Set but do not export (common interactive usage).
      'FORMULA_CARGO_JOBS=7',
      'export DISPLAY=:99',
      'source scripts/agent-init.sh >/dev/null',
      'env | grep "^FORMULA_CARGO_JOBS="',
    ].join(' && '),
  );
  assert.equal(out, 'FORMULA_CARGO_JOBS=7');
});

test('agent-init exports FORMULA_ALLOW_GLOBAL_CARGO_HOME when set without export', { skip: !hasBash }, () => {
  const out = runBash(
    [
      'unset FORMULA_ALLOW_GLOBAL_CARGO_HOME',
      // Set but do not export (common interactive usage).
      'FORMULA_ALLOW_GLOBAL_CARGO_HOME=1',
      'export DISPLAY=:99',
      'source scripts/agent-init.sh >/dev/null',
      'env | grep "^FORMULA_ALLOW_GLOBAL_CARGO_HOME="',
    ].join(' && '),
  );
  assert.equal(out, 'FORMULA_ALLOW_GLOBAL_CARGO_HOME=1');
});

test('agent-init can be sourced with nounset enabled and DISPLAY unset (bash)', { skip: !hasBash }, () => {
  const stubDir = mkdtempSync(resolve(tmpdir(), 'formula-xvfb-stub-'));
  const stubPath = resolve(stubDir, 'Xvfb');
  writeFileSync(stubPath, '#!/bin/sh\nexit 0\n', { mode: 0o755 });

  const out = runBash(
    [
      `export PATH="${stubDir}:$PATH"`,
      'set -u',
      'unset DISPLAY',
      'source scripts/agent-init.sh >/dev/null',
      'printf "ok"',
    ].join(' && '),
  );

  assert.equal(out, 'ok');
});

test('agent-init can be sourced with nounset enabled and IFS unset (bash)', { skip: !hasBash }, () => {
  const stubDir = mkdtempSync(resolve(tmpdir(), 'formula-xvfb-stub-'));
  const stubPath = resolve(stubDir, 'Xvfb');
  writeFileSync(stubPath, '#!/bin/sh\nexit 0\n', { mode: 0o755 });

  const out = runBash(
    [
      `export PATH="${stubDir}:$PATH"`,
      'set -u',
      'unset DISPLAY',
      'unset IFS',
      'source scripts/agent-init.sh >/dev/null',
      'if [ -z "${IFS+x}" ]; then printf "ok"; else printf "leak"; fi',
    ].join(' && '),
  );

  assert.equal(out, 'ok');
});

test('agent-init restores IFS when it is set to a custom value (bash)', { skip: !hasBash }, () => {
  const out = runBash(
    [
      // Prevent agent-init from spawning Xvfb during this test.
      'export DISPLAY=:99',
      'IFS=","',
      'source scripts/agent-init.sh >/dev/null',
      'printf "%s" "$IFS"',
    ].join(' && '),
  );

  assert.equal(out, ',');
});

test('agent-init does not enable errexit in bash when it was previously disabled', { skip: !hasBash }, () => {
  const out = runBash(
    [
      'set +e',
      'export DISPLAY=:99',
      'before=$-',
      'source scripts/agent-init.sh >/dev/null',
      'after=$-',
      'printf "%s\\n%s" "$before" "$after"',
    ].join(' && '),
  );

  const [before, after] = out.split('\n');
  assert.ok(!before.includes('e'), `expected errexit disabled before sourcing; got $-=${before}`);
  assert.ok(!after.includes('e'), `expected errexit disabled after sourcing; got $-=${after}`);
});

test('agent-init preserves errexit in bash when it was previously enabled', { skip: !hasBash }, () => {
  const out = runBash(
    [
      'set -e',
      'export DISPLAY=:99',
      'before=$-',
      'source scripts/agent-init.sh >/dev/null',
      'after=$-',
      'printf "%s\\n%s" "$before" "$after"',
    ].join(' && '),
  );

  const [before, after] = out.split('\n');
  assert.ok(before.includes('e'), `expected errexit enabled before sourcing; got $-=${before}`);
  assert.ok(after.includes('e'), `expected errexit enabled after sourcing; got $-=${after}`);
});

test(
  'agent-init treats CARGO_HOME=$HOME/.cargo as unset in local runs (defaults to repo-local cargo-home)',
  { skip: !hasBash },
  () => {
    const fakeHome = mkdtempSync(resolve(tmpdir(), 'formula-home-'));
    const cargoHome = runBash(
      [
        `export HOME="${fakeHome}"`,
        'export CARGO_HOME="$HOME/.cargo/"',
        'unset CI',
        'unset FORMULA_ALLOW_GLOBAL_CARGO_HOME',
        'export DISPLAY=:99',
        'source scripts/agent-init.sh >/dev/null',
        'printf "%s" "$CARGO_HOME"',
      ].join(' && '),
    );

    assert.equal(cargoHome, resolve(repoRoot, 'target', 'cargo-home'));
  },
);

test(
  'agent-init preserves CARGO_HOME=$HOME/.cargo when FORMULA_ALLOW_GLOBAL_CARGO_HOME=1',
  { skip: !hasBash },
  () => {
    const fakeHome = mkdtempSync(resolve(tmpdir(), 'formula-home-'));
    const cargoHome = runBash(
      [
        `export HOME="${fakeHome}"`,
        'export CARGO_HOME="$HOME/.cargo"',
        'unset CI',
        'export FORMULA_ALLOW_GLOBAL_CARGO_HOME=1',
        'export DISPLAY=:99',
        'source scripts/agent-init.sh >/dev/null',
        'printf "%s" "$CARGO_HOME"',
      ].join(' && '),
    );

    assert.equal(cargoHome, resolve(fakeHome, '.cargo'));
  },
);

test('agent-init preserves CARGO_HOME=$HOME/.cargo when CI is set', { skip: !hasBash }, () => {
  const fakeHome = mkdtempSync(resolve(tmpdir(), 'formula-home-'));
  const cargoHome = runBash(
    [
      `export HOME="${fakeHome}"`,
      'export CARGO_HOME="$HOME/.cargo"',
      'export CI=1',
      'unset FORMULA_ALLOW_GLOBAL_CARGO_HOME',
      'export DISPLAY=:99',
      'source scripts/agent-init.sh >/dev/null',
      'printf "%s" "$CARGO_HOME"',
    ].join(' && '),
  );

  assert.equal(cargoHome, resolve(fakeHome, '.cargo'));
});

test('agent-init can be sourced under /bin/sh (no bash-only syntax)', { skip: !hasSh }, () => {
  const { stdout, stderr } = runSh(
    [
      'unset CARGO_HOME',
      // Prevent agent-init from spawning Xvfb during this test.
      'export DISPLAY=:99',
      '. scripts/agent-init.sh >/dev/null',
      'printf "%s" "$CARGO_HOME"',
    ].join(' && '),
  );

  assert.equal(stderr, '');
  assert.equal(stdout, resolve(repoRoot, 'target', 'cargo-home'));
});

test('agent-init does not leak REPO_ROOT helper variable under /bin/sh', { skip: !hasSh }, () => {
  const { stdout, stderr } = runSh(
    [
      'unset REPO_ROOT',
      // Prevent agent-init from spawning Xvfb during this test.
      'export DISPLAY=:99',
      '. scripts/agent-init.sh >/dev/null',
      'if [ -z "${REPO_ROOT+x}" ]; then echo ok; else echo leak; fi',
    ].join(' && '),
  );

  assert.equal(stderr, '');
  assert.equal(stdout, 'ok');
});

test(
  'agent-init preserves existing NODE_OPTIONS flags while adding heap cap under /bin/sh',
  { skip: !hasSh },
  () => {
    const { stdout, stderr } = runSh(
      [
        'export NODE_OPTIONS="--trace-warnings"',
        // Prevent agent-init from spawning Xvfb during this test.
        'export DISPLAY=:99',
        '. scripts/agent-init.sh >/dev/null',
        'printf "%s" "$NODE_OPTIONS"',
      ].join(' && '),
    );

    assert.equal(stderr, '');
    assert.equal(stdout, '--max-old-space-size=3072 --trace-warnings');
  },
);

test('agent-init does not duplicate NODE_OPTIONS heap cap under /bin/sh', { skip: !hasSh }, () => {
  const { stdout, stderr } = runSh(
    [
      'export NODE_OPTIONS="--max-old-space-size=4096 --trace-warnings"',
      // Prevent agent-init from spawning Xvfb during this test.
      'export DISPLAY=:99',
      '. scripts/agent-init.sh >/dev/null',
      'printf "%s" "$NODE_OPTIONS"',
    ].join(' && '),
  );

  assert.equal(stderr, '');
  assert.equal(stdout, '--max-old-space-size=4096 --trace-warnings');
});

test('agent-init exports CARGO_HOME when set without export under /bin/sh', { skip: !hasSh }, () => {
  const override = mkdtempSync(resolve(tmpdir(), 'formula-cargo-home-unexported-'));
  const { stdout, stderr } = runSh(
    [
      `CARGO_HOME="${override}"`,
      'export DISPLAY=:99',
      '. scripts/agent-init.sh >/dev/null',
      'env | grep "^CARGO_HOME="',
    ].join(' && '),
  );

  assert.equal(stderr, '');
  assert.equal(stdout, `CARGO_HOME=${override}`);
});

test('agent-init exports CARGO_BUILD_JOBS when set without export under /bin/sh', { skip: !hasSh }, () => {
  const { stdout, stderr } = runSh(
    [
      'unset FORMULA_CARGO_JOBS',
      'unset CARGO_BUILD_JOBS',
      // Set but do not export (common interactive usage).
      'CARGO_BUILD_JOBS=7',
      'export DISPLAY=:99',
      '. scripts/agent-init.sh >/dev/null',
      'env | grep "^CARGO_BUILD_JOBS="',
    ].join(' && '),
  );

  assert.equal(stderr, '');
  assert.equal(stdout, 'CARGO_BUILD_JOBS=7');
});

test('agent-init exports MAKEFLAGS when set without export under /bin/sh', { skip: !hasSh }, () => {
  const { stdout, stderr } = runSh(
    [
      'unset MAKEFLAGS',
      // Set but do not export (common interactive usage).
      'MAKEFLAGS="-j9 -l2"',
      'export DISPLAY=:99',
      '. scripts/agent-init.sh >/dev/null',
      'env | grep "^MAKEFLAGS="',
    ].join(' && '),
  );

  assert.equal(stderr, '');
  assert.equal(stdout, 'MAKEFLAGS=-j9 -l2');
});

test('agent-init exports RAYON_NUM_THREADS when set without export under /bin/sh', { skip: !hasSh }, () => {
  const { stdout, stderr } = runSh(
    [
      'unset RAYON_NUM_THREADS',
      // Set but do not export (common interactive usage).
      'RAYON_NUM_THREADS=11',
      'export DISPLAY=:99',
      '. scripts/agent-init.sh >/dev/null',
      'env | grep "^RAYON_NUM_THREADS="',
    ].join(' && '),
  );

  assert.equal(stderr, '');
  assert.equal(stdout, 'RAYON_NUM_THREADS=11');
});

test('agent-init exports FORMULA_CARGO_JOBS when set without export under /bin/sh', { skip: !hasSh }, () => {
  const { stdout, stderr } = runSh(
    [
      'unset FORMULA_CARGO_JOBS',
      'FORMULA_CARGO_JOBS=7',
      'export DISPLAY=:99',
      '. scripts/agent-init.sh >/dev/null',
      'env | grep "^FORMULA_CARGO_JOBS="',
    ].join(' && '),
  );

  assert.equal(stderr, '');
  assert.equal(stdout, 'FORMULA_CARGO_JOBS=7');
});

test(
  'agent-init updates MAKEFLAGS and RAYON_NUM_THREADS when re-sourced after changing FORMULA_CARGO_JOBS under /bin/sh',
  { skip: !hasSh },
  () => {
    const { stdout, stderr } = runSh(
      [
        'unset MAKEFLAGS',
        'unset RAYON_NUM_THREADS',
        'unset CARGO_BUILD_JOBS',
        'export DISPLAY=:99',
        'FORMULA_CARGO_JOBS=7',
        '. scripts/agent-init.sh >/dev/null',
        'first=$(printf "%s,%s,%s" "$CARGO_BUILD_JOBS" "$MAKEFLAGS" "$RAYON_NUM_THREADS")',
        'FORMULA_CARGO_JOBS=3',
        '. scripts/agent-init.sh >/dev/null',
        'second=$(printf "%s,%s,%s" "$CARGO_BUILD_JOBS" "$MAKEFLAGS" "$RAYON_NUM_THREADS")',
        'printf "%s\\n%s" "$first" "$second"',
      ].join(' && '),
    );

    assert.equal(stderr, '');
    const [first, second] = stdout.split('\n');
    assert.equal(first, '7,-j7,7');
    assert.equal(second, '3,-j3,3');
  },
);

test('agent-init can be sourced with nounset enabled and DISPLAY unset under /bin/sh', { skip: !hasSh }, () => {
  const stubDir = mkdtempSync(resolve(tmpdir(), 'formula-xvfb-stub-'));
  const stubPath = resolve(stubDir, 'Xvfb');
  writeFileSync(stubPath, '#!/bin/sh\nexit 0\n', { mode: 0o755 });

  const { stdout, stderr } = runSh(
    [
      `export PATH="${stubDir}:$PATH"`,
      'set -u',
      'unset DISPLAY',
      '. scripts/agent-init.sh >/dev/null',
      'printf "ok"',
    ].join(' && '),
  );

  assert.equal(stderr, '');
  assert.equal(stdout, 'ok');
});

test('agent-init can be sourced with nounset enabled and IFS unset under /bin/sh', { skip: !hasSh }, () => {
  const stubDir = mkdtempSync(resolve(tmpdir(), 'formula-xvfb-stub-'));
  const stubPath = resolve(stubDir, 'Xvfb');
  writeFileSync(stubPath, '#!/bin/sh\nexit 0\n', { mode: 0o755 });

  const { stdout, stderr } = runSh(
    [
      `export PATH="${stubDir}:$PATH"`,
      'set -u',
      'unset DISPLAY',
      'unset IFS',
      '. scripts/agent-init.sh >/dev/null',
      'if [ -z "${IFS+x}" ]; then printf "ok"; else printf "leak"; fi',
    ].join(' && '),
  );

  assert.equal(stderr, '');
  assert.equal(stdout, 'ok');
});

test('agent-init restores IFS when it is set to a custom value under /bin/sh', { skip: !hasSh }, () => {
  const { stdout, stderr } = runSh(
    [
      'export DISPLAY=:99',
      'IFS=","',
      '. scripts/agent-init.sh >/dev/null',
      'printf "%s" "$IFS"',
    ].join(' && '),
  );

  assert.equal(stderr, '');
  assert.equal(stdout, ',');
});

test('agent-init does not enable errexit in /bin/sh when it was previously disabled', { skip: !hasSh }, () => {
  const { stdout, stderr } = runSh(
    [
      'set +e',
      'export DISPLAY=:99',
      'before=$-',
      '. scripts/agent-init.sh >/dev/null',
      'after=$-',
      'printf "before=%s\\nafter=%s" "$before" "$after"',
    ].join(' && '),
  );

  assert.equal(stderr, '');
  const [beforeLine, afterLine] = stdout.split('\n');
  const before = beforeLine.replace(/^before=/, '');
  const after = afterLine.replace(/^after=/, '');
  assert.ok(!before.includes('e'), `expected errexit disabled before sourcing; got $-=${before}`);
  assert.ok(!after.includes('e'), `expected errexit disabled after sourcing; got $-=${after}`);
});

test('agent-init preserves errexit in /bin/sh when it was previously enabled', { skip: !hasSh }, () => {
  const { stdout, stderr } = runSh(
    [
      'set -e',
      'export DISPLAY=:99',
      'before=$-',
      '. scripts/agent-init.sh >/dev/null',
      'after=$-',
      'printf "before=%s\\nafter=%s" "$before" "$after"',
    ].join(' && '),
  );

  assert.equal(stderr, '');
  const [beforeLine, afterLine] = stdout.split('\n');
  const before = beforeLine.replace(/^before=/, '');
  const after = afterLine.replace(/^after=/, '');
  assert.ok(before.includes('e'), `expected errexit enabled before sourcing; got $-=${before}`);
  assert.ok(after.includes('e'), `expected errexit enabled after sourcing; got $-=${after}`);
});
