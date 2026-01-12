import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { dirname, resolve } from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');

const hasBash = (() => {
  if (process.platform === 'win32') return false;
  const probe = spawnSync('bash', ['-lc', 'exit 0'], { stdio: 'ignore' });
  return probe.status === 0;
})();

function runBash(command) {
  const proc = spawnSync('bash', ['-lc', command], {
    encoding: 'utf8',
    cwd: repoRoot,
  });
  if (proc.error) throw proc.error;
  assert.equal(proc.status, 0, proc.stderr);
  return { stdout: proc.stdout, stderr: proc.stderr };
}

test('cargo_agent can run cargo fmt (subcommand does not accept -j)', { skip: !hasBash }, () => {
  const { stdout, stderr } = runBash('bash scripts/cargo_agent.sh fmt -- --version');
  assert.ok(!stderr.includes("unexpected argument '-j'"), stderr);
  assert.match(stdout, /rustfmt/i);
});

test('cargo_agent can run cargo clean (subcommand does not accept -j)', { skip: !hasBash }, () => {
  const { stderr } = runBash('bash scripts/cargo_agent.sh clean -n');
  assert.ok(!stderr.includes("unexpected argument '-j'"), stderr);
});

