import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import http from "node:http";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

// Note: this file uses the `.test.js` extension so it is picked up by `pnpm test:node`
// (scripts/run-node-tests.mjs collects `*.test.js` suites).

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const scriptPath = path.join(repoRoot, "scripts", "check-tag-ci-status.mjs");

/**
 * @param {string[]} args
 * @param {Record<string, string | undefined>} env
 */
function runScript(args, env) {
  return new Promise((resolve, reject) => {
    const child = spawn(process.execPath, [scriptPath, ...args], {
      cwd: repoRoot,
      env: { ...process.env, ...env },
      stdio: ["ignore", "pipe", "pipe"],
    });

    /** @type {Buffer[]} */
    const stdout = [];
    /** @type {Buffer[]} */
    const stderr = [];
    child.stdout.on("data", (chunk) => stdout.push(Buffer.from(chunk)));
    child.stderr.on("data", (chunk) => stderr.push(Buffer.from(chunk)));

    child.on("error", reject);
    child.on("close", (code, signal) => {
      resolve({
        code: code ?? 0,
        signal,
        stdout: Buffer.concat(stdout).toString("utf8"),
        stderr: Buffer.concat(stderr).toString("utf8"),
      });
    });
  });
}

/**
 * @param {{
 *   workflows: Array<{ id: number; name: string; path: string }>;
 *   runsByWorkflowId: Record<string, any[]>;
 * }} config
 */
async function withMockGitHub(config, fn) {
  const server = http.createServer((req, res) => {
    try {
      const url = new URL(req.url ?? "/", "http://127.0.0.1");

      if (url.pathname === "/repos/acme/widgets/actions/workflows") {
        const page = Number.parseInt(url.searchParams.get("page") ?? "1", 10);
        const workflows =
          "workflowsByPage" in config && config.workflowsByPage
            ? config.workflowsByPage[String(page)] ?? []
            : config.workflows;
        res.writeHead(200, { "Content-Type": "application/json" });
        res.end(JSON.stringify({ workflows }));
        return;
      }

      const runsMatch = url.pathname.match(
        /^\/repos\/acme\/widgets\/actions\/workflows\/(\d+)\/runs$/,
      );
      if (runsMatch) {
        const id = runsMatch[1];
        res.writeHead(200, { "Content-Type": "application/json" });
        res.end(JSON.stringify({ workflow_runs: config.runsByWorkflowId[id] ?? [] }));
        return;
      }

      res.writeHead(404, { "Content-Type": "application/json" });
      res.end(JSON.stringify({ message: `Unhandled path: ${url.pathname}` }));
    } catch (err) {
      res.writeHead(500, { "Content-Type": "application/json" });
      res.end(JSON.stringify({ message: String(err) }));
    }
  });

  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  const addr = server.address();
  assert.ok(addr && typeof addr === "object");
  const baseUrl = `http://127.0.0.1:${addr.port}`;

  try {
    return await fn(baseUrl);
  } finally {
    await new Promise((resolve) => server.close(resolve));
  }
}

test("finds the CI workflow even when it is on a later workflows API page", async () => {
  const sha = "abababababababababababababababababababab";

  const filler = Array.from({ length: 100 }, (_, i) => ({
    id: i + 1,
    name: `Workflow ${i + 1}`,
    path: `.github/workflows/${i + 1}.yml`,
  }));

  await withMockGitHub(
    {
      workflowsByPage: {
        "1": filler,
        "2": [{ id: 1234, name: "CI", path: ".github/workflows/ci.yml" }],
      },
      workflows: [],
      runsByWorkflowId: {
        "1234": [
          {
            id: 1,
            head_sha: sha,
            status: "completed",
            conclusion: "success",
            html_url: "http://example.local/runs/1",
          },
        ],
      },
    },
    async (baseUrl) => {
      const result = await runScript(["--repo", "acme/widgets", "--sha", sha, "--workflow", "CI"], {
        GITHUB_TOKEN: "test-token",
        GITHUB_API_URL: baseUrl,
        GITHUB_SERVER_URL: "http://example.local",
      });

      assert.equal(result.code, 0, `expected exit 0\nstdout:\n${result.stdout}\nstderr:\n${result.stderr}`);
      assert.match(result.stdout, /CI status check passed/i);
    },
  );
});

test("passes when CI has a successful completed run for the given commit", async () => {
  const sha = "deadbeefdeadbeefdeadbeefdeadbeefdeadbeef";

  await withMockGitHub(
    {
      workflows: [{ id: 123, name: "CI", path: ".github/workflows/ci.yml" }],
      runsByWorkflowId: {
        "123": [
          {
            id: 1,
            head_sha: sha,
            status: "completed",
            conclusion: "failure",
            html_url: "http://example.local/runs/1",
          },
          {
            id: 2,
            head_sha: sha,
            status: "completed",
            conclusion: "success",
            html_url: "http://example.local/runs/2",
          },
        ],
      },
    },
    async (baseUrl) => {
      const result = await runScript(["--repo", "acme/widgets", "--sha", sha, "--workflow", "CI"], {
        GITHUB_TOKEN: "test-token",
        GITHUB_API_URL: baseUrl,
        GITHUB_SERVER_URL: "http://example.local",
      });

      assert.equal(result.code, 0, `expected exit 0\nstdout:\n${result.stdout}\nstderr:\n${result.stderr}`);
      assert.match(result.stdout, /CI status check passed/i);
    },
  );
});

test("fails with a clear error when no successful CI run exists", async () => {
  const sha = "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa";

  await withMockGitHub(
    {
      workflows: [{ id: 123, name: "CI", path: ".github/workflows/ci.yml" }],
      runsByWorkflowId: {
        "123": [
          {
            id: 42,
            head_sha: sha,
            status: "completed",
            conclusion: "failure",
            html_url: "http://example.local/runs/42",
          },
        ],
      },
    },
    async (baseUrl) => {
      const result = await runScript(["--repo", "acme/widgets", "--sha", sha, "--workflow", "CI"], {
        GITHUB_TOKEN: "test-token",
        GITHUB_API_URL: baseUrl,
        GITHUB_SERVER_URL: "http://example.local",
      });

      assert.notEqual(result.code, 0, "expected non-zero exit code");
      assert.match(result.stderr, /no successful "CI" workflow run found/i);
      assert.match(result.stderr, new RegExp(sha));
      assert.match(result.stderr, /Newest completed run: conclusion=failure/i);
    },
  );
});

test("fails when a workflow named CI exists but is not the main ci.yml workflow", async () => {
  const sha = "bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb";

  await withMockGitHub(
    {
      workflows: [{ id: 123, name: "CI", path: ".github/workflows/other.yml" }],
      runsByWorkflowId: { "123": [] },
    },
    async (baseUrl) => {
      const result = await runScript(["--repo", "acme/widgets", "--sha", sha, "--workflow", "CI"], {
        GITHUB_TOKEN: "test-token",
        GITHUB_API_URL: baseUrl,
        GITHUB_SERVER_URL: "http://example.local",
      });

      assert.notEqual(result.code, 0, "expected non-zero exit code");
      assert.match(result.stderr, /intentionally restricted/i);
      assert.match(result.stderr, /\.github\/workflows\/ci\.yml/);
    },
  );
});

test("ignores successful runs for other commits (defensive head_sha validation)", async () => {
  const sha = "cccccccccccccccccccccccccccccccccccccccc";
  const otherSha = "dddddddddddddddddddddddddddddddddddddddd";

  await withMockGitHub(
    {
      workflows: [{ id: 123, name: "CI", path: ".github/workflows/ci.yml" }],
      runsByWorkflowId: {
        "123": [
          {
            id: 999,
            head_sha: otherSha,
            status: "completed",
            conclusion: "success",
            html_url: "http://example.local/runs/999",
          },
        ],
      },
    },
    async (baseUrl) => {
      const result = await runScript(["--repo", "acme/widgets", "--sha", sha, "--workflow", "CI"], {
        GITHUB_TOKEN: "test-token",
        GITHUB_API_URL: baseUrl,
        GITHUB_SERVER_URL: "http://example.local",
      });

      assert.notEqual(result.code, 0, "expected non-zero exit code");
      assert.match(result.stderr, /no successful "CI" workflow run found/i);
      assert.match(result.stderr, /shows no workflow runs/i);
    },
  );
});

test("reports when CI is still running (newest run not completed)", async () => {
  const sha = "eeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeee";

  await withMockGitHub(
    {
      workflows: [{ id: 123, name: "CI", path: ".github/workflows/ci.yml" }],
      runsByWorkflowId: {
        "123": [
          {
            id: 7,
            head_sha: sha,
            status: "in_progress",
            conclusion: null,
            html_url: "http://example.local/runs/7",
          },
        ],
      },
    },
    async (baseUrl) => {
      const result = await runScript(["--repo", "acme/widgets", "--sha", sha, "--workflow", "CI"], {
        GITHUB_TOKEN: "test-token",
        GITHUB_API_URL: baseUrl,
        GITHUB_SERVER_URL: "http://example.local",
      });

      assert.notEqual(result.code, 0, "expected non-zero exit code");
      assert.match(result.stderr, /no successful "CI" workflow run found/i);
      assert.match(result.stderr, /Newest run: status=in_progress/i);
    },
  );
});

test("includes a permissions hint when GitHub returns 403", async () => {
  const sha = "ffffffffffffffffffffffffffffffffffffffff";

  const server = http.createServer((req, res) => {
    const url = new URL(req.url ?? "/", "http://127.0.0.1");
    if (url.pathname === "/repos/acme/widgets/actions/workflows") {
      res.writeHead(403, { "Content-Type": "application/json" });
      res.end(JSON.stringify({ message: "Forbidden" }));
      return;
    }

    res.writeHead(404, { "Content-Type": "application/json" });
    res.end(JSON.stringify({ message: `Unhandled path: ${url.pathname}` }));
  });

  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  const addr = server.address();
  assert.ok(addr && typeof addr === "object");
  const baseUrl = `http://127.0.0.1:${addr.port}`;

  try {
    const result = await runScript(["--repo", "acme/widgets", "--sha", sha, "--workflow", "CI"], {
      GITHUB_TOKEN: "test-token",
      GITHUB_API_URL: baseUrl,
      GITHUB_SERVER_URL: "http://example.local",
    });

    assert.notEqual(result.code, 0, "expected non-zero exit code");
    assert.match(result.stderr, /403/i);
    assert.match(result.stderr, /actions: read/i);
  } finally {
    await new Promise((resolve) => server.close(resolve));
  }
});
