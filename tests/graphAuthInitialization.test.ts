import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";

const authUrl = new URL("../src/graphAuth.ts", import.meta.url).href;

function runWithPersistence(source: string, script: string, env = process.env) {
  const result = spawnSync(process.execPath, ["--input-type=module", "-e", `
    import { registerHooks } from "node:module";
    import assert from "node:assert/strict";
    registerHooks({
      resolve(specifier, context, nextResolve) {
        if (specifier === "@azure/msal-node-extensions") {
          return { url: "test:persistence", shortCircuit: true };
        }
        return nextResolve(specifier, context);
      },
      load(url, context, nextLoad) {
        if (url === "test:persistence") {
          return { format: "module", source: ${JSON.stringify(source)}, shortCircuit: true };
        }
        return nextLoad(url, context);
      }
    });
    const auth = await import(${JSON.stringify(authUrl)});
    ${script}
  `], { encoding: "utf8", env, timeout: 15_000 });
  assert.equal(result.error, undefined);
  assert.equal(result.status, 0, result.stderr);
}

const unavailablePersistence = `
  export const DataProtectionScope = {};
  export const PersistenceCreator = {};
  export class PersistenceCachePlugin {}
  throw new Error("Native persistence unavailable");
`;

test("token override works when native persistence cannot load", () => {
  runWithPersistence(unavailablePersistence, `
    assert.equal(await auth.resolveGraphAccessToken({ GRAPH_ACCESS_TOKEN: " manual-token " }), "manual-token");
  `);
});

test("unavailable native persistence produces a protected-cache error", () => {
  runWithPersistence(unavailablePersistence, `
    await assert.rejects(auth.createGraphAccessTokenProvider({ clientId: "test-client" }), (error) => {
      assert.ok(error instanceof auth.GraphAuthenticationError);
      assert.match(error.message, /protected.*cache/);
      assert.equal(error.cause.message, "Native persistence unavailable");
      return true;
    });
  `);
});

for (const customPath of [false, true]) {
  test(`creates missing ${customPath ? "custom" : "default"} cache parents before persistence`, async (t) => {
    const root = await mkdtemp(join(tmpdir(), "graph-auth-test-"));
    t.after(() => rm(root, { recursive: true, force: true }));
    const cachePath = customPath
      ? join(root, "custom", "nested", "cache.json")
      : join(root, "microsoft-todo-api", "msal-cache.json");
    const source = `
      import assert from "node:assert/strict";
      import { writeFile } from "node:fs/promises";
      export const DataProtectionScope = { CurrentUser: "current-user" };
      export class PersistenceCachePlugin {}
      export const PersistenceCreator = {
        async createPersistence(options) {
          assert.equal(options.cachePath, ${JSON.stringify(cachePath)});
          assert.equal(options.dataProtectionScope, "current-user");
          assert.equal(options.usePlaintextFileOnLinux, false);
          await writeFile(options.cachePath, "test-cache");
          return {};
        }
      };
    `;
    runWithPersistence(source, `
      const options = { clientId: "test-client", ${customPath ? `cachePath: ${JSON.stringify(cachePath)}` : ""} };
      await auth.createGraphAccessTokenProvider(options);
      await auth.createGraphAccessTokenProvider(options);
    `, { ...process.env, LOCALAPPDATA: root, XDG_CACHE_HOME: root });
  });
}
