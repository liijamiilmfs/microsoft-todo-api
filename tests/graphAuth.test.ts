import assert from "node:assert/strict";
import test from "node:test";

import { InteractionRequiredAuthError } from "@azure/msal-node";

import {
  GRAPH_TASKS_SCOPE,
  GraphAuthenticationError,
  MsalGraphTokenProvider,
  resolveGraphAccessToken,
  type MsalTokenClient
} from "../src/graphAuth.ts";

function authenticationResult(accessToken: string) {
  return { accessToken } as never;
}

test("rejects multiple cached accounts before requesting any token", async () => {
  const client: MsalTokenClient = {
    getTokenCache: () => ({
      getAllAccounts: async () => [
        { homeAccountId: "account-1" },
        { homeAccountId: "account-2" }
      ] as never
    }),
    acquireTokenSilent: async () => assert.fail("Must not select a cached account"),
    acquireTokenByDeviceCode: async () => assert.fail("Must not start sign-in")
  };
  const provider = new MsalGraphTokenProvider(client, () => undefined);

  await assert.rejects(() => provider.getAccessToken(), (error: unknown) => {
    assert.ok(error instanceof GraphAuthenticationError);
    assert.match(error.message, /multiple.*accounts/i);
    return true;
  });
});

test("returns a silently acquired cached token without starting device code", async () => {
  let deviceCodeCalls = 0;
  let silentScopes: string[] = [];
  const client: MsalTokenClient = {
    getTokenCache: () => ({
      getAllAccounts: async () => [{ homeAccountId: "account-1" }] as never
    }),
    acquireTokenSilent: async (request) => {
      silentScopes = request.scopes;
      return authenticationResult("cached-token");
    },
    acquireTokenByDeviceCode: async () => {
      deviceCodeCalls += 1;
      return authenticationResult("device-token");
    }
  };

  const provider = new MsalGraphTokenProvider(client, () => undefined);

  assert.equal(await provider.getAccessToken(), "cached-token");
  assert.deepEqual(silentScopes, [GRAPH_TASKS_SCOPE]);
  assert.equal(deviceCodeCalls, 0);
});

test("uses device code when no cached account exists", async () => {
  let callbackMessage = "";
  let deviceCodeCalls = 0;
  const client: MsalTokenClient = {
    getTokenCache: () => ({ getAllAccounts: async () => [] }),
    acquireTokenSilent: async () => authenticationResult("unexpected-token"),
    acquireTokenByDeviceCode: async (request) => {
      deviceCodeCalls += 1;
      request.deviceCodeCallback({ message: "Use the Microsoft device code." } as never);
      assert.deepEqual(request.scopes, [GRAPH_TASKS_SCOPE]);
      return authenticationResult("device-token");
    }
  };

  const provider = new MsalGraphTokenProvider(client, (message) => {
    callbackMessage = message;
  });

  assert.equal(await provider.getAccessToken(), "device-token");
  assert.equal(deviceCodeCalls, 1);
  assert.equal(callbackMessage, "Use the Microsoft device code.");
});

test("falls back to device code when silent acquisition requires interaction", async () => {
  let deviceCodeCalls = 0;
  const client: MsalTokenClient = {
    getTokenCache: () => ({
      getAllAccounts: async () => [{ homeAccountId: "account-1" }] as never
    }),
    acquireTokenSilent: async () => {
      throw new InteractionRequiredAuthError("interaction_required", "interaction_required");
    },
    acquireTokenByDeviceCode: async () => {
      deviceCodeCalls += 1;
      return authenticationResult("device-token");
    }
  };

  const provider = new MsalGraphTokenProvider(client, () => undefined);

  assert.equal(await provider.getAccessToken(), "device-token");
  assert.equal(deviceCodeCalls, 1);
});

test("propagates unexpected silent authentication errors", async () => {
  const failure = new Error("network failure");
  const client: MsalTokenClient = {
    getTokenCache: () => ({
      getAllAccounts: async () => [{ homeAccountId: "account-1" }] as never
    }),
    acquireTokenSilent: async () => {
      throw failure;
    },
    acquireTokenByDeviceCode: async () => authenticationResult("device-token")
  };

  const provider = new MsalGraphTokenProvider(client, () => undefined);

  await assert.rejects(() => provider.getAccessToken(), (error: unknown) => error === failure);
});

test("throws a typed error when device code does not return an access token", async () => {
  const client: MsalTokenClient = {
    getTokenCache: () => ({ getAllAccounts: async () => [] }),
    acquireTokenSilent: async () => null,
    acquireTokenByDeviceCode: async () => null
  };

  const provider = new MsalGraphTokenProvider(client, () => undefined);

  await assert.rejects(
    () => provider.getAccessToken(),
    (error: unknown) => {
      assert.ok(error instanceof GraphAuthenticationError);
      assert.match(error.message, /access token/i);
      return true;
    }
  );
});

test("uses GRAPH_ACCESS_TOKEN as an explicit override", async () => {
  const accessToken = await resolveGraphAccessToken(
    { GRAPH_ACCESS_TOKEN: " manual-token " },
    async () => ({ getAccessToken: async () => "unexpected-token" })
  );

  assert.equal(accessToken, "manual-token");
});

test("requires GRAPH_CLIENT_ID when no token override exists", async () => {
  await assert.rejects(
    () => resolveGraphAccessToken({}, async () => ({ getAccessToken: async () => "unexpected" })),
    /GRAPH_CLIENT_ID is required/i
  );
});

test("uses trimmed client and tenant configuration when authenticating", async () => {
  let receivedClientId = "";
  let receivedTenantId = "";

  const accessToken = await resolveGraphAccessToken(
    {
      GRAPH_CLIENT_ID: " client-123 ",
      GRAPH_TENANT_ID: " tenant-456 "
    },
    async (options) => {
      receivedClientId = options.clientId;
      receivedTenantId = options.tenantId ?? "";
      return { getAccessToken: async () => "authenticated-token" };
    }
  );

  assert.equal(accessToken, "authenticated-token");
  assert.equal(receivedClientId, "client-123");
  assert.equal(receivedTenantId, "tenant-456");
});
