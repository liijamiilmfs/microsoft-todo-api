import { mkdir } from "node:fs/promises";
import { homedir } from "node:os";
import { dirname, join } from "node:path";

import {
  InteractionRequiredAuthError,
  PublicClientApplication,
  type AccountInfo,
  type AuthenticationResult,
  type DeviceCodeRequest,
  type SilentFlowRequest
} from "@azure/msal-node";
import type { PersistenceCachePlugin } from "@azure/msal-node-extensions";

export const GRAPH_TASKS_SCOPE = "Tasks.ReadWrite";

export interface GraphAccessTokenProvider {
  getAccessToken(): Promise<string>;
}

export interface MsalTokenCache {
  getAllAccounts(): Promise<AccountInfo[]>;
}

export interface MsalTokenClient {
  getTokenCache(): MsalTokenCache;
  acquireTokenSilent(request: SilentFlowRequest): Promise<AuthenticationResult | null>;
  acquireTokenByDeviceCode(request: DeviceCodeRequest): Promise<AuthenticationResult | null>;
}

export interface GraphAuthOptions {
  clientId: string;
  tenantId?: string;
  cachePath?: string;
  onDeviceCode?: (message: string) => void;
}

export class GraphAuthenticationError extends Error {
  constructor(message: string, options?: ErrorOptions) {
    super(message, options);
    this.name = "GraphAuthenticationError";
  }
}

export class MsalGraphTokenProvider implements GraphAccessTokenProvider {
  private readonly client: MsalTokenClient;
  private readonly onDeviceCode: (message: string) => void;

  constructor(
    client: MsalTokenClient,
    onDeviceCode: (message: string) => void
  ) {
    this.client = client;
    this.onDeviceCode = onDeviceCode;
  }

  async getAccessToken(): Promise<string> {
    const accounts = await this.client.getTokenCache().getAllAccounts();
    if (accounts.length > 1) {
      throw new GraphAuthenticationError(
        "Multiple Microsoft accounts are cached. Clear this application's protected token cache, then sign in with the intended account."
      );
    }
    const account = accounts[0];

    if (account) {
      try {
        const result = await this.client.acquireTokenSilent({
          account,
          scopes: [GRAPH_TASKS_SCOPE]
        });

        if (result?.accessToken) {
          return result.accessToken;
        }
      } catch (error: unknown) {
        if (!(error instanceof InteractionRequiredAuthError)) {
          throw error;
        }
      }
    }

    const result = await this.client.acquireTokenByDeviceCode({
      scopes: [GRAPH_TASKS_SCOPE],
      deviceCodeCallback: (response) => {
        this.onDeviceCode(response.message);
      }
    });

    if (!result?.accessToken) {
      throw new GraphAuthenticationError("Microsoft Graph did not return an access token.");
    }

    return result.accessToken;
  }
}

export async function createGraphAccessTokenProvider(
  options: GraphAuthOptions
): Promise<GraphAccessTokenProvider> {
  if (!options.clientId.trim()) {
    throw new GraphAuthenticationError("GRAPH_CLIENT_ID is required.");
  }

  const cachePlugin = await createProtectedCache(options.cachePath ?? defaultCachePath());
  const client = new PublicClientApplication({
    auth: {
      clientId: options.clientId,
      authority: `https://login.microsoftonline.com/${options.tenantId ?? "common"}`
    },
    cache: {
      cachePlugin
    }
  });

  return new MsalGraphTokenProvider(
    client,
    options.onDeviceCode ?? ((message) => console.log(message))
  );
}

async function createProtectedCache(cachePath: string): Promise<PersistenceCachePlugin> {
  try {
    const { DataProtectionScope, PersistenceCachePlugin, PersistenceCreator } =
      await import("@azure/msal-node-extensions");
    await mkdir(dirname(cachePath), { recursive: true, mode: 0o700 });
    const persistence = await PersistenceCreator.createPersistence({
      cachePath,
      dataProtectionScope: DataProtectionScope.CurrentUser,
      serviceName: "microsoft-todo-api",
      accountName: "default",
      usePlaintextFileOnLinux: false
    });
    return new PersistenceCachePlugin(persistence);
  } catch (error: unknown) {
    throw new GraphAuthenticationError(
      "Unable to initialize the protected Microsoft Graph token cache. On Linux, install a Secret Service/LibSecret provider; plaintext cache fallback is disabled.",
      { cause: error }
    );
  }
}

export type GraphEnvironment = Pick<
  NodeJS.ProcessEnv,
  "GRAPH_ACCESS_TOKEN" | "GRAPH_CLIENT_ID" | "GRAPH_TENANT_ID"
>;

export type GraphAccessTokenProviderFactory = (
  options: GraphAuthOptions
) => Promise<GraphAccessTokenProvider>;

export async function resolveGraphAccessToken(
  env: GraphEnvironment,
  createProvider: GraphAccessTokenProviderFactory = createGraphAccessTokenProvider
): Promise<string> {
  const override = env.GRAPH_ACCESS_TOKEN?.trim();

  if (override) {
    return override;
  }

  const clientId = env.GRAPH_CLIENT_ID?.trim();

  if (!clientId) {
    throw new GraphAuthenticationError(
      "GRAPH_CLIENT_ID is required when GRAPH_ACCESS_TOKEN is not set."
    );
  }

  const provider = await createProvider({
    clientId,
    tenantId: env.GRAPH_TENANT_ID?.trim() || "common"
  });

  return provider.getAccessToken();
}

function defaultCachePath(): string {
  const cacheRoot =
    process.env.LOCALAPPDATA ?? process.env.XDG_CACHE_HOME ?? join(homedir(), ".cache");

  return join(cacheRoot, "microsoft-todo-api", "msal-cache.json");
}
