import {
  PublicClientApplication,
  ConfidentialClientApplication,
  Configuration,
  ICachePlugin,
  TokenCacheContext,
  AccountInfo,
  SilentFlowRequest,
} from '@azure/msal-node';
import * as fs from 'fs';
import * as path from 'path';

const TOKEN_CACHE_PATH = process.env.TOKEN_CACHE_PATH || '/data/tokens.json';

const DEFAULT_SCOPES = [
  'User.Read',
  'User.ReadWrite.All',
  'Group.ReadWrite.All',
  'GroupMember.ReadWrite.All',
  'Mail.ReadWrite',
  'Mail.Send',
  'Calendars.ReadWrite',
  'Files.ReadWrite.All',
  'Sites.ReadWrite.All',
  'Tasks.ReadWrite',
  'Contacts.ReadWrite',
  // Teams: Team.ReadWrite.All does not exist as a delegated permission — use granular scopes
  'Team.ReadBasic.All',
  'Team.Create',
  'TeamMember.ReadWrite.All',
  'Channel.ReadBasic.All',
  'Channel.Create',
  'Channel.Delete.All',
  'ChannelMessage.Send',
  'Directory.ReadWrite.All',
  'DeviceManagementApps.ReadWrite.All',
  'DeviceManagementConfiguration.ReadWrite.All',
  'DeviceManagementManagedDevices.ReadWrite.All',
  'DeviceManagementServiceConfig.ReadWrite.All',
  'offline_access',
];

// Scopes for delegated (device code / auth code) flow — overridable via GRAPH_SCOPES env var.
export const SCOPES: string[] = process.env.GRAPH_SCOPES
  ? process.env.GRAPH_SCOPES.split(' ').filter(Boolean)
  : DEFAULT_SCOPES;

// The .default scope instructs Entra ID to grant all application permissions
// that have been pre-consented for the app registration.
// Used exclusively with the client credentials (app-only) flow.
const APP_ONLY_SCOPES = ['https://graph.microsoft.com/.default'];

function createCachePlugin(): ICachePlugin {
  const cachePath = TOKEN_CACHE_PATH;

  return {
    async beforeCacheAccess(ctx: TokenCacheContext) {
      if (fs.existsSync(cachePath)) {
        const data = fs.readFileSync(cachePath, 'utf8');
        ctx.tokenCache.deserialize(data);
      }
    },
    async afterCacheAccess(ctx: TokenCacheContext) {
      if (ctx.cacheHasChanged) {
        const dir = path.dirname(cachePath);
        if (!fs.existsSync(dir)) {
          fs.mkdirSync(dir, { recursive: true });
        }
        // mode 0o600: only the owning process user may read/write the token cache
        fs.writeFileSync(cachePath, ctx.tokenCache.serialize(), { mode: 0o600 });
      }
    },
  };
}

// ── Auth mode detection ────────────────────────────────────────────────────────
//
// The server supports four authentication modes selected by environment variables:
//
// 1. Authorization Code (confidential client, delegated — recommended for HTTP mode)
//    Set: AZURE_CLIENT_SECRET + AZURE_REDIRECT_URI
//    Flow: OAuth 2.0 Authorization Code + PKCE → user authenticates in browser
//    Visit /auth/login to start the flow. Redirect URI must be registered in Entra.
//    CA compliance: evaluated against the USER'S browser device.
//
// 2. Client Secret (confidential client, app-only)
//    Set: AZURE_CLIENT_SECRET (without AZURE_REDIRECT_URI)
//    Flow: client_credentials → acquireTokenByClientCredential()
//    Requires Application permissions (not delegated) + admin consent.
//
// 3. Client Certificate (confidential client, app-only, preferred for K8s)
//    Set: AZURE_CLIENT_CERTIFICATE_PATH + AZURE_CLIENT_CERTIFICATE_THUMBPRINT
//    Flow: client_credentials with cert assertion
//    Requires Application permissions + admin consent.
//
// 4. Device Code (public client, delegated — local / stdio use only)
//    Set: none of the above
//    Flow: device_code → user authenticates in browser → acquireTokenByDeviceCode()
//    Not suitable for containers if CA compliance policies are active.

export type AuthMode = 'authorization-code' | 'client-secret' | 'client-certificate' | 'device-code';

export class AuthRequiredError extends Error {
  constructor(message = 'Not authenticated — visit /auth/login to sign in with Microsoft') {
    super(message);
    this.name = 'AuthRequiredError';
  }
}

export interface TokenManagerOptions {
  /** When false, tokens are kept in-memory only (no file I/O). Use for per-session isolation. */
  persistCache?: boolean;
}

export class TokenManager {
  private app: PublicClientApplication | ConfidentialClientApplication;
  private isConfidential: boolean;
  readonly authMode: AuthMode;

  constructor(options: TokenManagerOptions = {}) {
    const { persistCache = true } = options;
    const clientId = process.env.AZURE_CLIENT_ID;
    const tenantId = process.env.AZURE_TENANT_ID || 'common';
    const clientSecret = process.env.AZURE_CLIENT_SECRET;
    const certPath = process.env.AZURE_CLIENT_CERTIFICATE_PATH;
    const certThumbprint = process.env.AZURE_CLIENT_CERTIFICATE_THUMBPRINT;
    const redirectUri = process.env.AZURE_REDIRECT_URI;

    if (!clientId) throw new Error('AZURE_CLIENT_ID environment variable is required');

    const authority = `https://login.microsoftonline.com/${tenantId}`;

    // Build certificate configuration if both path and thumbprint are provided
    let clientCertificate:
      | { thumbprintSha256?: string; thumbprint?: string; privateKey: string }
      | undefined;

    if (certPath && certThumbprint) {
      if (!fs.existsSync(certPath)) {
        throw new Error(`AZURE_CLIENT_CERTIFICATE_PATH does not exist: ${certPath}`);
      }
      const privateKey = fs.readFileSync(certPath, 'utf8');
      // SHA-256 thumbprints are 64 hex chars; SHA-1 are 40 hex chars (legacy ADFS only)
      clientCertificate =
        certThumbprint.length === 64
          ? { thumbprintSha256: certThumbprint, privateKey }
          : { thumbprint: certThumbprint, privateKey };
    } else if (certPath || certThumbprint) {
      throw new Error(
        'Both AZURE_CLIENT_CERTIFICATE_PATH and AZURE_CLIENT_CERTIFICATE_THUMBPRINT ' +
          'must be set together.'
      );
    }

    // Authorization code flow takes priority over client credentials when AZURE_REDIRECT_URI is set.
    this.isConfidential = Boolean(clientSecret || clientCertificate);
    this.authMode = redirectUri && clientSecret
      ? 'authorization-code'
      : clientCertificate
        ? 'client-certificate'
        : clientSecret
          ? 'client-secret'
          : 'device-code';

    if (this.authMode === 'authorization-code' && !clientSecret) {
      throw new Error('AZURE_CLIENT_SECRET is required for authorization-code flow');
    }

    const msalConfig: Configuration = {
      auth: {
        clientId,
        authority,
        ...(clientSecret ? { clientSecret } : {}),
        ...(clientCertificate ? { clientCertificate } : {}),
      },
      cache: persistCache ? { cachePlugin: createCachePlugin() } : {},
      system: { loggerOptions: { loggerCallback: () => {}, piiLoggingEnabled: false } },
    };

    this.app = this.isConfidential
      ? new ConfidentialClientApplication(msalConfig)
      : new PublicClientApplication(msalConfig as Configuration);
  }

  private async getAccount(): Promise<AccountInfo | null> {
    const cache = this.app.getTokenCache();
    const accounts = await cache.getAllAccounts();
    return accounts.length > 0 ? accounts[0] : null;
  }

  async getAccessToken(): Promise<string> {
    // ── Authorization Code flow ─────────────────────────────────────────────
    // Triggered when AZURE_REDIRECT_URI + AZURE_CLIENT_SECRET are both set.
    // Tokens are acquired interactively via the browser at /auth/login and cached.
    // After initial login, subsequent calls use silent refresh.
    if (this.authMode === 'authorization-code') {
      const account = await this.getAccount();
      if (account) {
        try {
          const cca = this.app as ConfidentialClientApplication;
          const result = await cca.acquireTokenSilent({ account, scopes: SCOPES });
          if (result?.accessToken) return result.accessToken;
        } catch {
          // refresh failed — fall through to AuthRequiredError
        }
      }
      throw new AuthRequiredError();
    }

    // ── App-only (client credentials) flow ─────────────────────────────────
    // Triggered when AZURE_CLIENT_SECRET or AZURE_CLIENT_CERTIFICATE_PATH is set
    // (without AZURE_REDIRECT_URI).
    //
    // Requires Application permissions + admin consent on the app registration.
    // Tool parameters like userId='me' will not resolve — use explicit UPNs/object IDs.
    if (this.authMode === 'client-secret' || this.authMode === 'client-certificate') {
      const cca = this.app as ConfidentialClientApplication;
      const result = await cca.acquireTokenByClientCredential({
        scopes: APP_ONLY_SCOPES,
      });
      if (result?.accessToken) return result.accessToken;
      throw new Error('Client credentials flow returned no access token');
    }

    // ── Delegated (device code) flow ────────────────────────────────────────
    // Used when no client secret or certificate is configured.
    // Suitable for local / stdio usage where a human can interact with a browser.
    //
    // CA compliance warning: token refresh requests are issued from this process.
    // If Entra ID CA policies require device compliance on the refreshing device,
    // and this process runs on a non-enrolled container, silent refresh WILL fail.
    const account = await this.getAccount();

    if (account) {
      try {
        const silentRequest: SilentFlowRequest = { account, scopes: SCOPES };
        const result = await (this.app as PublicClientApplication).acquireTokenSilent(
          silentRequest
        );
        if (result?.accessToken) return result.accessToken;
      } catch {
        // fall through to interactive device code
      }
    }

    const pca = this.app as PublicClientApplication;
    const result = await pca.acquireTokenByDeviceCode({
      scopes: SCOPES,
      deviceCodeCallback: (response) => {
        process.stderr.write(`\n${'='.repeat(60)}\n`);
        process.stderr.write(response.message + '\n');
        process.stderr.write(`${'='.repeat(60)}\n\n`);
      },
    });

    if (!result?.accessToken) throw new Error('Authentication failed: no access token returned');
    return result.accessToken;
  }

  // ── Authorization Code helpers ──────────────────────────────────────────────

  /** Generates the Microsoft login URL for the OAuth authorization code flow. */
  async getAuthCodeUrl(redirectUri: string, state: string, codeChallenge: string): Promise<string> {
    if (this.authMode !== 'authorization-code') {
      throw new Error('getAuthCodeUrl is only available in authorization-code mode');
    }
    const cca = this.app as ConfidentialClientApplication;
    return cca.getAuthCodeUrl({
      scopes: SCOPES,
      redirectUri,
      state,
      codeChallenge,
      codeChallengeMethod: 'S256',
    });
  }

  /** Exchanges an authorization code (from /auth/callback) for tokens. */
  async acquireTokenByAuthCode(
    code: string,
    redirectUri: string,
    codeVerifier: string
  ): Promise<void> {
    if (this.authMode !== 'authorization-code') {
      throw new Error('acquireTokenByAuthCode is only available in authorization-code mode');
    }
    const cca = this.app as ConfidentialClientApplication;
    const result = await cca.acquireTokenByCode({ code, scopes: SCOPES, redirectUri, codeVerifier });
    if (!result?.accessToken) throw new Error('Authorization code exchange returned no access token');
  }

  /** Returns true if a valid cached account exists (delegated flows only). */
  async isAuthenticated(): Promise<boolean> {
    if (this.authMode === 'client-secret' || this.authMode === 'client-certificate') return true;
    const account = await this.getAccount();
    return account !== null;
  }

  /** Returns the UPN and display name of the authenticated account, or null if not yet authenticated. */
  async getAccountInfo(): Promise<{ upn: string; name: string } | null> {
    if (this.authMode === 'client-secret' || this.authMode === 'client-certificate') {
      return { upn: 'app-only', name: 'app-only' };
    }
    const account = await this.getAccount();
    if (!account) return null;
    return { upn: account.username, name: account.name ?? account.username };
  }
}
