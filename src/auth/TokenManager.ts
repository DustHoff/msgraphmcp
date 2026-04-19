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
  'Team.ReadWrite.All',
  'Channel.ReadWrite.All',
  'ChannelMessage.Send',
  'Directory.ReadWrite.All',
  'DeviceManagementApps.ReadWrite.All',
  'DeviceManagementConfiguration.ReadWrite.All',
  'DeviceManagementManagedDevices.ReadWrite.All',
  'DeviceManagementServiceConfig.ReadWrite.All',
  'offline_access',
];

export const SCOPES: string[] = process.env.GRAPH_SCOPES
  ? process.env.GRAPH_SCOPES.split(' ').filter(Boolean)
  : DEFAULT_SCOPES;

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
        fs.writeFileSync(cachePath, ctx.tokenCache.serialize());
      }
    },
  };
}

export class TokenManager {
  private app: PublicClientApplication | ConfidentialClientApplication;
  private isConfidential: boolean;

  constructor() {
    const clientId = process.env.AZURE_CLIENT_ID;
    const tenantId = process.env.AZURE_TENANT_ID || 'common';
    const clientSecret = process.env.AZURE_CLIENT_SECRET;

    if (!clientId) throw new Error('AZURE_CLIENT_ID environment variable is required');

    const authority = `https://login.microsoftonline.com/${tenantId}`;
    this.isConfidential = Boolean(clientSecret);

    const msalConfig: Configuration = {
      auth: { clientId, authority, clientSecret },
      cache: { cachePlugin: createCachePlugin() },
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
    const account = await this.getAccount();

    if (account) {
      try {
        const silentRequest: SilentFlowRequest = { account, scopes: SCOPES };
        const result = await this.app.acquireTokenSilent(silentRequest);
        if (result?.accessToken) return result.accessToken;
      } catch {
        // fall through to interactive
      }
    }

    // Device code flow (works for both public and confidential clients)
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
}
