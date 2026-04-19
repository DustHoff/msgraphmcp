import path from 'path';
import fs from 'fs';
import os from 'os';

// ── MSAL mock ────────────────────────────────────────────────────────────────
const mockGetAllAccounts = jest.fn();
const mockAcquireTokenSilent = jest.fn();
const mockAcquireTokenByDeviceCode = jest.fn();
const mockSerialize = jest.fn().mockReturnValue('{}');
const mockDeserialize = jest.fn();

jest.mock('@azure/msal-node', () => ({
  PublicClientApplication: jest.fn().mockImplementation(() => ({
    getTokenCache: () => ({
      getAllAccounts: mockGetAllAccounts,
      serialize: mockSerialize,
      deserialize: mockDeserialize,
    }),
    acquireTokenSilent: mockAcquireTokenSilent,
    acquireTokenByDeviceCode: mockAcquireTokenByDeviceCode,
  })),
  ConfidentialClientApplication: jest.fn().mockImplementation(() => ({
    getTokenCache: () => ({
      getAllAccounts: mockGetAllAccounts,
      serialize: mockSerialize,
      deserialize: mockDeserialize,
    }),
    acquireTokenSilent: mockAcquireTokenSilent,
  })),
}));

// ── helpers ──────────────────────────────────────────────────────────────────
function makeEnv(overrides: Record<string, string> = {}) {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'mcp-test-'));
  process.env.AZURE_CLIENT_ID = overrides.AZURE_CLIENT_ID ?? 'test-client-id';
  process.env.AZURE_TENANT_ID = overrides.AZURE_TENANT_ID ?? 'test-tenant-id';
  process.env.AZURE_CLIENT_SECRET = overrides.AZURE_CLIENT_SECRET ?? '';
  process.env.TOKEN_CACHE_PATH = path.join(tmpDir, 'tokens.json');
  return tmpDir;
}

function cleanEnv() {
  delete process.env.AZURE_CLIENT_ID;
  delete process.env.AZURE_TENANT_ID;
  delete process.env.AZURE_CLIENT_SECRET;
  delete process.env.TOKEN_CACHE_PATH;
}

// ── tests ────────────────────────────────────────────────────────────────────
describe('TokenManager', () => {
  let tmpDir: string;

  beforeEach(() => {
    jest.resetModules();
    jest.clearAllMocks();
    tmpDir = makeEnv();
  });

  afterEach(() => {
    cleanEnv();
    fs.rmSync(tmpDir, { recursive: true, force: true });
  });

  it('throws when AZURE_CLIENT_ID is missing', async () => {
    delete process.env.AZURE_CLIENT_ID;
    const { TokenManager } = await import('../../src/auth/TokenManager');
    expect(() => new TokenManager()).toThrow('AZURE_CLIENT_ID');
  });

  it('uses device code flow when no cached account exists', async () => {
    mockGetAllAccounts.mockResolvedValue([]);
    mockAcquireTokenByDeviceCode.mockResolvedValue({ accessToken: 'device-code-token' });

    const { TokenManager } = await import('../../src/auth/TokenManager');
    const manager = new TokenManager();
    const token = await manager.getAccessToken();

    expect(token).toBe('device-code-token');
    expect(mockAcquireTokenByDeviceCode).toHaveBeenCalledTimes(1);
  });

  it('uses silent flow when a cached account exists', async () => {
    const fakeAccount = { homeAccountId: 'id', environment: 'login.windows.net', tenantId: 'tid', username: 'user' };
    mockGetAllAccounts.mockResolvedValue([fakeAccount]);
    mockAcquireTokenSilent.mockResolvedValue({ accessToken: 'silent-token' });

    const { TokenManager } = await import('../../src/auth/TokenManager');
    const manager = new TokenManager();
    const token = await manager.getAccessToken();

    expect(token).toBe('silent-token');
    expect(mockAcquireTokenSilent).toHaveBeenCalledTimes(1);
    expect(mockAcquireTokenByDeviceCode).not.toHaveBeenCalled();
  });

  it('falls back to device code when silent flow fails', async () => {
    const fakeAccount = { homeAccountId: 'id', environment: 'login.windows.net', tenantId: 'tid', username: 'user' };
    mockGetAllAccounts.mockResolvedValue([fakeAccount]);
    mockAcquireTokenSilent.mockRejectedValue(new Error('interaction_required'));
    mockAcquireTokenByDeviceCode.mockResolvedValue({ accessToken: 'fallback-token' });

    const { TokenManager } = await import('../../src/auth/TokenManager');
    const manager = new TokenManager();
    const token = await manager.getAccessToken();

    expect(token).toBe('fallback-token');
    expect(mockAcquireTokenByDeviceCode).toHaveBeenCalledTimes(1);
  });

  it('throws when authentication returns no token', async () => {
    mockGetAllAccounts.mockResolvedValue([]);
    mockAcquireTokenByDeviceCode.mockResolvedValue(null);

    const { TokenManager } = await import('../../src/auth/TokenManager');
    const manager = new TokenManager();
    await expect(manager.getAccessToken()).rejects.toThrow('Authentication failed');
  });
});
