import path from 'path';
import fs from 'fs';
import os from 'os';

// ── MSAL mock ────────────────────────────────────────────────────────────────
const mockGetAllAccounts = jest.fn();
const mockAcquireTokenSilent = jest.fn();
const mockAcquireTokenByDeviceCode = jest.fn();
const mockAcquireTokenByClientCredential = jest.fn();
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
    acquireTokenByClientCredential: mockAcquireTokenByClientCredential,
  })),
}));

// ── helpers ──────────────────────────────────────────────────────────────────
function makeEnv(overrides: Record<string, string | undefined> = {}) {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'mcp-test-'));
  process.env.AZURE_CLIENT_ID = overrides.AZURE_CLIENT_ID ?? 'test-client-id';
  process.env.AZURE_TENANT_ID = overrides.AZURE_TENANT_ID ?? 'test-tenant-id';
  process.env.TOKEN_CACHE_PATH = path.join(tmpDir, 'tokens.json');

  delete process.env.AZURE_CLIENT_SECRET;
  delete process.env.AZURE_CLIENT_CERTIFICATE_PATH;
  delete process.env.AZURE_CLIENT_CERTIFICATE_THUMBPRINT;

  if (overrides.AZURE_CLIENT_SECRET !== undefined) {
    process.env.AZURE_CLIENT_SECRET = overrides.AZURE_CLIENT_SECRET;
  }
  if (overrides.AZURE_CLIENT_CERTIFICATE_PATH !== undefined) {
    process.env.AZURE_CLIENT_CERTIFICATE_PATH = overrides.AZURE_CLIENT_CERTIFICATE_PATH;
  }
  if (overrides.AZURE_CLIENT_CERTIFICATE_THUMBPRINT !== undefined) {
    process.env.AZURE_CLIENT_CERTIFICATE_THUMBPRINT = overrides.AZURE_CLIENT_CERTIFICATE_THUMBPRINT;
  }

  return tmpDir;
}

function cleanEnv() {
  delete process.env.AZURE_CLIENT_ID;
  delete process.env.AZURE_TENANT_ID;
  delete process.env.AZURE_CLIENT_SECRET;
  delete process.env.AZURE_CLIENT_CERTIFICATE_PATH;
  delete process.env.AZURE_CLIENT_CERTIFICATE_THUMBPRINT;
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

  // ── device-code flow ───────────────────────────────────────────────────────

  it('throws when AZURE_CLIENT_ID is missing', async () => {
    delete process.env.AZURE_CLIENT_ID;
    const { TokenManager } = await import('../../src/auth/TokenManager');
    expect(() => new TokenManager()).toThrow('AZURE_CLIENT_ID');
  });

  it('reports authMode as device-code when no secret or cert is set', async () => {
    const { TokenManager } = await import('../../src/auth/TokenManager');
    const manager = new TokenManager();
    expect(manager.authMode).toBe('device-code');
  });

  it('uses device code flow when no cached account exists', async () => {
    mockGetAllAccounts.mockResolvedValue([]);
    mockAcquireTokenByDeviceCode.mockResolvedValue({ accessToken: 'device-code-token' });

    const { TokenManager } = await import('../../src/auth/TokenManager');
    const manager = new TokenManager();
    const token = await manager.getAccessToken();

    expect(token).toBe('device-code-token');
    expect(mockAcquireTokenByDeviceCode).toHaveBeenCalledTimes(1);
    expect(mockAcquireTokenByClientCredential).not.toHaveBeenCalled();
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

  it('throws when device code authentication returns no token', async () => {
    mockGetAllAccounts.mockResolvedValue([]);
    mockAcquireTokenByDeviceCode.mockResolvedValue(null);

    const { TokenManager } = await import('../../src/auth/TokenManager');
    const manager = new TokenManager();
    await expect(manager.getAccessToken()).rejects.toThrow('Authentication failed');
  });

  // ── client-secret flow ─────────────────────────────────────────────────────

  it('reports authMode as client-secret when AZURE_CLIENT_SECRET is set', async () => {
    makeEnv({ AZURE_CLIENT_SECRET: 'my-secret' });
    const { TokenManager } = await import('../../src/auth/TokenManager');
    const manager = new TokenManager();
    expect(manager.authMode).toBe('client-secret');
  });

  it('uses client credentials flow when client secret is configured', async () => {
    makeEnv({ AZURE_CLIENT_SECRET: 'my-secret' });
    mockAcquireTokenByClientCredential.mockResolvedValue({ accessToken: 'app-only-token' });

    const { TokenManager } = await import('../../src/auth/TokenManager');
    const manager = new TokenManager();
    const token = await manager.getAccessToken();

    expect(token).toBe('app-only-token');
    expect(mockAcquireTokenByClientCredential).toHaveBeenCalledTimes(1);
    expect(mockAcquireTokenByClientCredential).toHaveBeenCalledWith({
      scopes: ['https://graph.microsoft.com/.default'],
    });
    expect(mockAcquireTokenByDeviceCode).not.toHaveBeenCalled();
    expect(mockAcquireTokenSilent).not.toHaveBeenCalled();
  });

  it('throws when client credentials flow returns no token', async () => {
    makeEnv({ AZURE_CLIENT_SECRET: 'my-secret' });
    mockAcquireTokenByClientCredential.mockResolvedValue(null);

    const { TokenManager } = await import('../../src/auth/TokenManager');
    const manager = new TokenManager();
    await expect(manager.getAccessToken()).rejects.toThrow('Client credentials flow returned no access token');
  });

  // ── client-certificate flow ────────────────────────────────────────────────

  it('reports authMode as client-certificate when cert env vars are set', async () => {
    const certFile = path.join(tmpDir, 'cert.pem');
    fs.writeFileSync(certFile, '-----BEGIN RSA PRIVATE KEY-----\nfake\n-----END RSA PRIVATE KEY-----');
    const thumbprint = 'a'.repeat(64); // 64-char = SHA-256
    makeEnv({ AZURE_CLIENT_CERTIFICATE_PATH: certFile, AZURE_CLIENT_CERTIFICATE_THUMBPRINT: thumbprint });

    const { TokenManager } = await import('../../src/auth/TokenManager');
    const manager = new TokenManager();
    expect(manager.authMode).toBe('client-certificate');
  });

  it('uses client credentials flow when certificate is configured', async () => {
    const certFile = path.join(tmpDir, 'cert.pem');
    fs.writeFileSync(certFile, '-----BEGIN RSA PRIVATE KEY-----\nfake\n-----END RSA PRIVATE KEY-----');
    const thumbprint = 'a'.repeat(64);
    makeEnv({ AZURE_CLIENT_CERTIFICATE_PATH: certFile, AZURE_CLIENT_CERTIFICATE_THUMBPRINT: thumbprint });
    mockAcquireTokenByClientCredential.mockResolvedValue({ accessToken: 'cert-token' });

    const { TokenManager } = await import('../../src/auth/TokenManager');
    const manager = new TokenManager();
    const token = await manager.getAccessToken();

    expect(token).toBe('cert-token');
    expect(mockAcquireTokenByClientCredential).toHaveBeenCalledTimes(1);
    expect(mockAcquireTokenByDeviceCode).not.toHaveBeenCalled();
  });

  it('throws when only AZURE_CLIENT_CERTIFICATE_PATH is set (missing thumbprint)', async () => {
    const certFile = path.join(tmpDir, 'cert.pem');
    fs.writeFileSync(certFile, 'fake-pem');
    makeEnv({ AZURE_CLIENT_CERTIFICATE_PATH: certFile });

    const { TokenManager } = await import('../../src/auth/TokenManager');
    expect(() => new TokenManager()).toThrow(
      'Both AZURE_CLIENT_CERTIFICATE_PATH and AZURE_CLIENT_CERTIFICATE_THUMBPRINT'
    );
  });

  it('throws when only AZURE_CLIENT_CERTIFICATE_THUMBPRINT is set (missing path)', async () => {
    makeEnv({ AZURE_CLIENT_CERTIFICATE_THUMBPRINT: 'a'.repeat(64) });

    const { TokenManager } = await import('../../src/auth/TokenManager');
    expect(() => new TokenManager()).toThrow(
      'Both AZURE_CLIENT_CERTIFICATE_PATH and AZURE_CLIENT_CERTIFICATE_THUMBPRINT'
    );
  });

  it('throws when AZURE_CLIENT_CERTIFICATE_PATH does not exist', async () => {
    makeEnv({
      AZURE_CLIENT_CERTIFICATE_PATH: path.join(tmpDir, 'nonexistent.pem'),
      AZURE_CLIENT_CERTIFICATE_THUMBPRINT: 'a'.repeat(64),
    });

    const { TokenManager } = await import('../../src/auth/TokenManager');
    expect(() => new TokenManager()).toThrow('does not exist');
  });
});
