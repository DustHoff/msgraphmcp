import {
  detectHttpAuthMode,
  resolveHttpSecurityPolicy,
  isAuthorizedMcpRequest,
  HttpSecurityError,
} from '../../src/http/security';

function env(overrides: Record<string, string | undefined>): NodeJS.ProcessEnv {
  return overrides as NodeJS.ProcessEnv;
}

const TOKEN = 'a'.repeat(48); // high-entropy-length stand-in

describe('detectHttpAuthMode', () => {
  it('returns authorization-code when redirect URI + secret are set', () => {
    expect(
      detectHttpAuthMode(env({ AZURE_REDIRECT_URI: 'https://h/auth/callback', AZURE_CLIENT_SECRET: 's' }))
    ).toBe('authorization-code');
  });

  it('returns client-certificate when a cert path is set (no redirect)', () => {
    expect(detectHttpAuthMode(env({ AZURE_CLIENT_CERTIFICATE_PATH: '/mnt/cert/tls.key' }))).toBe(
      'client-certificate'
    );
  });

  it('returns client-secret when only a secret is set', () => {
    expect(detectHttpAuthMode(env({ AZURE_CLIENT_SECRET: 's' }))).toBe('client-secret');
  });

  it('returns device-code when nothing confidential is set', () => {
    expect(detectHttpAuthMode(env({ AZURE_CLIENT_ID: 'id' }))).toBe('device-code');
  });
});

describe('resolveHttpSecurityPolicy', () => {
  it('accepts authorization-code with an inbound token', () => {
    const policy = resolveHttpSecurityPolicy(
      env({ AZURE_REDIRECT_URI: 'https://h/auth/callback', AZURE_CLIENT_SECRET: 's', MCP_AUTH_TOKEN: TOKEN })
    );
    expect(policy.authMode).toBe('authorization-code');
    expect(policy.gatingEnabled).toBe(true);
    expect(policy.mcpAuthToken).toBe(TOKEN);
    expect(policy.warnings).toHaveLength(0);
  });

  it('fails closed on app-only (client-secret) over HTTP', () => {
    expect(() =>
      resolveHttpSecurityPolicy(env({ AZURE_CLIENT_SECRET: 's', MCP_AUTH_TOKEN: TOKEN }))
    ).toThrow(HttpSecurityError);
  });

  it('fails closed on client-certificate over HTTP', () => {
    expect(() =>
      resolveHttpSecurityPolicy(env({ AZURE_CLIENT_CERTIFICATE_PATH: '/c.key', MCP_AUTH_TOKEN: TOKEN }))
    ).toThrow(/AUTH-1/);
  });

  it('fails closed on device-code over HTTP', () => {
    expect(() => resolveHttpSecurityPolicy(env({ AZURE_CLIENT_ID: 'id', MCP_AUTH_TOKEN: TOKEN }))).toThrow(
      /authorization-code/
    );
  });

  it('permits app-only with explicit MCP_ALLOW_INSECURE_AUTH_MODE override (with warning)', () => {
    const policy = resolveHttpSecurityPolicy(
      env({ AZURE_CLIENT_SECRET: 's', MCP_AUTH_TOKEN: TOKEN, MCP_ALLOW_INSECURE_AUTH_MODE: 'true' })
    );
    expect(policy.authMode).toBe('client-secret');
    expect(policy.warnings.join(' ')).toMatch(/trusted network boundary/);
  });

  it('fails closed in authorization-code mode when no inbound token is configured', () => {
    expect(() =>
      resolveHttpSecurityPolicy(env({ AZURE_REDIRECT_URI: 'https://h/auth/callback', AZURE_CLIENT_SECRET: 's' }))
    ).toThrow(/MCP_AUTH_TOKEN/);
  });

  it('permits no token only with explicit MCP_ALLOW_UNAUTHENTICATED override (with warning)', () => {
    const policy = resolveHttpSecurityPolicy(
      env({
        AZURE_REDIRECT_URI: 'https://h/auth/callback',
        AZURE_CLIENT_SECRET: 's',
        MCP_ALLOW_UNAUTHENTICATED: 'true',
      })
    );
    expect(policy.gatingEnabled).toBe(false);
    expect(policy.mcpAuthToken).toBeUndefined();
    expect(policy.warnings.join(' ')).toMatch(/UNAUTHENTICATED/);
  });

  it('warns when the inbound token is too short', () => {
    const policy = resolveHttpSecurityPolicy(
      env({ AZURE_REDIRECT_URI: 'https://h/auth/callback', AZURE_CLIENT_SECRET: 's', MCP_AUTH_TOKEN: 'short' })
    );
    expect(policy.warnings.join(' ')).toMatch(/high-entropy/);
  });
});

describe('isAuthorizedMcpRequest', () => {
  it('allows any request when gating is disabled (no expected token)', () => {
    expect(isAuthorizedMcpRequest(undefined, undefined)).toBe(true);
    expect(isAuthorizedMcpRequest('Bearer whatever', undefined)).toBe(true);
  });

  it('accepts a matching Bearer token', () => {
    expect(isAuthorizedMcpRequest(`Bearer ${TOKEN}`, TOKEN)).toBe(true);
  });

  it('accepts a matching bare token (no Bearer prefix)', () => {
    expect(isAuthorizedMcpRequest(TOKEN, TOKEN)).toBe(true);
  });

  it('is case-insensitive on the Bearer scheme', () => {
    expect(isAuthorizedMcpRequest(`bearer ${TOKEN}`, TOKEN)).toBe(true);
  });

  it('rejects a wrong token', () => {
    expect(isAuthorizedMcpRequest(`Bearer ${'b'.repeat(48)}`, TOKEN)).toBe(false);
  });

  it('rejects a missing header when gating is enabled', () => {
    expect(isAuthorizedMcpRequest(undefined, TOKEN)).toBe(false);
    expect(isAuthorizedMcpRequest('', TOKEN)).toBe(false);
  });

  it('rejects a token that is a prefix of the expected (no length leak / partial match)', () => {
    expect(isAuthorizedMcpRequest(`Bearer ${TOKEN.slice(0, 40)}`, TOKEN)).toBe(false);
  });
});
