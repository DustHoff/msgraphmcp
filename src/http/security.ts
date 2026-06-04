import { createHash, timingSafeEqual } from 'crypto';

// ── HTTP-mode security policy ────────────────────────────────────────────────
//
// When the server runs in HTTP mode (PORT set, Kubernetes / public exposure) the
// network allowlist can no longer be assumed to be the security boundary. This
// module fails closed on configurations that would let an external attacker reach
// Microsoft Graph, and provides the inbound auth gate for the /mcp endpoint.
//
// See SECURITY-AUDIT-2026-06-04.md findings AUTH-1, AUTH-4, CSRF-1.

export type HttpAuthMode =
  | 'authorization-code'
  | 'client-secret'
  | 'client-certificate'
  | 'device-code';

/**
 * Derives the auth mode from the environment, mirroring TokenManager's own
 * detection so the HTTP startup guard and the TokenManager always agree.
 */
export function detectHttpAuthMode(env: NodeJS.ProcessEnv): HttpAuthMode {
  const redirectUri = env.AZURE_REDIRECT_URI;
  const clientSecret = env.AZURE_CLIENT_SECRET;
  const certPath = env.AZURE_CLIENT_CERTIFICATE_PATH;

  if (redirectUri && clientSecret) return 'authorization-code';
  if (certPath) return 'client-certificate';
  if (clientSecret) return 'client-secret';
  return 'device-code';
}

export interface HttpSecurityPolicy {
  authMode: HttpAuthMode;
  /** The bearer token required on every /mcp request, or undefined when gating is disabled. */
  mcpAuthToken?: string;
  /** True when an inbound bearer token is enforced on /mcp. */
  gatingEnabled: boolean;
  /** Non-fatal configuration warnings to surface to the operator at startup. */
  warnings: string[];
}

/** Raised when the HTTP configuration is unsafe to start. The message is operator-facing. */
export class HttpSecurityError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'HttpSecurityError';
  }
}

/**
 * Validates the HTTP-mode security configuration and fails closed on insecure
 * setups. Returns the resolved policy, or throws HttpSecurityError with an
 * actionable message when the server must not start.
 *
 * Rules (overridable only by explicit opt-in env flags, for trusted-network use):
 *  1. Only 'authorization-code' (delegated, per-user) may run over HTTP.
 *     app-only modes hand ANY caller full app-identity Graph access (AUTH-1);
 *     device-code leaks the device code to logs and is interactive/stdio-only (AUTH-4).
 *     Override: MCP_ALLOW_INSECURE_AUTH_MODE=true
 *  2. An inbound auth gate (MCP_AUTH_TOKEN) must protect /mcp so that only
 *     authorized MCP clients can open a session / obtain a login URL (CSRF-1).
 *     Override: MCP_ALLOW_UNAUTHENTICATED=true
 */
export function resolveHttpSecurityPolicy(env: NodeJS.ProcessEnv): HttpSecurityPolicy {
  const authMode = detectHttpAuthMode(env);
  const warnings: string[] = [];

  // 1. Fail closed for auth modes that are unsafe over HTTP.
  const allowInsecureMode = env.MCP_ALLOW_INSECURE_AUTH_MODE === 'true';
  if (authMode !== 'authorization-code') {
    if (!allowInsecureMode) {
      throw new HttpSecurityError(
        `Refusing to start the HTTP server in '${authMode}' mode.\n` +
          `Only 'authorization-code' (delegated, per-user) is permitted for HTTP / public exposure:\n` +
          `  - app-only modes (client-secret / client-certificate) grant ANY caller full\n` +
          `    app-identity Graph access with no user interaction (audit finding AUTH-1).\n` +
          `  - device-code is interactive / stdio-only and leaks the device code to logs\n` +
          `    over HTTP (audit finding AUTH-4).\n` +
          `Set AZURE_REDIRECT_URI and AZURE_CLIENT_SECRET to enable authorization-code mode.\n` +
          `To override on a trusted, non-public network only, set MCP_ALLOW_INSECURE_AUTH_MODE=true.`
      );
    }
    warnings.push(
      `auth mode '${authMode}' permitted over HTTP via MCP_ALLOW_INSECURE_AUTH_MODE — ` +
        `the endpoint MUST sit behind a trusted network boundary; never expose it publicly`
    );
  }

  // 2. Inbound auth gate on /mcp.
  const mcpAuthToken = env.MCP_AUTH_TOKEN?.trim() || undefined;
  const allowUnauthenticated = env.MCP_ALLOW_UNAUTHENTICATED === 'true';
  if (!mcpAuthToken) {
    if (!allowUnauthenticated) {
      throw new HttpSecurityError(
        `Refusing to start the HTTP server without an inbound auth gate on /mcp.\n` +
          `Set MCP_AUTH_TOKEN to a high-entropy secret (e.g. \`openssl rand -hex 32\`) so that only\n` +
          `authorized MCP clients can reach /mcp. Configure Claude Code with:\n` +
          `  claude mcp add --transport http msgraphmcp https://<host>/mcp \\\n` +
          `      --header "Authorization: Bearer <token>"\n` +
          `To run without the gate on a trusted, non-public network only, set MCP_ALLOW_UNAUTHENTICATED=true.`
      );
    }
    warnings.push(
      'MCP_AUTH_TOKEN not set — /mcp is UNAUTHENTICATED (MCP_ALLOW_UNAUTHENTICATED=true); ' +
        'access control relies entirely on network controls'
    );
  } else if (mcpAuthToken.length < 32) {
    warnings.push('MCP_AUTH_TOKEN is shorter than 32 characters — use a high-entropy secret (>= 32 chars)');
  }

  return {
    authMode,
    mcpAuthToken,
    gatingEnabled: Boolean(mcpAuthToken),
    warnings,
  };
}

/**
 * Extracts the presented credential from an Authorization header. Accepts both
 * `Bearer <token>` and a bare token value (some MCP clients send the raw token).
 */
function extractPresentedToken(authHeader: string | undefined): string | undefined {
  if (!authHeader) return undefined;
  const trimmed = authHeader.trim();
  if (!trimmed) return undefined;
  const match = /^Bearer\s+(.+)$/i.exec(trimmed);
  return match ? match[1].trim() : trimmed;
}

/**
 * Constant-time check of an inbound /mcp request's Authorization header against
 * the expected bearer token. Returns true when gating is disabled (no expected
 * token). Both sides are SHA-256 hashed to a fixed length so timingSafeEqual
 * never throws on length mismatch and the comparison leaks no length information.
 */
export function isAuthorizedMcpRequest(
  authHeader: string | undefined,
  expectedToken: string | undefined
): boolean {
  if (!expectedToken) return true; // gating disabled
  const presented = extractPresentedToken(authHeader);
  if (!presented) return false;

  const presentedDigest = createHash('sha256').update(presented).digest();
  const expectedDigest = createHash('sha256').update(expectedToken).digest();
  return timingSafeEqual(presentedDigest, expectedDigest);
}
