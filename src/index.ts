import http, { IncomingMessage, ServerResponse } from 'http';
import { createHash, randomBytes, randomUUID } from 'crypto';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js';
import { TokenManager, AuthRequiredError } from './auth/TokenManager';
import { GraphClient } from './graph/GraphClient';
import { registerAllTools } from './tools/index';
import { logger } from './logger';

// ── Helpers ──────────────────────────────────────────────────────────────────

// Guard against request bodies that could cause OOM before JSON parsing
const MAX_BODY_BYTES = 4 * 1024 * 1024; // 4 MB — ample for any MCP JSON-RPC message

function parseBody(req: IncomingMessage): Promise<unknown> {
  return new Promise((resolve, reject) => {
    const chunks: Buffer[] = [];
    let bytesReceived = 0;

    req.on('data', (chunk: Buffer) => {
      bytesReceived += chunk.length;
      if (bytesReceived > MAX_BODY_BYTES) {
        req.destroy(new Error('Request body exceeds 4 MB limit'));
        return;
      }
      chunks.push(chunk);
    });
    req.on('end', () => {
      const raw = Buffer.concat(chunks).toString('utf8');
      try {
        resolve(raw ? JSON.parse(raw) : undefined);
      } catch {
        resolve(undefined);
      }
    });
    req.on('error', reject);
  });
}

function createMcpServer(graphClient: GraphClient): McpServer {
  const server = new McpServer({ name: 'msgraphmcp', version: '1.0.0' });
  registerAllTools(server, graphClient);
  return server;
}

// ── PKCE helpers ─────────────────────────────────────────────────────────────

function generateCodeVerifier(): string {
  return randomBytes(32).toString('base64url');
}

function generateCodeChallenge(verifier: string): string {
  return createHash('sha256').update(verifier).digest('base64url');
}

// ── stdio mode (Claude Code / local) ─────────────────────────────────────────

async function startStdio(graphClient: GraphClient, tokenManager: TokenManager): Promise<void> {
  const server = createMcpServer(graphClient);

  server.server.oninitialized = () => {
    const ci = server.server.getClientVersion();
    logger.info('mcp client connected', { client: ci?.name, clientVersion: ci?.version });
  };

  logger.info('authenticating with microsoft graph');
  await tokenManager.getAccessToken();
  logger.info('authentication successful — stdio mcp server ready');

  const transport = new StdioServerTransport();
  await server.connect(transport);
}

// ── HTTP / Streamable-HTTP mode (Kubernetes) ──────────────────────────────────

interface SessionData {
  transport: StreamableHTTPServerTransport;
  tokenManager: TokenManager;
  graphClient: GraphClient;
  lastActivityAt: number;
}

// Maximum number of concurrent MCP sessions. Prevents OOM via session flooding.
const MAX_SESSIONS = parseInt(process.env.MAX_SESSIONS ?? '50', 10);

// Idle sessions are closed after this many minutes of inactivity (default 60).
const SESSION_IDLE_TIMEOUT_MS =
  parseInt(process.env.SESSION_IDLE_TIMEOUT_MINUTES ?? '60', 10) * 60 * 1000;

async function startHttp(port: number): Promise<void> {
  // Map of active sessions: sessionId → { transport, tokenManager, graphClient }
  // Each session gets its own isolated token — no token bleed between users.
  const sessions = new Map<string, SessionData>();

  const isAuthCodeMode = Boolean(process.env.AZURE_REDIRECT_URI && process.env.AZURE_CLIENT_SECRET);
  const REDIRECT_URI = process.env.AZURE_REDIRECT_URI ?? '';
  const BASE_LOGIN_URL = REDIRECT_URI.replace('/auth/callback', '/auth/login');

  if (isAuthCodeMode) {
    logger.info('auth mode: authorization-code (delegated, per-session)', {
      redirectUri: REDIRECT_URI,
      loginUrl: BASE_LOGIN_URL,
    });
  } else if (process.env.AZURE_CLIENT_SECRET) {
    logger.warn(
      'auth mode: client-credentials (app-only) — AZURE_REDIRECT_URI is not set. ' +
      'Set AZURE_REDIRECT_URI to enable delegated (per-user) authentication.',
      { hint: 'Each session will authenticate as the application identity, not as a user.' }
    );
  } else {
    logger.info('auth mode: device-code (delegated, interactive)');
  }

  // Pending OAuth state: state-value → { codeVerifier, sessionId }
  // sessionId binds the OAuth callback to the exact MCP session that initiated the login.
  // Entries expire after 10 minutes to avoid unbounded growth.
  const pendingAuth = new Map<string, { codeVerifier: string; sessionId: string; expiresAt: number }>();

  // One-time login tokens: token → { sessionId, expiresAt }
  // Session IDs are never exposed in URLs — instead a short-lived, one-time-use token is
  // issued per authentication attempt. This prevents session token injection: even if an
  // attacker sees the login URL (from browser history, logs, etc.) they cannot reuse it
  // after the legitimate user has clicked it, and the session ID itself is never revealed.
  const loginTokens = new Map<string, { sessionId: string; expiresAt: number }>();
  const LOGIN_TOKEN_TTL_MS = 15 * 60 * 1000; // 15 minutes

  function generateLoginUrl(sid: string): string {
    const token = randomBytes(32).toString('hex');
    loginTokens.set(token, { sessionId: sid, expiresAt: Date.now() + LOGIN_TOKEN_TTL_MS });
    return `${BASE_LOGIN_URL}?token=${token}`;
  }

  // Close sessions that have been idle longer than SESSION_IDLE_TIMEOUT_MS.
  const idleCheckInterval = setInterval(() => {
    const cutoff = Date.now() - SESSION_IDLE_TIMEOUT_MS;
    for (const [id, session] of sessions) {
      if (session.lastActivityAt < cutoff) {
        logger.info('mcp session idle timeout — closing', { sessionId: id });
        session.transport.close();
        sessions.delete(id);
      }
    }
  }, 5 * 60 * 1000);
  idleCheckInterval.unref();

  function purgeStalePending() {
    const now = Date.now();
    for (const [key, val] of pendingAuth) {
      if (val.expiresAt < now) pendingAuth.delete(key);
    }
    for (const [tok, entry] of loginTokens) {
      if (entry.expiresAt < now) loginTokens.delete(tok);
    }
  }

  const httpServer = http.createServer(async (req: IncomingMessage, res: ServerResponse) => {
    const url = new URL(req.url ?? '/', `http://localhost:${port}`);

    // ── Health/readiness probe ───────────────────────────────────────────────
    // Intentionally returns no session IDs, UPNs, or per-session detail —
    // that data would let anyone reaching /health enumerate sessions and UPNs.
    if (url.pathname === '/health') {
      res.writeHead(200, { 'Content-Type': 'application/json' });
      if (isAuthCodeMode) {
        const authStates = await Promise.all(
          [...sessions.values()].map(s => s.tokenManager.isAuthenticated().catch(() => false))
        );
        const authenticatedCount = authStates.filter(Boolean).length;
        res.end(JSON.stringify({
          status: 'ok',
          service: 'msgraphmcp',
          authMode: 'authorization-code',
          sessions: sessions.size,
          authenticatedSessions: authenticatedCount,
        }));
      } else {
        const anySession = [...sessions.values()][0];
        const authenticated = anySession
          ? await anySession.tokenManager.isAuthenticated().catch(() => false)
          : undefined;
        res.end(JSON.stringify({
          status: 'ok',
          service: 'msgraphmcp',
          sessions: sessions.size,
          authMode: anySession?.tokenManager.authMode ?? 'per-session',
          ...(authenticated !== undefined && { authenticated }),
        }));
      }
      return;
    }

    // ── OAuth Authorization Code: start flow ────────────────────────────────
    // GET /auth/login?token=<one-time-token>  → redirects to Microsoft login page.
    // The token is a short-lived (15 min), one-time-use value that maps server-side to
    // the MCP session ID — the session ID itself is never exposed in URLs.
    // Prerequisites: AZURE_REDIRECT_URI + AZURE_CLIENT_SECRET env vars set.
    if (url.pathname === '/auth/login') {
      if (!isAuthCodeMode) {
        res.writeHead(400, { 'Content-Type': 'text/plain' });
        res.end('Authorization code mode not configured. Set AZURE_REDIRECT_URI and AZURE_CLIENT_SECRET.');
        return;
      }

      // Look up the one-time login token — delete immediately to prevent replay.
      const token = url.searchParams.get('token');
      const loginEntry = token ? loginTokens.get(token) : undefined;
      if (!token || !loginEntry || loginEntry.expiresAt < Date.now()) {
        res.writeHead(400, { 'Content-Type': 'text/plain' });
        res.end(
          'Invalid or expired login token.\n' +
          'Trigger a tool call via your MCP client to receive a fresh login URL.'
        );
        return;
      }
      loginTokens.delete(token); // one-time use — consumed immediately

      const sessionId = loginEntry.sessionId;
      const session = sessions.get(sessionId);
      if (!session) {
        res.writeHead(400, { 'Content-Type': 'text/plain' });
        res.end('MCP session has expired or disconnected. Please reconnect your MCP client.');
        return;
      }

      // Reject login attempts for already-authenticated sessions — prevents token
      // replacement even if a token was somehow obtained for an active session.
      const alreadyAuthenticated = await session.tokenManager.isAuthenticated().catch(() => false);
      if (alreadyAuthenticated) {
        res.writeHead(409, { 'Content-Type': 'text/plain' });
        res.end('Session is already authenticated. Disconnect and reconnect to start a new session.');
        return;
      }

      try {
        purgeStalePending();
        const codeVerifier = generateCodeVerifier();
        const codeChallenge = generateCodeChallenge(codeVerifier);
        const state = randomUUID();

        // Cancel any existing pending auth for this session — prevents multiple
        // simultaneous OAuth flows for the same session (e.g. user clicks login twice).
        for (const [key, val] of pendingAuth) {
          if (val.sessionId === sessionId) pendingAuth.delete(key);
        }

        pendingAuth.set(state, { codeVerifier, sessionId, expiresAt: Date.now() + 10 * 60 * 1000 });

        const authUrl = await session.tokenManager.getAuthCodeUrl(REDIRECT_URI, state, codeChallenge);
        logger.info('auth: redirecting to microsoft login', { state, sessionId });
        res.writeHead(302, { Location: authUrl });
        res.end();
      } catch (err) {
        logger.error('auth: failed to build auth URL', { error: String(err) });
        res.writeHead(500, { 'Content-Type': 'text/plain' });
        res.end('Failed to initiate authentication: ' + String(err));
      }
      return;
    }

    // ── OAuth Authorization Code: callback ──────────────────────────────────
    // GET /auth/callback?code=...&state=...  (Microsoft redirects here after login)
    if (url.pathname === '/auth/callback') {
      const code = url.searchParams.get('code');
      const state = url.searchParams.get('state');
      const error = url.searchParams.get('error');
      const errorDesc = url.searchParams.get('error_description');

      if (error) {
        logger.warn('auth: callback error from microsoft', { error, errorDesc });
        res.writeHead(400, { 'Content-Type': 'text/html; charset=utf-8' });
        res.end(errorPage(`${error}: ${errorDesc ?? 'unknown error'}`));
        return;
      }

      const pending = state ? pendingAuth.get(state) : undefined;
      if (!code || !pending) {
        res.writeHead(400, { 'Content-Type': 'text/plain' });
        res.end('Invalid callback: missing authorization code or unknown state parameter.');
        return;
      }

      pendingAuth.delete(state!);

      const session = sessions.get(pending.sessionId);
      if (!session) {
        res.writeHead(400, { 'Content-Type': 'text/html; charset=utf-8' });
        res.end(errorPage(
          'MCP session has expired or disconnected. Please reconnect your MCP client and authenticate again.'
        ));
        return;
      }

      try {
        await session.tokenManager.acquireTokenByAuthCode(code, REDIRECT_URI, pending.codeVerifier);
        const accountInfo = await session.tokenManager.getAccountInfo().catch(() => null);
        logger.info('auth: authorization code exchange successful', {
          sessionId: pending.sessionId,
          user: accountInfo?.upn,
        });
        res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
        res.end(successPage());
      } catch (err) {
        logger.error('auth: token exchange failed', { error: String(err) });
        res.writeHead(500, { 'Content-Type': 'text/html; charset=utf-8' });
        res.end(errorPage('Token exchange failed: ' + String(err)));
      }
      return;
    }

    if (url.pathname !== '/mcp') {
      res.writeHead(404, { 'Content-Type': 'text/plain' });
      res.end('Not Found');
      return;
    }

    const incomingSessionId = req.headers['mcp-session-id'] as string | undefined;
    try {
      const existingSession = incomingSessionId ? sessions.get(incomingSessionId) : undefined;
      let transport = existingSession?.transport;

      // If the client provides a session ID that no longer exists (e.g. after a pod restart),
      // return 404 so the client knows to re-initialize rather than sending tool calls to a
      // brand-new uninitialised transport, which would yield "Server not initialized" errors.
      if (incomingSessionId && !existingSession) {
        res.writeHead(404, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: 'Session not found', sessionId: incomingSessionId }));
        return;
      }

      if (!transport) {
        if (sessions.size >= MAX_SESSIONS) {
          logger.warn('mcp session limit reached', { limit: MAX_SESSIONS, active: sessions.size });
          res.writeHead(503, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ error: 'Service Unavailable', message: 'Maximum concurrent sessions reached. Try again later.' }));
          return;
        }

        // Each MCP session gets its own isolated TokenManager and GraphClient.
        // In auth-code mode this ensures tokens are never shared across users —
        // each session authenticates independently via a one-time login token.
        const sessionTokenManager = new TokenManager({ persistCache: false });

        // getLoginUrl issues a fresh one-time login token on each call.
        // sessionId is set by onsessioninitialized before any tool call can fire.
        let sessionId: string | undefined;
        const getLoginUrl = isAuthCodeMode
          ? () => generateLoginUrl(sessionId!)
          : undefined;

        const sessionGraphClient = new GraphClient(sessionTokenManager, getLoginUrl);

        let resolveSessionId: (id: string) => void;
        const sessionIdPromise = new Promise<string>((r) => { resolveSessionId = r; });

        transport = new StreamableHTTPServerTransport({
          sessionIdGenerator: () => randomUUID(),
          onsessioninitialized: (id) => {
            sessionId = id;
            sessions.set(id, { transport: transport!, tokenManager: sessionTokenManager, graphClient: sessionGraphClient, lastActivityAt: Date.now() });
            resolveSessionId(id);
            logger.info('mcp session opened', { sessionId: id, activeSessions: sessions.size });
          },
          onsessionclosed: (id) => {
            sessions.delete(id);
            logger.info('mcp session closed', { sessionId: id, activeSessions: sessions.size });
          },
        });

        const mcpServer = createMcpServer(sessionGraphClient);
        mcpServer.server.oninitialized = () => {
          const ci = mcpServer.server.getClientVersion();
          sessionIdPromise.then((sid) => {
            logger.info('mcp client connected', {
              client: ci?.name,
              clientVersion: ci?.version,
              sessionId: sid,
            });
          });
        };

        await mcpServer.connect(transport);
      }

      // Parse body only for POST requests
      const body = req.method === 'POST' ? await parseBody(req) : undefined;
      await transport.handleRequest(req, res, body);

      // Update last-activity timestamp so idle-timeout checker keeps the session alive.
      if (incomingSessionId) {
        const session = sessions.get(incomingSessionId);
        if (session) session.lastActivityAt = Date.now();
      }

    } catch (err) {
      if (err instanceof AuthRequiredError) {
        const loginUrl = incomingSessionId
          ? generateLoginUrl(incomingSessionId)
          : BASE_LOGIN_URL;
        logger.warn('auth: unauthenticated mcp request', { sessionId: incomingSessionId });
        if (!res.headersSent) {
          res.writeHead(401, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({
            error: 'Unauthorized',
            message: err.message,
            loginUrl,
          }));
        }
      } else {
        logger.error('mcp request error', { error: String(err) });
        if (!res.headersSent) {
          res.writeHead(500, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ error: 'Internal Server Error' }));
        }
      }
    }
  });

  httpServer.listen(port, '0.0.0.0', () => {
    logger.info('http server listening', {
      port,
      endpoints: isAuthCodeMode
        ? ['GET /auth/login', 'GET /auth/callback', 'POST /mcp', 'GET /mcp', 'DELETE /mcp', 'GET /health']
        : ['POST /mcp', 'GET /mcp', 'DELETE /mcp', 'GET /health'],
    });
    if (isAuthCodeMode) {
      logger.info('visit to authenticate', {
        url: REDIRECT_URI.replace('/auth/callback', '/auth/login'),
      });
    }
  });
}

// ── HTML helpers ──────────────────────────────────────────────────────────────

function successPage(): string {
  return `<!DOCTYPE html>
<html lang="en">
<head><meta charset="utf-8"><title>Signed in</title>
<style>body{font-family:system-ui,sans-serif;display:flex;align-items:center;justify-content:center;
min-height:100vh;margin:0;background:#f0fdf4}
.card{background:#fff;border-radius:12px;padding:2.5rem 3rem;box-shadow:0 4px 24px #0001;text-align:center}
h1{color:#16a34a;margin:0 0 .5rem}p{color:#555;margin:0}</style></head>
<body><div class="card">
<h1>&#10003; Signed in successfully</h1>
<p>You can close this window and return to Claude.</p>
</div></body></html>`;
}

function errorPage(detail: string): string {
  return `<!DOCTYPE html>
<html lang="en">
<head><meta charset="utf-8"><title>Auth error</title>
<style>body{font-family:system-ui,sans-serif;display:flex;align-items:center;justify-content:center;
min-height:100vh;margin:0;background:#fef2f2}
.card{background:#fff;border-radius:12px;padding:2.5rem 3rem;box-shadow:0 4px 24px #0001;text-align:center}
h1{color:#dc2626;margin:0 0 .5rem}p{color:#555;margin:0;font-size:.9rem}</style></head>
<body><div class="card">
<h1>Authentication failed</h1>
<p>${detail.replace(/</g, '&lt;')}</p>
</div></body></html>`;
}

// ── Entry point ───────────────────────────────────────────────────────────────

async function main(): Promise<void> {
  const requiredEnv = ['AZURE_CLIENT_ID'];
  for (const key of requiredEnv) {
    if (!process.env[key]) {
      process.stderr.write(`ERROR: ${key} environment variable is required\n`);
      process.exit(1);
    }
  }

  const portEnv = process.env.PORT;
  if (portEnv) {
    const port = parseInt(portEnv, 10);
    if (isNaN(port) || port < 1 || port > 65535) {
      process.stderr.write(`ERROR: PORT must be a valid port number, got: ${portEnv}\n`);
      process.exit(1);
    }
    await startHttp(port);
  } else {
    const tokenManager = new TokenManager();
    const graphClient = new GraphClient(tokenManager);
    await startStdio(graphClient, tokenManager);
  }
}

main().catch((err) => {
  logger.error('fatal error', { error: String(err) });
  process.exit(1);
});
