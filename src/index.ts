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

async function startHttp(port: number): Promise<void> {
  // Map of active sessions: sessionId → transport
  const sessions = new Map<string, StreamableHTTPServerTransport>();

  // ── Auth-code mode: single global TokenManager with persistent cache ───────
  // In authorization-code mode all sessions share one authenticated identity.
  // The user signs in once via /auth/login; tokens are cached on disk.
  //
  // Device-code mode keeps per-session isolated in-memory caches to prevent
  // delegated token bleed between concurrent users.
  const isAuthCodeMode = Boolean(process.env.AZURE_REDIRECT_URI && process.env.AZURE_CLIENT_SECRET);
  const REDIRECT_URI = process.env.AZURE_REDIRECT_URI ?? '';

  let globalTokenManager: TokenManager | undefined;
  let globalGraphClient: GraphClient | undefined;

  if (isAuthCodeMode) {
    globalTokenManager = new TokenManager({ persistCache: true });
    globalGraphClient = new GraphClient(globalTokenManager);
    logger.info('auth mode: authorization-code (delegated)', {
      redirectUri: REDIRECT_URI,
      loginUrl: REDIRECT_URI.replace('/auth/callback', '/auth/login'),
    });
  }

  // Pending OAuth state: state-value → { codeVerifier }
  // Entries expire after 10 minutes to avoid unbounded growth.
  const pendingAuth = new Map<string, { codeVerifier: string; expiresAt: number }>();

  function purgeStalePending() {
    const now = Date.now();
    for (const [key, val] of pendingAuth) {
      if (val.expiresAt < now) pendingAuth.delete(key);
    }
  }

  const httpServer = http.createServer(async (req: IncomingMessage, res: ServerResponse) => {
    const url = new URL(req.url ?? '/', `http://localhost:${port}`);

    // ── Health/readiness probe ───────────────────────────────────────────────
    if (url.pathname === '/health') {
      const authenticated = globalTokenManager
        ? await globalTokenManager.isAuthenticated().catch(() => false)
        : undefined;
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({
        status: 'ok',
        service: 'msgraphmcp',
        sessions: sessions.size,
        authMode: globalTokenManager?.authMode ?? 'per-session',
        ...(authenticated !== undefined && { authenticated }),
        ...(isAuthCodeMode && !authenticated && {
          loginUrl: REDIRECT_URI.replace('/auth/callback', '/auth/login'),
        }),
      }));
      return;
    }

    // ── OAuth Authorization Code: start flow ────────────────────────────────
    // GET /auth/login  → redirects to Microsoft login page
    // Prerequisites: AZURE_REDIRECT_URI + AZURE_CLIENT_SECRET env vars set.
    // The redirect URI must be registered in the Entra ID app registration.
    if (url.pathname === '/auth/login') {
      if (!isAuthCodeMode || !globalTokenManager) {
        res.writeHead(400, { 'Content-Type': 'text/plain' });
        res.end('Authorization code mode not configured. Set AZURE_REDIRECT_URI and AZURE_CLIENT_SECRET.');
        return;
      }
      try {
        purgeStalePending();
        const codeVerifier = generateCodeVerifier();
        const codeChallenge = generateCodeChallenge(codeVerifier);
        const state = randomUUID();
        pendingAuth.set(state, { codeVerifier, expiresAt: Date.now() + 10 * 60 * 1000 });

        const authUrl = await globalTokenManager.getAuthCodeUrl(REDIRECT_URI, state, codeChallenge);
        logger.info('auth: redirecting to microsoft login', { state });
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

      try {
        await globalTokenManager!.acquireTokenByAuthCode(code, REDIRECT_URI, pending.codeVerifier);
        logger.info('auth: authorization code exchange successful');
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

    try {
      const incomingSessionId = req.headers['mcp-session-id'] as string | undefined;
      let transport = incomingSessionId ? sessions.get(incomingSessionId) : undefined;

      // If the client provides a session ID that no longer exists (e.g. after a pod restart),
      // return 404 so the client knows to re-initialize rather than sending tool calls to a
      // brand-new uninitialised transport, which would yield "Server not initialized" errors.
      if (incomingSessionId && !transport) {
        res.writeHead(404, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: 'Session not found', sessionId: incomingSessionId }));
        return;
      }

      if (!transport) {
        // In auth-code mode all sessions share the global authenticated client.
        // In device-code mode each session gets its own isolated in-memory cache.
        const sessionTokenManager = isAuthCodeMode
          ? globalTokenManager!
          : new TokenManager({ persistCache: false });
        const sessionGraphClient = isAuthCodeMode
          ? globalGraphClient!
          : new GraphClient(sessionTokenManager);

        let resolveSessionId: (id: string) => void;
        const sessionIdPromise = new Promise<string>((r) => { resolveSessionId = r; });

        transport = new StreamableHTTPServerTransport({
          sessionIdGenerator: () => randomUUID(),
          onsessioninitialized: (id) => {
            sessions.set(id, transport!);
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

    } catch (err) {
      if (err instanceof AuthRequiredError) {
        const loginUrl = REDIRECT_URI.replace('/auth/callback', '/auth/login');
        logger.warn('auth: unauthenticated mcp request', { loginUrl });
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
