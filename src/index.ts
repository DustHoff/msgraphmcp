import http, { IncomingMessage, ServerResponse } from 'http';
import { randomUUID } from 'crypto';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js';
import { TokenManager } from './auth/TokenManager';
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
  // Each session owns its own TokenManager + GraphClient so delegated tokens
  // are never shared across users.
  const sessions = new Map<string, StreamableHTTPServerTransport>();

  const httpServer = http.createServer(async (req: IncomingMessage, res: ServerResponse) => {
    const url = new URL(req.url ?? '/', `http://localhost:${port}`);

    // ── Health/readiness probe ───────────────────────────────────────────────
    if (url.pathname === '/health') {
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ status: 'ok', service: 'msgraphmcp', sessions: sessions.size }));
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

      if (!transport) {
        // New session: isolated TokenManager (in-memory cache) + GraphClient so
        // that User A's tokens can never bleed into User B's session.
        const sessionTokenManager = new TokenManager({ persistCache: false });
        const sessionGraphClient = new GraphClient(sessionTokenManager);

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
      logger.error('mcp request error', { error: String(err) });
      if (!res.headersSent) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: 'Internal Server Error' }));
      }
    }
  });

  httpServer.listen(port, '0.0.0.0', () => {
    logger.info('http server listening', {
      port,
      endpoints: ['POST /mcp', 'GET /mcp', 'DELETE /mcp', 'GET /health'],
    });
  });
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
