import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { TokenManager } from './auth/TokenManager';
import { GraphClient } from './graph/GraphClient';
import { registerAllTools } from './tools/index';
import { logger } from './logger';

async function main() {
  const requiredEnv = ['AZURE_CLIENT_ID'];
  for (const key of requiredEnv) {
    if (!process.env[key]) {
      process.stderr.write(`ERROR: ${key} environment variable is required\n`);
      process.exit(1);
    }
  }

  const tokenManager = new TokenManager();
  const graphClient = new GraphClient(tokenManager);

  const server = new McpServer({
    name: 'msgraphmcp',
    version: '1.0.0',
  });

  // Log client name + version as soon as the MCP handshake completes
  server.server.oninitialized = () => {
    const clientInfo = server.server.getClientVersion();
    const capabilities = server.server.getClientCapabilities();
    logger.info('mcp client connected', {
      client: clientInfo?.name,
      clientVersion: clientInfo?.version,
      experimental: capabilities?.experimental,
    });
  };

  registerAllTools(server, graphClient);

  logger.info('authenticating with microsoft graph');
  try {
    await tokenManager.getAccessToken();
    logger.info('authentication successful — mcp server ready');
  } catch (err) {
    logger.error('authentication failed', { error: String(err) });
    process.exit(1);
  }

  const transport = new StdioServerTransport();
  await server.connect(transport);
}

main().catch((err) => {
  logger.error('fatal error', { error: String(err) });
  process.exit(1);
});
