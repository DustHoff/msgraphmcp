import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { TokenManager } from './auth/TokenManager';
import { GraphClient } from './graph/GraphClient';
import { registerAllTools } from './tools/index';

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

  registerAllTools(server, graphClient);

  // Warm up authentication before accepting MCP requests.
  // The device code prompt is written to stderr and does not interfere with
  // the MCP stdio protocol running on stdout/stdin.
  process.stderr.write('Initializing Microsoft Graph authentication...\n');
  try {
    await tokenManager.getAccessToken();
    process.stderr.write('Authentication successful. MCP server ready.\n');
  } catch (err) {
    process.stderr.write(`Authentication failed: ${err}\n`);
    process.exit(1);
  }

  const transport = new StdioServerTransport();
  await server.connect(transport);
}

main().catch((err) => {
  process.stderr.write(`Fatal error: ${err}\n`);
  process.exit(1);
});
