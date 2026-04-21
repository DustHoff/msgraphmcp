import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { GraphClient } from '../graph/GraphClient';

export function registerAuthTools(server: McpServer, graph: GraphClient): void {
  server.tool(
    'get_login_url',
    'Returns authentication status and the login URL if sign-in is required. Call this first if other tools return auth errors.',
    {},
    async () => {
      const status = await graph.getAuthStatus();
      if (status.authenticated) {
        return {
          content: [{ type: 'text', text: JSON.stringify({ authenticated: true }) }],
        };
      }
      return {
        content: [{
          type: 'text',
          text: JSON.stringify({
            authenticated: false,
            loginUrl: status.loginUrl ?? null,
            message: status.loginUrl
              ? `Not authenticated. Visit the loginUrl to sign in with Microsoft, then retry your request.`
              : 'Not authenticated. No login URL available — ensure the MCP server is configured for auth-code mode.',
          }),
        }],
      };
    },
  );
}
