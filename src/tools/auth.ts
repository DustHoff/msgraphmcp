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
        const isAppOnly = status.mode === 'client-secret' || status.mode === 'client-certificate';
        return {
          content: [{
            type: 'text',
            text: JSON.stringify({
              authenticated: true,
              mode: status.mode,
              ...(isAppOnly && {
                message: 'Authenticated as application identity (app-only). ' +
                  'To use delegated (per-user) authentication, set AZURE_REDIRECT_URI on the server.',
              }),
            }),
          }],
        };
      }
      return {
        content: [{
          type: 'text',
          text: JSON.stringify({
            authenticated: false,
            mode: status.mode,
            loginUrl: status.loginUrl ?? null,
            message: status.loginUrl
              ? 'Not authenticated. Visit the loginUrl to sign in with Microsoft, then retry your request.'
              : 'Not authenticated. No login URL available — ensure AZURE_REDIRECT_URI is set on the server.',
          }),
        }],
      };
    },
  );
}
