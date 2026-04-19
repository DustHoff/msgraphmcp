import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { GraphClient } from '../graph/GraphClient';
import { registerUserTools } from './users';
import { registerMailTools } from './mail';
import { registerCalendarTools } from './calendar';
import { registerFileTools } from './files';
import { registerGroupTools } from './groups';
import { registerTeamsTools } from './teams';
import { registerContactTools } from './contacts';
import { registerTaskTools } from './tasks';
import { registerSiteTools } from './sites';
import { registerIntuneTools } from './intune';

export function registerAllTools(server: McpServer, graph: GraphClient): void {
  registerUserTools(server, graph);
  registerMailTools(server, graph);
  registerCalendarTools(server, graph);
  registerFileTools(server, graph);
  registerGroupTools(server, graph);
  registerTeamsTools(server, graph);
  registerContactTools(server, graph);
  registerTaskTools(server, graph);
  registerSiteTools(server, graph);
  registerIntuneTools(server, graph);
}
