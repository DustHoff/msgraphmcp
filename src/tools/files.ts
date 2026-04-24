import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { GraphClient } from '../graph/GraphClient';
import { userPath, odataQuote, encodeDrivePath } from './shared';

export function registerFileTools(server: McpServer, graph: GraphClient) {
  server.tool(
    'list_drive_items',
    'List items in a OneDrive folder.',
    {
      userId: z.string().default('me').describe('User id or "me"'),
      itemPath: z.string().default('/').describe('Folder path relative to drive root, e.g. "/Documents"'),
      top: z.number().int().min(1).max(999).default(50),
    },
    async ({ userId, itemPath, top }) => {
      const driveBase = `${userPath(userId)}/drive`;
      const url = itemPath === '/'
        ? `${driveBase}/root/children`
        : `${driveBase}/root:${encodeDrivePath(itemPath)}:/children`;
      const items = await graph.get(url, { $top: top });
      return { content: [{ type: 'text', text: JSON.stringify(items, null, 2) }] };
    }
  );

  server.tool(
    'get_drive_item',
    'Get metadata for a OneDrive item by path or id.',
    {
      userId: z.string().default('me'),
      itemPath: z.string().optional().describe('Path relative to drive root, e.g. "/Documents/file.docx"'),
      itemId: z.string().optional().describe('Item id (alternative to itemPath)'),
    },
    async ({ userId, itemPath, itemId }) => {
      if (!itemPath && !itemId) {
        throw new Error('Either itemPath or itemId must be provided.');
      }
      const driveBase = `${userPath(userId)}/drive`;
      const url = itemId
        ? `${driveBase}/items/${encodeURIComponent(itemId)}`
        : `${driveBase}/root:${encodeDrivePath(itemPath!)}`;
      const item = await graph.get(url);
      return { content: [{ type: 'text', text: JSON.stringify(item, null, 2) }] };
    }
  );

  server.tool(
    'create_drive_folder',
    'Create a folder in OneDrive.',
    {
      userId: z.string().default('me'),
      parentPath: z.string().default('/').describe('Parent folder path, e.g. "/Documents"'),
      folderName: z.string(),
      conflictBehavior: z.enum(['rename', 'fail', 'replace']).default('rename'),
    },
    async ({ userId, parentPath, folderName, conflictBehavior }) => {
      const driveBase = `${userPath(userId)}/drive`;
      const url = parentPath === '/'
        ? `${driveBase}/root/children`
        : `${driveBase}/root:${encodeDrivePath(parentPath)}:/children`;
      const folder = await graph.post(url, {
        name: folderName,
        folder: {},
        '@microsoft.graph.conflictBehavior': conflictBehavior,
      });
      return { content: [{ type: 'text', text: JSON.stringify(folder, null, 2) }] };
    }
  );

  server.tool(
    'upload_drive_file',
    'Upload a small file (≤4 MB) to OneDrive. For larger files use upload sessions.',
    {
      userId: z.string().default('me'),
      filePath: z.string().describe('Destination path including filename, e.g. "/Documents/report.txt"'),
      content: z.string().describe('File content (text)'),
      conflictBehavior: z.enum(['rename', 'fail', 'replace']).default('replace'),
    },
    async ({ userId, filePath, content, conflictBehavior }) => {
      const url = `${userPath(userId)}/drive/root:${encodeDrivePath(filePath)}:/content`;
      const item = await graph.put(url, content, {
        params: { '@microsoft.graph.conflictBehavior': conflictBehavior },
        headers: { 'Content-Type': 'text/plain' },
      });
      return { content: [{ type: 'text', text: JSON.stringify(item, null, 2) }] };
    }
  );

  server.tool(
    'delete_drive_item',
    'Delete a OneDrive item.',
    {
      userId: z.string().default('me'),
      itemPath: z.string().optional(),
      itemId: z.string().optional(),
    },
    async ({ userId, itemPath, itemId }) => {
      if (!itemPath && !itemId) {
        throw new Error('Either itemPath or itemId must be provided.');
      }
      const driveBase = `${userPath(userId)}/drive`;
      const url = itemId
        ? `${driveBase}/items/${encodeURIComponent(itemId)}`
        : `${driveBase}/root:${encodeDrivePath(itemPath!)}`;
      await graph.delete(url);
      return { content: [{ type: 'text', text: 'Item deleted.' }] };
    }
  );

  server.tool(
    'copy_drive_item',
    'Copy a OneDrive item to another location.',
    {
      userId: z.string().default('me'),
      itemId: z.string().describe('Source item id'),
      destinationParentId: z.string().describe('Destination parent folder id'),
      newName: z.string().optional().describe('New filename (optional)'),
    },
    async ({ userId, itemId, destinationParentId, newName }) => {
      const body: Record<string, unknown> = {
        parentReference: { id: destinationParentId },
      };
      if (newName) body.name = newName;
      const result = await graph.post(
        `${userPath(userId)}/drive/items/${encodeURIComponent(itemId)}/copy`,
        body
      );
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }
  );

  server.tool(
    'search_drive',
    'Search for files/folders in OneDrive.',
    {
      userId: z.string().default('me'),
      query: z.string().describe('Search query'),
      top: z.number().int().min(1).max(200).default(25),
    },
    async ({ userId, query, top }) => {
      // Single quotes inside OData string literals are escaped by doubling
      // per the OData spec. Without this, a query containing `'` would
      // truncate the expression and return an error or unintended results.
      const quoted = encodeURIComponent(odataQuote(query));
      const results = await graph.get(
        `${userPath(userId)}/drive/root/search(q='${quoted}')`,
        { $top: top }
      );
      return { content: [{ type: 'text', text: JSON.stringify(results, null, 2) }] };
    }
  );

  server.tool(
    'list_shared_with_me',
    'List OneDrive items shared with the signed-in user.',
    { userId: z.string().default('me') },
    async ({ userId }) => {
      const items = await graph.getAll(`${userPath(userId)}/drive/sharedWithMe`);
      return { content: [{ type: 'text', text: JSON.stringify(items, null, 2) }] };
    }
  );
}
