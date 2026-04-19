import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { GraphClient } from '../graph/GraphClient';

export function registerSiteTools(server: McpServer, graph: GraphClient) {
  server.tool(
    'list_sites',
    'List SharePoint sites in the tenant.',
    {
      filter: z.string().optional(),
      top: z.number().int().min(1).max(200).default(25),
    },
    async ({ filter, top }) => {
      const params: Record<string, unknown> = { $top: top };
      if (filter) params['$filter'] = filter;
      const sites = await graph.get('/sites', params);
      return { content: [{ type: 'text', text: JSON.stringify(sites, null, 2) }] };
    }
  );

  server.tool(
    'get_site',
    'Get a SharePoint site by id or by hostname+path.',
    {
      siteId: z.string().optional().describe('Site id (GUID or hostname,sitePath)'),
      hostname: z.string().optional().describe('e.g. contoso.sharepoint.com'),
      sitePath: z.string().optional().describe('e.g. /sites/Marketing'),
    },
    async ({ siteId, hostname, sitePath }) => {
      let url: string;
      if (siteId) {
        url = `/sites/${siteId}`;
      } else if (hostname && sitePath) {
        url = `/sites/${hostname}:${sitePath}`;
      } else {
        url = '/sites/root';
      }
      const site = await graph.get(url);
      return { content: [{ type: 'text', text: JSON.stringify(site, null, 2) }] };
    }
  );

  server.tool(
    'search_sites',
    'Search for SharePoint sites by keyword.',
    { query: z.string() },
    async ({ query }) => {
      const sites = await graph.get(`/sites?search=${encodeURIComponent(query)}`);
      return { content: [{ type: 'text', text: JSON.stringify(sites, null, 2) }] };
    }
  );

  server.tool(
    'list_site_lists',
    'List lists/libraries within a SharePoint site.',
    { siteId: z.string() },
    async ({ siteId }) => {
      const lists = await graph.getAll(`/sites/${siteId}/lists`);
      return { content: [{ type: 'text', text: JSON.stringify(lists, null, 2) }] };
    }
  );

  server.tool(
    'get_site_list',
    'Get a specific list in a SharePoint site.',
    {
      siteId: z.string(),
      listId: z.string().describe('List id or display name'),
    },
    async ({ siteId, listId }) => {
      const list = await graph.get(`/sites/${siteId}/lists/${listId}`);
      return { content: [{ type: 'text', text: JSON.stringify(list, null, 2) }] };
    }
  );

  server.tool(
    'list_site_list_items',
    'List items in a SharePoint list.',
    {
      siteId: z.string(),
      listId: z.string(),
      filter: z.string().optional(),
      top: z.number().int().min(1).max(999).default(50),
      expand: z.string().optional().describe("e.g. 'fields' to include field values"),
    },
    async ({ siteId, listId, filter, top, expand }) => {
      const params: Record<string, unknown> = { $top: top };
      if (filter) params['$filter'] = filter;
      if (expand) params['$expand'] = expand;
      const items = await graph.get(`/sites/${siteId}/lists/${listId}/items`, params);
      return { content: [{ type: 'text', text: JSON.stringify(items, null, 2) }] };
    }
  );

  server.tool(
    'get_site_list_item',
    'Get a single SharePoint list item.',
    {
      siteId: z.string(),
      listId: z.string(),
      itemId: z.string(),
    },
    async ({ siteId, listId, itemId }) => {
      const item = await graph.get(`/sites/${siteId}/lists/${listId}/items/${itemId}?$expand=fields`);
      return { content: [{ type: 'text', text: JSON.stringify(item, null, 2) }] };
    }
  );

  server.tool(
    'create_site_list_item',
    'Create a new item in a SharePoint list.',
    {
      siteId: z.string(),
      listId: z.string(),
      fields: z.record(z.unknown()).describe('Key-value pairs of column names and values'),
    },
    async ({ siteId, listId, fields }) => {
      const item = await graph.post(`/sites/${siteId}/lists/${listId}/items`, { fields });
      return { content: [{ type: 'text', text: JSON.stringify(item, null, 2) }] };
    }
  );

  server.tool(
    'update_site_list_item',
    'Update fields on a SharePoint list item.',
    {
      siteId: z.string(),
      listId: z.string(),
      itemId: z.string(),
      fields: z.record(z.unknown()),
    },
    async ({ siteId, listId, itemId, fields }) => {
      const item = await graph.patch(`/sites/${siteId}/lists/${listId}/items/${itemId}/fields`, fields);
      return { content: [{ type: 'text', text: JSON.stringify(item, null, 2) }] };
    }
  );

  server.tool(
    'delete_site_list_item',
    'Delete a SharePoint list item.',
    {
      siteId: z.string(),
      listId: z.string(),
      itemId: z.string(),
    },
    async ({ siteId, listId, itemId }) => {
      await graph.delete(`/sites/${siteId}/lists/${listId}/items/${itemId}`);
      return { content: [{ type: 'text', text: 'List item deleted.' }] };
    }
  );
}
