import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { GraphClient } from '../graph/GraphClient';
import { encodeId, needsEventualConsistency } from './shared';

export function registerGroupTools(server: McpServer, graph: GraphClient) {
  server.tool(
    'list_groups',
    'List groups in the directory.',
    {
      filter: z.string().optional().describe("OData filter, e.g. \"groupTypes/any(c:c eq 'Unified')\""),
      select: z.string().optional(),
      top: z.number().int().min(1).max(999).default(50),
      search: z.string().optional(),
    },
    async ({ filter, select, top, search }) => {
      const params: Record<string, unknown> = { $top: top };
      if (filter) params['$filter'] = filter;
      if (select) params['$select'] = select;
      if (search) {
        params['$search'] = search;
        params['$count'] = true;
      }
      const config = needsEventualConsistency(params)
        ? { headers: { ConsistencyLevel: 'eventual' } }
        : undefined;
      const groups = await graph.get('/groups', params, config);
      return { content: [{ type: 'text', text: JSON.stringify(groups, null, 2) }] };
    }
  );

  server.tool(
    'get_group',
    'Get a single group by id.',
    {
      groupId: z.string(),
      select: z.string().optional(),
    },
    async ({ groupId, select }) => {
      const group = await graph.get(`/groups/${encodeId(groupId)}`, select ? { $select: select } : undefined);
      return { content: [{ type: 'text', text: JSON.stringify(group, null, 2) }] };
    }
  );

  server.tool(
    'create_group',
    'Create a new group (Microsoft 365 or Security).',
    {
      displayName: z.string(),
      mailNickname: z.string(),
      description: z.string().optional(),
      groupType: z.enum(['Microsoft365', 'Security']).default('Microsoft365'),
      mailEnabled: z.boolean().optional(),
      securityEnabled: z.boolean().optional(),
      visibility: z.enum(['Public', 'Private', 'HiddenMembership']).optional(),
    },
    async ({ displayName, mailNickname, description, groupType, mailEnabled, securityEnabled, visibility }) => {
      const isM365 = groupType === 'Microsoft365';
      const body: Record<string, unknown> = {
        displayName,
        mailNickname,
        groupTypes: isM365 ? ['Unified'] : [],
        mailEnabled: mailEnabled ?? isM365,
        securityEnabled: securityEnabled ?? !isM365,
      };
      if (description) body.description = description;
      if (visibility) body.visibility = visibility;

      const group = await graph.post('/groups', body);
      return { content: [{ type: 'text', text: JSON.stringify(group, null, 2) }] };
    }
  );

  server.tool(
    'update_group',
    'Update group properties.',
    {
      groupId: z.string(),
      displayName: z.string().optional(),
      description: z.string().optional(),
      visibility: z.enum(['Public', 'Private', 'HiddenMembership']).optional(),
      mailNickname: z.string().optional(),
    },
    async ({ groupId, ...props }) => {
      const body = Object.fromEntries(Object.entries(props).filter(([, v]) => v !== undefined));
      await graph.patch(`/groups/${encodeId(groupId)}`, body);
      return { content: [{ type: 'text', text: `Group ${groupId} updated.` }] };
    }
  );

  server.tool(
    'delete_group',
    'Delete a group.',
    { groupId: z.string() },
    async ({ groupId }) => {
      await graph.delete(`/groups/${encodeId(groupId)}`);
      return { content: [{ type: 'text', text: `Group ${groupId} deleted.` }] };
    }
  );

  server.tool(
    'list_group_members',
    'List members of a group.',
    {
      groupId: z.string(),
      select: z.string().optional(),
    },
    async ({ groupId, select }) => {
      const members = await graph.getAll(
        `/groups/${encodeId(groupId)}/members`,
        select ? { $select: select } : undefined
      );
      return { content: [{ type: 'text', text: JSON.stringify(members, null, 2) }] };
    }
  );

  server.tool(
    'add_group_member',
    'Add a user (or other directory object) to a group.',
    {
      groupId: z.string(),
      memberId: z.string().describe('Object id of the user/service principal to add'),
    },
    async ({ groupId, memberId }) => {
      await graph.post(`/groups/${encodeId(groupId)}/members/$ref`, {
        '@odata.id': `https://graph.microsoft.com/v1.0/directoryObjects/${encodeId(memberId)}`,
      });
      return { content: [{ type: 'text', text: `Member ${memberId} added to group ${groupId}.` }] };
    }
  );

  server.tool(
    'remove_group_member',
    'Remove a member from a group.',
    {
      groupId: z.string(),
      memberId: z.string(),
    },
    async ({ groupId, memberId }) => {
      await graph.delete(`/groups/${encodeId(groupId)}/members/${encodeId(memberId)}/$ref`);
      return { content: [{ type: 'text', text: `Member ${memberId} removed from group ${groupId}.` }] };
    }
  );

  server.tool(
    'list_group_owners',
    'List owners of a group.',
    { groupId: z.string() },
    async ({ groupId }) => {
      const owners = await graph.getAll(`/groups/${encodeId(groupId)}/owners`);
      return { content: [{ type: 'text', text: JSON.stringify(owners, null, 2) }] };
    }
  );

  server.tool(
    'add_group_owner',
    'Add an owner to a group.',
    {
      groupId: z.string(),
      ownerId: z.string().describe('Object id of the user to add as owner'),
    },
    async ({ groupId, ownerId }) => {
      await graph.post(`/groups/${encodeId(groupId)}/owners/$ref`, {
        '@odata.id': `https://graph.microsoft.com/v1.0/users/${encodeId(ownerId)}`,
      });
      return { content: [{ type: 'text', text: `Owner ${ownerId} added to group ${groupId}.` }] };
    }
  );
}
