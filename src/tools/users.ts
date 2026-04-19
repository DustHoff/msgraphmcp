import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { GraphClient } from '../graph/GraphClient';

export function registerUserTools(server: McpServer, graph: GraphClient) {
  server.tool(
    'list_users',
    'List users in the directory. Supports filtering, sorting and field selection.',
    {
      filter: z.string().optional().describe("OData filter, e.g. \"displayName eq 'Alice'\""),
      select: z.string().optional().describe("Comma-separated fields to return, e.g. 'id,displayName,mail'"),
      top: z.number().int().min(1).max(999).optional().describe('Max number of results (1-999)'),
      search: z.string().optional().describe('Search query, e.g. \'"displayName:Alice"\''),
    },
    async ({ filter, select, top, search }) => {
      const params: Record<string, unknown> = {};
      if (filter) params['$filter'] = filter;
      if (select) params['$select'] = select;
      if (top) params['$top'] = top;
      if (search) {
        params['$search'] = search;
        // search requires ConsistencyLevel header – handled via $count
        params['$count'] = true;
      }
      const users = await graph.getAll('/users', params);
      return { content: [{ type: 'text', text: JSON.stringify(users, null, 2) }] };
    }
  );

  server.tool(
    'get_user',
    'Get a single user by id or userPrincipalName.',
    {
      userId: z.string().describe('User id or userPrincipalName. Use "me" for the signed-in user.'),
      select: z.string().optional().describe("Comma-separated fields to return"),
    },
    async ({ userId, select }) => {
      const params = select ? { $select: select } : undefined;
      const user = await graph.get(`/users/${encodeURIComponent(userId)}`, params);
      return { content: [{ type: 'text', text: JSON.stringify(user, null, 2) }] };
    }
  );

  server.tool(
    'create_user',
    'Create a new user in the directory.',
    {
      displayName: z.string().describe('Display name'),
      userPrincipalName: z.string().describe('UPN, e.g. alice@contoso.com'),
      mailNickname: z.string().describe('Mail alias (without @domain)'),
      password: z.string().describe('Initial password'),
      accountEnabled: z.boolean().default(true).describe('Whether account is enabled'),
      givenName: z.string().optional(),
      surname: z.string().optional(),
      jobTitle: z.string().optional(),
      department: z.string().optional(),
      mobilePhone: z.string().optional(),
    },
    async ({ displayName, userPrincipalName, mailNickname, password, accountEnabled, givenName, surname, jobTitle, department, mobilePhone }) => {
      const body: Record<string, unknown> = {
        displayName,
        userPrincipalName,
        mailNickname,
        accountEnabled,
        passwordProfile: { password, forceChangePasswordNextSignIn: true },
      };
      if (givenName) body.givenName = givenName;
      if (surname) body.surname = surname;
      if (jobTitle) body.jobTitle = jobTitle;
      if (department) body.department = department;
      if (mobilePhone) body.mobilePhone = mobilePhone;

      const user = await graph.post('/users', body);
      return { content: [{ type: 'text', text: JSON.stringify(user, null, 2) }] };
    }
  );

  server.tool(
    'update_user',
    'Update properties of a user.',
    {
      userId: z.string().describe('User id or userPrincipalName'),
      displayName: z.string().optional(),
      givenName: z.string().optional(),
      surname: z.string().optional(),
      jobTitle: z.string().optional(),
      department: z.string().optional(),
      mobilePhone: z.string().optional(),
      officeLocation: z.string().optional(),
      businessPhones: z.array(z.string()).optional(),
      accountEnabled: z.boolean().optional(),
    },
    async ({ userId, ...props }) => {
      const body = Object.fromEntries(Object.entries(props).filter(([, v]) => v !== undefined));
      await graph.patch(`/users/${encodeURIComponent(userId)}`, body);
      return { content: [{ type: 'text', text: `User ${userId} updated successfully.` }] };
    }
  );

  server.tool(
    'delete_user',
    'Delete a user from the directory.',
    { userId: z.string().describe('User id or userPrincipalName') },
    async ({ userId }) => {
      await graph.delete(`/users/${encodeURIComponent(userId)}`);
      return { content: [{ type: 'text', text: `User ${userId} deleted.` }] };
    }
  );

  server.tool(
    'get_user_member_of',
    'Get groups and directory roles a user is a member of.',
    {
      userId: z.string().describe('User id or userPrincipalName. Use "me" for the signed-in user.'),
    },
    async ({ userId }) => {
      const groups = await graph.getAll(`/users/${encodeURIComponent(userId)}/memberOf`);
      return { content: [{ type: 'text', text: JSON.stringify(groups, null, 2) }] };
    }
  );

  server.tool(
    'reset_user_password',
    'Reset a user password.',
    {
      userId: z.string().describe('User id or userPrincipalName'),
      newPassword: z.string().describe('New password'),
      forceChangePasswordNextSignIn: z.boolean().default(true),
    },
    async ({ userId, newPassword, forceChangePasswordNextSignIn }) => {
      await graph.patch(`/users/${encodeURIComponent(userId)}`, {
        passwordProfile: { password: newPassword, forceChangePasswordNextSignIn },
      });
      return { content: [{ type: 'text', text: `Password for ${userId} reset successfully.` }] };
    }
  );
}
