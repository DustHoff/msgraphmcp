import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { GraphClient } from '../graph/GraphClient';
import { userPath } from './shared';

export function registerContactTools(server: McpServer, graph: GraphClient) {
  server.tool(
    'list_contacts',
    'List personal contacts for a user.',
    {
      userId: z.string().default('me'),
      filter: z.string().optional(),
      select: z.string().optional(),
      top: z.number().int().min(1).max(999).default(50),
    },
    async ({ userId, filter, select, top }) => {
      const params: Record<string, unknown> = { $top: top };
      if (filter) params['$filter'] = filter;
      if (select) params['$select'] = select;
      const contacts = await graph.get(`${userPath(userId)}/contacts`, params);
      return { content: [{ type: 'text', text: JSON.stringify(contacts, null, 2) }] };
    }
  );

  server.tool(
    'get_contact',
    'Get a specific contact.',
    {
      userId: z.string().default('me'),
      contactId: z.string(),
    },
    async ({ userId, contactId }) => {
      const contact = await graph.get(`${userPath(userId)}/contacts/${encodeURIComponent(contactId)}`);
      return { content: [{ type: 'text', text: JSON.stringify(contact, null, 2) }] };
    }
  );

  server.tool(
    'create_contact',
    'Create a new contact.',
    {
      userId: z.string().default('me'),
      givenName: z.string().optional(),
      surname: z.string().optional(),
      displayName: z.string().optional(),
      emailAddresses: z.array(z.object({ address: z.string().email(), name: z.string().optional() })).optional(),
      businessPhones: z.array(z.string()).optional(),
      mobilePhone: z.string().optional(),
      jobTitle: z.string().optional(),
      companyName: z.string().optional(),
      department: z.string().optional(),
    },
    async ({ userId, givenName, surname, displayName, emailAddresses, businessPhones, mobilePhone, jobTitle, companyName, department }) => {
      const body: Record<string, unknown> = {};
      if (givenName) body.givenName = givenName;
      if (surname) body.surname = surname;
      if (displayName) body.displayName = displayName;
      if (emailAddresses) body.emailAddresses = emailAddresses;
      if (businessPhones) body.businessPhones = businessPhones;
      if (mobilePhone) body.mobilePhone = mobilePhone;
      if (jobTitle) body.jobTitle = jobTitle;
      if (companyName) body.companyName = companyName;
      if (department) body.department = department;

      const contact = await graph.post(`${userPath(userId)}/contacts`, body);
      return { content: [{ type: 'text', text: JSON.stringify(contact, null, 2) }] };
    }
  );

  server.tool(
    'update_contact',
    'Update an existing contact.',
    {
      userId: z.string().default('me'),
      contactId: z.string(),
      givenName: z.string().optional(),
      surname: z.string().optional(),
      displayName: z.string().optional(),
      emailAddresses: z.array(z.object({ address: z.string().email(), name: z.string().optional() })).optional(),
      businessPhones: z.array(z.string()).optional(),
      mobilePhone: z.string().optional(),
      jobTitle: z.string().optional(),
      companyName: z.string().optional(),
    },
    async ({ userId, contactId, ...props }) => {
      const body = Object.fromEntries(Object.entries(props).filter(([, v]) => v !== undefined));
      const contact = await graph.patch(
        `${userPath(userId)}/contacts/${encodeURIComponent(contactId)}`,
        body
      );
      return { content: [{ type: 'text', text: JSON.stringify(contact, null, 2) }] };
    }
  );

  server.tool(
    'delete_contact',
    'Delete a contact.',
    {
      userId: z.string().default('me'),
      contactId: z.string(),
    },
    async ({ userId, contactId }) => {
      await graph.delete(`${userPath(userId)}/contacts/${encodeURIComponent(contactId)}`);
      return { content: [{ type: 'text', text: 'Contact deleted.' }] };
    }
  );
}
