import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { GraphClient } from '../graph/GraphClient';

const recipientSchema = z.object({
  name: z.string().optional(),
  address: z.string().email(),
});

export function registerMailTools(server: McpServer, graph: GraphClient) {
  server.tool(
    'list_messages',
    'List mail messages in a mailbox folder.',
    {
      userId: z.string().default('me').describe('User id or "me"'),
      folderId: z.string().default('inbox').describe('Folder id or well-known name (inbox, sentitems, drafts, deleteditems)'),
      filter: z.string().optional().describe("OData filter"),
      select: z.string().optional().describe("Comma-separated fields"),
      top: z.number().int().min(1).max(999).default(25),
      search: z.string().optional().describe('Search query'),
    },
    async ({ userId, folderId, filter, select, top, search }) => {
      const params: Record<string, unknown> = { $top: top };
      if (filter) params['$filter'] = filter;
      if (select) params['$select'] = select;
      if (search) params['$search'] = search;
      const messages = await graph.get(
        `/users/${encodeURIComponent(userId)}/mailFolders/${folderId}/messages`,
        params
      );
      return { content: [{ type: 'text', text: JSON.stringify(messages, null, 2) }] };
    }
  );

  server.tool(
    'get_message',
    'Get a specific mail message.',
    {
      userId: z.string().default('me'),
      messageId: z.string(),
    },
    async ({ userId, messageId }) => {
      const msg = await graph.get(`/users/${encodeURIComponent(userId)}/messages/${messageId}`);
      return { content: [{ type: 'text', text: JSON.stringify(msg, null, 2) }] };
    }
  );

  server.tool(
    'send_mail',
    'Send an email message.',
    {
      userId: z.string().default('me').describe('Sender user id or "me"'),
      subject: z.string(),
      body: z.string().describe('Message body content'),
      bodyContentType: z.enum(['Text', 'HTML']).default('Text'),
      toRecipients: z.array(recipientSchema).min(1),
      ccRecipients: z.array(recipientSchema).optional(),
      bccRecipients: z.array(recipientSchema).optional(),
      saveToSentItems: z.boolean().default(true),
    },
    async ({ userId, subject, body, bodyContentType, toRecipients, ccRecipients, bccRecipients, saveToSentItems }) => {
      const toAddr = (r: { name?: string; address: string }) => ({ emailAddress: { name: r.name, address: r.address } });
      const message: Record<string, unknown> = {
        subject,
        body: { contentType: bodyContentType, content: body },
        toRecipients: toRecipients.map(toAddr),
      };
      if (ccRecipients?.length) message.ccRecipients = ccRecipients.map(toAddr);
      if (bccRecipients?.length) message.bccRecipients = bccRecipients.map(toAddr);

      await graph.post(`/users/${encodeURIComponent(userId)}/sendMail`, { message, saveToSentItems });
      return { content: [{ type: 'text', text: 'Mail sent successfully.' }] };
    }
  );

  server.tool(
    'reply_to_message',
    'Reply to a mail message.',
    {
      userId: z.string().default('me'),
      messageId: z.string(),
      comment: z.string().describe('Reply body text'),
    },
    async ({ userId, messageId, comment }) => {
      await graph.post(`/users/${encodeURIComponent(userId)}/messages/${messageId}/reply`, { comment });
      return { content: [{ type: 'text', text: 'Reply sent.' }] };
    }
  );

  server.tool(
    'forward_message',
    'Forward a mail message to recipients.',
    {
      userId: z.string().default('me'),
      messageId: z.string(),
      toRecipients: z.array(recipientSchema).min(1),
      comment: z.string().optional().describe('Additional text prepended to the forwarded message'),
    },
    async ({ userId, messageId, toRecipients, comment }) => {
      const toAddr = (r: { name?: string; address: string }) => ({ emailAddress: { name: r.name, address: r.address } });
      const body: Record<string, unknown> = { toRecipients: toRecipients.map(toAddr) };
      if (comment) body.comment = comment;
      await graph.post(`/users/${encodeURIComponent(userId)}/messages/${messageId}/forward`, body);
      return { content: [{ type: 'text', text: 'Message forwarded.' }] };
    }
  );

  server.tool(
    'delete_message',
    'Delete (move to Deleted Items) a mail message.',
    {
      userId: z.string().default('me'),
      messageId: z.string(),
    },
    async ({ userId, messageId }) => {
      await graph.delete(`/users/${encodeURIComponent(userId)}/messages/${messageId}`);
      return { content: [{ type: 'text', text: 'Message deleted.' }] };
    }
  );

  server.tool(
    'move_message',
    'Move a mail message to a different folder.',
    {
      userId: z.string().default('me'),
      messageId: z.string(),
      destinationFolderId: z.string().describe('Target folder id or well-known name'),
    },
    async ({ userId, messageId, destinationFolderId }) => {
      const msg = await graph.post(
        `/users/${encodeURIComponent(userId)}/messages/${messageId}/move`,
        { destinationId: destinationFolderId }
      );
      return { content: [{ type: 'text', text: JSON.stringify(msg, null, 2) }] };
    }
  );

  server.tool(
    'list_mail_folders',
    'List mail folders for a user.',
    {
      userId: z.string().default('me'),
      includeHiddenFolders: z.boolean().default(false),
    },
    async ({ userId, includeHiddenFolders }) => {
      const folders = await graph.getAll(
        `/users/${encodeURIComponent(userId)}/mailFolders`,
        includeHiddenFolders ? { includeHiddenFolders: true } : undefined
      );
      return { content: [{ type: 'text', text: JSON.stringify(folders, null, 2) }] };
    }
  );

  server.tool(
    'create_mail_folder',
    'Create a new mail folder.',
    {
      userId: z.string().default('me'),
      displayName: z.string(),
      parentFolderId: z.string().optional().describe('Parent folder id (omit for top-level)'),
    },
    async ({ userId, displayName, parentFolderId }) => {
      const url = parentFolderId
        ? `/users/${encodeURIComponent(userId)}/mailFolders/${parentFolderId}/childFolders`
        : `/users/${encodeURIComponent(userId)}/mailFolders`;
      const folder = await graph.post(url, { displayName });
      return { content: [{ type: 'text', text: JSON.stringify(folder, null, 2) }] };
    }
  );
}
