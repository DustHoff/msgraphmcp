import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { GraphClient } from '../graph/GraphClient';
import { userPath } from './shared';

export function registerTeamsTools(server: McpServer, graph: GraphClient) {
  server.tool(
    'list_joined_teams',
    'List Microsoft Teams the signed-in user has joined.',
    { userId: z.string().default('me') },
    async ({ userId }) => {
      const teams = await graph.getAll(`${userPath(userId)}/joinedTeams`);
      return { content: [{ type: 'text', text: JSON.stringify(teams, null, 2) }] };
    }
  );

  server.tool(
    'get_team',
    'Get details of a specific team.',
    { teamId: z.string() },
    async ({ teamId }) => {
      const team = await graph.get(`/teams/${teamId}`);
      return { content: [{ type: 'text', text: JSON.stringify(team, null, 2) }] };
    }
  );

  server.tool(
    'create_team',
    'Create a new Microsoft Team.',
    {
      displayName: z.string(),
      description: z.string().optional(),
      visibility: z.enum(['Public', 'Private']).default('Private'),
      template: z.string().default('standard').describe('Team template, e.g. "standard", "educationClass"'),
    },
    async ({ displayName, description, visibility, template }) => {
      const body: Record<string, unknown> = {
        'template@odata.bind': `https://graph.microsoft.com/v1.0/teamsTemplates('${template}')`,
        displayName,
        visibility,
      };
      if (description) body.description = description;
      const team = await graph.post('/teams', body);
      return { content: [{ type: 'text', text: JSON.stringify(team, null, 2) }] };
    }
  );

  server.tool(
    'list_channels',
    'List channels in a team.',
    { teamId: z.string() },
    async ({ teamId }) => {
      const channels = await graph.getAll(`/teams/${teamId}/channels`);
      return { content: [{ type: 'text', text: JSON.stringify(channels, null, 2) }] };
    }
  );

  server.tool(
    'get_channel',
    'Get a specific channel.',
    { teamId: z.string(), channelId: z.string() },
    async ({ teamId, channelId }) => {
      const channel = await graph.get(`/teams/${teamId}/channels/${channelId}`);
      return { content: [{ type: 'text', text: JSON.stringify(channel, null, 2) }] };
    }
  );

  server.tool(
    'create_channel',
    'Create a channel in a team.',
    {
      teamId: z.string(),
      displayName: z.string(),
      description: z.string().optional(),
      membershipType: z.enum(['standard', 'private', 'shared']).default('standard'),
    },
    async ({ teamId, displayName, description, membershipType }) => {
      const body: Record<string, unknown> = { displayName, membershipType };
      if (description) body.description = description;
      const channel = await graph.post(`/teams/${teamId}/channels`, body);
      return { content: [{ type: 'text', text: JSON.stringify(channel, null, 2) }] };
    }
  );

  server.tool(
    'delete_channel',
    'Delete a channel from a team.',
    { teamId: z.string(), channelId: z.string() },
    async ({ teamId, channelId }) => {
      await graph.delete(`/teams/${teamId}/channels/${channelId}`);
      return { content: [{ type: 'text', text: 'Channel deleted.' }] };
    }
  );

  server.tool(
    'list_channel_messages',
    'List messages in a Teams channel.',
    {
      teamId: z.string(),
      channelId: z.string(),
      top: z.number().int().min(1).max(50).default(20),
    },
    async ({ teamId, channelId, top }) => {
      const messages = await graph.get(
        `/teams/${teamId}/channels/${channelId}/messages`,
        { $top: top }
      );
      return { content: [{ type: 'text', text: JSON.stringify(messages, null, 2) }] };
    }
  );

  server.tool(
    'send_channel_message',
    'Send a message to a Teams channel.',
    {
      teamId: z.string(),
      channelId: z.string(),
      content: z.string().describe('Message text'),
      contentType: z.enum(['text', 'html']).default('text'),
    },
    async ({ teamId, channelId, content, contentType }) => {
      const message = await graph.post(`/teams/${teamId}/channels/${channelId}/messages`, {
        body: { content, contentType },
      });
      return { content: [{ type: 'text', text: JSON.stringify(message, null, 2) }] };
    }
  );

  server.tool(
    'reply_to_channel_message',
    'Reply to a message in a Teams channel.',
    {
      teamId: z.string(),
      channelId: z.string(),
      messageId: z.string(),
      content: z.string(),
      contentType: z.enum(['text', 'html']).default('text'),
    },
    async ({ teamId, channelId, messageId, content, contentType }) => {
      const reply = await graph.post(
        `/teams/${teamId}/channels/${channelId}/messages/${messageId}/replies`,
        { body: { content, contentType } }
      );
      return { content: [{ type: 'text', text: JSON.stringify(reply, null, 2) }] };
    }
  );

  server.tool(
    'list_team_members',
    'List members of a team.',
    { teamId: z.string() },
    async ({ teamId }) => {
      const members = await graph.getAll(`/teams/${teamId}/members`);
      return { content: [{ type: 'text', text: JSON.stringify(members, null, 2) }] };
    }
  );

  server.tool(
    'add_team_member',
    'Add a member to a team.',
    {
      teamId: z.string(),
      userId: z.string().describe('Object id of the user to add'),
      roles: z.array(z.enum(['owner', 'member'])).default(['member']),
    },
    async ({ teamId, userId, roles }) => {
      const member = await graph.post(`/teams/${teamId}/members`, {
        '@odata.type': '#microsoft.graph.aadUserConversationMember',
        roles,
        'user@odata.bind': `https://graph.microsoft.com/v1.0/users/${userId}`,
      });
      return { content: [{ type: 'text', text: JSON.stringify(member, null, 2) }] };
    }
  );
}
