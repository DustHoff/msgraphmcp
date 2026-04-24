import { MockMcpServer } from '../helpers/MockMcpServer';
import { args, createMockGraphClient } from '../helpers/mockGraphClient';
import { registerTeamsTools } from '../../src/tools/teams';

describe('Teams Tools', () => {
  let server: MockMcpServer;
  let graph: ReturnType<typeof createMockGraphClient>;

  beforeEach(() => {
    server = new MockMcpServer();
    graph = createMockGraphClient();
    registerTeamsTools(server as never, graph as never);
  });

  afterEach(() => jest.clearAllMocks());

  const TOOLS = [
    'list_joined_teams', 'get_team', 'create_team',
    'list_channels', 'get_channel', 'create_channel', 'delete_channel',
    'list_channel_messages', 'send_channel_message', 'reply_to_channel_message',
    'list_team_members', 'add_team_member',
  ];

  it('registers all teams tools', () => {
    TOOLS.forEach(name => expect(server.isRegistered(name)).toBe(true));
  });

  describe('create_team', () => {
    it('includes odata.bind template reference', async () => {
      graph.post.mockResolvedValue({ id: 't1' });
      await server.call('create_team', { displayName: 'Engineering' });
      const [, body] = args(graph.post);
      expect(body['template@odata.bind']).toContain('teamsTemplates');
      expect(body.displayName).toBe('Engineering');
    });
  });

  describe('send_channel_message', () => {
    it('posts message with body structure', async () => {
      graph.post.mockResolvedValue({ id: 'm1' });
      await server.call('send_channel_message', {
        teamId: 't1', channelId: 'c1', content: 'Hello team!',
      });
      const [url, body] = args(graph.post);
      expect(url).toBe('/teams/t1/channels/c1/messages');
      expect(body.body.content).toBe('Hello team!');
    });
  });

  describe('reply_to_channel_message', () => {
    it('posts to replies endpoint', async () => {
      graph.post.mockResolvedValue({ id: 'r1' });
      await server.call('reply_to_channel_message', {
        teamId: 't1', channelId: 'c1', messageId: 'm1', content: 'Agreed!',
      });
      const [url] = args(graph.post);
      expect(url).toContain('m1/replies');
    });
  });

  describe('add_team_member', () => {
    it('sends correct odata type and user bind', async () => {
      graph.post.mockResolvedValue({ id: 'mem1' });
      await server.call('add_team_member', { teamId: 't1', userId: 'u1' });
      const [, body] = args(graph.post);
      expect(body['@odata.type']).toContain('aadUserConversationMember');
      expect(body['user@odata.bind']).toContain('u1');
    });
  });

  describe('delete_channel', () => {
    it('calls DELETE on the channel', async () => {
      graph.delete.mockResolvedValue(undefined);
      await server.call('delete_channel', { teamId: 't1', channelId: 'c1' });
      expect(graph.delete).toHaveBeenCalledWith('/teams/t1/channels/c1');
    });
  });

  describe('URL-encoding of opaque ids', () => {
    it('encodes teamId and channelId in path', async () => {
      graph.delete.mockResolvedValue(undefined);
      await server.call('delete_channel', { teamId: 't/1', channelId: 'c?1' });
      expect(graph.delete).toHaveBeenCalledWith('/teams/t%2F1/channels/c%3F1');
    });

    it('encodes userId in @odata.bind for add_team_member', async () => {
      graph.post.mockResolvedValue({ id: 'm1' });
      await server.call('add_team_member', { teamId: 't1', userId: 'u/1' });
      const [, body] = args(graph.post);
      expect(body['user@odata.bind']).toBe('https://graph.microsoft.com/v1.0/users/u%2F1');
    });
  });
});
