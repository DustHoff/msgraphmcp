import { MockMcpServer } from '../helpers/MockMcpServer';
import { args, createMockGraphClient } from '../helpers/mockGraphClient';
import { registerGroupTools } from '../../src/tools/groups';

describe('Group Tools', () => {
  let server: MockMcpServer;
  let graph: ReturnType<typeof createMockGraphClient>;

  beforeEach(() => {
    server = new MockMcpServer();
    graph = createMockGraphClient();
    registerGroupTools(server as never, graph as never);
  });

  afterEach(() => jest.clearAllMocks());

  const TOOLS = [
    'list_groups', 'get_group', 'create_group', 'update_group', 'delete_group',
    'list_group_members', 'add_group_member', 'remove_group_member',
    'list_group_owners', 'add_group_owner',
  ];

  it('registers all group tools', () => {
    TOOLS.forEach(name => expect(server.isRegistered(name)).toBe(true));
  });

  describe('create_group', () => {
    it('creates M365 group with Unified groupType', async () => {
      graph.post.mockResolvedValue({ id: 'g1' });
      await server.call('create_group', { displayName: 'Dev Team', mailNickname: 'devteam' });
      const [, body] = args(graph.post);
      expect(body.groupTypes).toContain('Unified');
      expect(body.mailEnabled).toBe(true);
    });

    it('creates security group without Unified type', async () => {
      graph.post.mockResolvedValue({ id: 'g2' });
      await server.call('create_group', {
        displayName: 'SecGroup',
        mailNickname: 'sec',
        groupType: 'Security',
      });
      const [, body] = args(graph.post);
      expect(body.groupTypes).toEqual([]);
      expect(body.securityEnabled).toBe(true);
      expect(body.mailEnabled).toBe(false);
    });
  });

  describe('add_group_member', () => {
    it('posts $ref with correct odata.id', async () => {
      graph.post.mockResolvedValue(undefined);
      await server.call('add_group_member', { groupId: 'g1', memberId: 'u1' });
      const [url, body] = args(graph.post);
      expect(url).toBe('/groups/g1/members/$ref');
      expect(body['@odata.id']).toContain('u1');
    });
  });

  describe('remove_group_member', () => {
    it('calls DELETE on member ref', async () => {
      graph.delete.mockResolvedValue(undefined);
      await server.call('remove_group_member', { groupId: 'g1', memberId: 'u1' });
      expect(graph.delete).toHaveBeenCalledWith('/groups/g1/members/u1/$ref');
    });
  });

  describe('update_group', () => {
    it('sends only provided fields', async () => {
      graph.patch.mockResolvedValue(undefined);
      await server.call('update_group', { groupId: 'g1', description: 'New desc' });
      const [, body] = args(graph.patch);
      expect(body).toEqual({ description: 'New desc' });
      expect(body.displayName).toBeUndefined();
    });
  });

  describe('URL-encoding of opaque ids', () => {
    it('encodes groupId before embedding it in the path', async () => {
      graph.get.mockResolvedValue({ id: 'g/1' });
      await server.call('get_group', { groupId: 'g/1' });
      const [url] = args(graph.get);
      expect(url).toBe('/groups/g%2F1');
    });

    it('encodes both groupId and memberId in add_group_member', async () => {
      graph.post.mockResolvedValue(undefined);
      await server.call('add_group_member', { groupId: 'g/1', memberId: 'u?2' });
      const [url, body] = args(graph.post);
      expect(url).toBe('/groups/g%2F1/members/$ref');
      expect(body['@odata.id']).toBe('https://graph.microsoft.com/v1.0/directoryObjects/u%3F2');
    });

    it('encodes both groupId and memberId in remove_group_member', async () => {
      graph.delete.mockResolvedValue(undefined);
      await server.call('remove_group_member', { groupId: 'g/1', memberId: 'u?2' });
      expect(graph.delete).toHaveBeenCalledWith('/groups/g%2F1/members/u%3F2/$ref');
    });
  });
});
