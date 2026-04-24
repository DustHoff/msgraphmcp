import { MockMcpServer } from '../helpers/MockMcpServer';
import { args, createMockGraphClient } from '../helpers/mockGraphClient';
import { registerUserTools } from '../../src/tools/users';

describe('User Tools', () => {
  let server: MockMcpServer;
  let graph: ReturnType<typeof createMockGraphClient>;

  beforeEach(() => {
    server = new MockMcpServer();
    graph = createMockGraphClient();
    registerUserTools(server as never, graph as never);
  });

  afterEach(() => jest.clearAllMocks());

  const TOOLS = [
    'list_users', 'get_user', 'create_user', 'update_user',
    'delete_user', 'get_user_member_of', 'reset_user_password',
  ];

  it('registers all user tools', () => {
    TOOLS.forEach(name => expect(server.isRegistered(name)).toBe(true));
  });

  describe('list_users', () => {
    it('calls getAll on /users', async () => {
      graph.getAll.mockResolvedValue([{ id: '1', displayName: 'Alice' }]);
      const result = await server.call('list_users', {});
      // getAll now accepts (url, params, config?); config is undefined when
      // the query does not require ConsistencyLevel: eventual.
      expect(graph.getAll).toHaveBeenCalledWith('/users', expect.any(Object), undefined);
      expect(result.content[0].text).toContain('Alice');
    });

    it('passes filter and select params', async () => {
      graph.getAll.mockResolvedValue([]);
      await server.call('list_users', { filter: "displayName eq 'Bob'", select: 'id,displayName' });
      const [, params] = args(graph.getAll);
      expect(params.$filter).toBe("displayName eq 'Bob'");
      expect(params.$select).toBe('id,displayName');
    });
  });

  describe('get_user', () => {
    it('calls get with userId in path', async () => {
      graph.get.mockResolvedValue({ id: 'u1', displayName: 'Alice' });
      const result = await server.call('get_user', { userId: 'u1' });
      expect(graph.get).toHaveBeenCalledWith('/users/u1', undefined);
      expect(result.content[0].text).toContain('u1');
    });

    it('encodes special characters in userId', async () => {
      graph.get.mockResolvedValue({ id: 'u1' });
      await server.call('get_user', { userId: 'alice@contoso.com' });
      const [url] = args(graph.get);
      expect(url).toContain('alice%40contoso.com');
    });
  });

  describe('create_user', () => {
    it('sends correct body to POST /users', async () => {
      graph.post.mockResolvedValue({ id: 'new-id' });
      await server.call('create_user', {
        displayName: 'New User',
        userPrincipalName: 'new@contoso.com',
        mailNickname: 'new',
        password: 'P@ssw0rd!',
      });
      const [url, body] = args(graph.post);
      expect(url).toBe('/users');
      expect(body.displayName).toBe('New User');
      expect(body.passwordProfile.password).toBe('P@ssw0rd!');
    });
  });

  describe('update_user', () => {
    it('sends PATCH with changed fields only', async () => {
      graph.patch.mockResolvedValue(undefined);
      await server.call('update_user', { userId: 'u1', jobTitle: 'Engineer' });
      const [url, body] = args(graph.patch);
      expect(url).toBe('/users/u1');
      expect(body).toEqual({ jobTitle: 'Engineer' });
    });
  });

  describe('delete_user', () => {
    it('calls DELETE with correct path', async () => {
      graph.delete.mockResolvedValue(undefined);
      const result = await server.call('delete_user', { userId: 'u1' });
      expect(graph.delete).toHaveBeenCalledWith('/users/u1');
      expect(result.content[0].text).toContain('deleted');
    });
  });

  describe('get_user_member_of', () => {
    it('returns group membership', async () => {
      graph.getAll.mockResolvedValue([{ id: 'g1', displayName: 'Admins' }]);
      const result = await server.call('get_user_member_of', { userId: 'u1' });
      expect(graph.getAll).toHaveBeenCalledWith('/users/u1/memberOf');
      expect(result.content[0].text).toContain('Admins');
    });
  });

  describe('reset_user_password', () => {
    it('sends passwordProfile PATCH', async () => {
      graph.patch.mockResolvedValue(undefined);
      await server.call('reset_user_password', { userId: 'u1', newPassword: 'New@Pass1' });
      const [, body] = args(graph.patch);
      expect(body.passwordProfile.password).toBe('New@Pass1');
    });
  });
});
