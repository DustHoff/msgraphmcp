import { MockMcpServer } from '../helpers/MockMcpServer';
import { args, createMockGraphClient } from '../helpers/mockGraphClient';
import { registerContactTools } from '../../src/tools/contacts';

describe('Contact Tools', () => {
  let server: MockMcpServer;
  let graph: ReturnType<typeof createMockGraphClient>;

  beforeEach(() => {
    server = new MockMcpServer();
    graph = createMockGraphClient();
    registerContactTools(server as never, graph as never);
  });

  afterEach(() => jest.clearAllMocks());

  it('registers all contact tools', () => {
    ['list_contacts', 'get_contact', 'create_contact', 'update_contact', 'delete_contact']
      .forEach(name => expect(server.isRegistered(name)).toBe(true));
  });

  describe('create_contact', () => {
    it('posts to contacts endpoint with provided fields', async () => {
      graph.post.mockResolvedValue({ id: 'c1' });
      await server.call('create_contact', {
        givenName: 'Alice',
        surname: 'Smith',
        emailAddresses: [{ address: 'alice@x.com' }],
      });
      const [url, body] = args(graph.post);
      expect(url).toBe('/users/me/contacts');
      expect(body.givenName).toBe('Alice');
      expect(body.emailAddresses[0].address).toBe('alice@x.com');
    });

    it('omits undefined fields from body', async () => {
      graph.post.mockResolvedValue({ id: 'c2' });
      await server.call('create_contact', { givenName: 'Bob' });
      const [, body] = args(graph.post);
      expect(body.surname).toBeUndefined();
      expect(body.mobilePhone).toBeUndefined();
    });
  });

  describe('update_contact', () => {
    it('patches only provided fields', async () => {
      graph.patch.mockResolvedValue({ id: 'c1' });
      await server.call('update_contact', { contactId: 'c1', jobTitle: 'Dev' });
      const [url, body] = args(graph.patch);
      expect(url).toContain('c1');
      expect(Object.keys(body)).toEqual(['jobTitle']);
    });
  });

  describe('delete_contact', () => {
    it('calls DELETE with contact path', async () => {
      graph.delete.mockResolvedValue(undefined);
      await server.call('delete_contact', { contactId: 'c1' });
      expect(graph.delete).toHaveBeenCalledWith('/users/me/contacts/c1');
    });
  });
});
