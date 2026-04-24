import { MockMcpServer } from '../helpers/MockMcpServer';
import { args, createMockGraphClient } from '../helpers/mockGraphClient';
import { registerSiteTools } from '../../src/tools/sites';

describe('SharePoint Site Tools', () => {
  let server: MockMcpServer;
  let graph: ReturnType<typeof createMockGraphClient>;

  beforeEach(() => {
    server = new MockMcpServer();
    graph = createMockGraphClient();
    registerSiteTools(server as never, graph as never);
  });

  afterEach(() => jest.clearAllMocks());

  const TOOLS = [
    'list_sites', 'get_site', 'search_sites', 'list_site_lists', 'get_site_list',
    'list_site_list_items', 'get_site_list_item', 'create_site_list_item',
    'update_site_list_item', 'delete_site_list_item',
  ];

  it('registers all site tools', () => {
    TOOLS.forEach(name => expect(server.isRegistered(name)).toBe(true));
  });

  describe('get_site', () => {
    it('uses root when no args provided', async () => {
      graph.get.mockResolvedValue({ id: 'root' });
      await server.call('get_site', {});
      expect(graph.get).toHaveBeenCalledWith('/sites/root');
    });

    it('uses siteId when provided', async () => {
      graph.get.mockResolvedValue({ id: 's1' });
      await server.call('get_site', { siteId: 's1' });
      expect(graph.get).toHaveBeenCalledWith('/sites/s1');
    });

    it('builds hostname:path URL', async () => {
      graph.get.mockResolvedValue({ id: 'sp1' });
      await server.call('get_site', { hostname: 'contoso.sharepoint.com', sitePath: '/sites/HR' });
      expect(graph.get).toHaveBeenCalledWith('/sites/contoso.sharepoint.com:/sites/HR');
    });
  });

  describe('create_site_list_item', () => {
    it('posts fields object to list items', async () => {
      graph.post.mockResolvedValue({ id: 'item1' });
      await server.call('create_site_list_item', {
        siteId: 's1',
        listId: 'l1',
        fields: { Title: 'New Item', Status: 'Active' },
      });
      const [url, body] = args(graph.post);
      expect(url).toBe('/sites/s1/lists/l1/items');
      expect(body.fields.Title).toBe('New Item');
    });
  });

  describe('update_site_list_item', () => {
    it('patches the fields subresource', async () => {
      graph.patch.mockResolvedValue({ Title: 'Updated' });
      await server.call('update_site_list_item', {
        siteId: 's1', listId: 'l1', itemId: 'i1',
        fields: { Title: 'Updated' },
      });
      const [url] = args(graph.patch);
      expect(url).toBe('/sites/s1/lists/l1/items/i1/fields');
    });
  });

  describe('delete_site_list_item', () => {
    it('calls DELETE with full item path', async () => {
      graph.delete.mockResolvedValue(undefined);
      await server.call('delete_site_list_item', { siteId: 's1', listId: 'l1', itemId: 'i1' });
      expect(graph.delete).toHaveBeenCalledWith('/sites/s1/lists/l1/items/i1');
    });
  });

  describe('URL-encoding of opaque ids', () => {
    it('preserves the commas in composite siteIds (hostname,guid,guid)', async () => {
      graph.get.mockResolvedValue({ id: 'x' });
      await server.call('get_site', { siteId: 'contoso.sharepoint.com,abc,def' });
      expect(graph.get).toHaveBeenCalledWith('/sites/contoso.sharepoint.com,abc,def');
    });

    it('encodes hostname and path when building hostname:path URLs', async () => {
      graph.get.mockResolvedValue({ id: 'x' });
      await server.call('get_site', {
        hostname: 'contoso.sharepoint.com',
        sitePath: '/sites/HR Team',
      });
      expect(graph.get).toHaveBeenCalledWith('/sites/contoso.sharepoint.com:/sites/HR%20Team');
    });

    it('encodes listId and itemId when building list item URLs', async () => {
      graph.get.mockResolvedValue({ id: 'x' });
      await server.call('get_site_list_item', {
        siteId: 's1', listId: 'l/1', itemId: 'i?1',
      });
      const [url, params] = args(graph.get);
      expect(url).toBe('/sites/s1/lists/l%2F1/items/i%3F1');
      expect(params).toEqual({ $expand: 'fields' });
    });
  });
});
