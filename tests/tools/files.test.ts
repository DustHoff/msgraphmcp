import { MockMcpServer } from '../helpers/MockMcpServer';
import { args, createMockGraphClient } from '../helpers/mockGraphClient';
import { registerFileTools } from '../../src/tools/files';

describe('File / OneDrive Tools', () => {
  let server: MockMcpServer;
  let graph: ReturnType<typeof createMockGraphClient>;

  beforeEach(() => {
    server = new MockMcpServer();
    graph = createMockGraphClient();
    registerFileTools(server as never, graph as never);
  });

  afterEach(() => jest.clearAllMocks());

  const TOOLS = [
    'list_drive_items', 'get_drive_item', 'create_drive_folder',
    'upload_drive_file', 'delete_drive_item', 'copy_drive_item',
    'search_drive', 'list_shared_with_me',
  ];

  it('registers all file tools', () => {
    TOOLS.forEach(name => expect(server.isRegistered(name)).toBe(true));
  });

  describe('list_drive_items', () => {
    it('uses root/children for root path', async () => {
      graph.get.mockResolvedValue({ value: [] });
      await server.call('list_drive_items', {});
      const [url] = args(graph.get);
      expect(url).toContain('root/children');
    });

    it('uses root:<path>:/children for sub-folders', async () => {
      graph.get.mockResolvedValue({ value: [] });
      await server.call('list_drive_items', { itemPath: '/Documents' });
      const [url] = args(graph.get);
      expect(url).toContain('root:/Documents:/children');
    });
  });

  describe('create_drive_folder', () => {
    it('includes folder: {} in body', async () => {
      graph.post.mockResolvedValue({ id: 'f1' });
      await server.call('create_drive_folder', { folderName: 'Reports' });
      const [, body] = args(graph.post);
      expect(body.folder).toEqual({});
      expect(body.name).toBe('Reports');
    });
  });

  describe('upload_drive_file', () => {
    it('puts content to root:<path>:/content', async () => {
      graph.put.mockResolvedValue({ id: 'fi1' });
      await server.call('upload_drive_file', {
        filePath: '/Documents/report.txt',
        content: 'Hello World',
      });
      const [url] = args(graph.put);
      expect(url).toContain('/Documents/report.txt:/content');
    });
  });

  describe('copy_drive_item', () => {
    it('sends parentReference with destination id', async () => {
      graph.post.mockResolvedValue({});
      await server.call('copy_drive_item', {
        itemId: 'item1',
        destinationParentId: 'dest1',
      });
      const [, body] = args(graph.post);
      expect(body.parentReference.id).toBe('dest1');
    });

    it('includes new name when provided', async () => {
      graph.post.mockResolvedValue({});
      await server.call('copy_drive_item', {
        itemId: 'item1',
        destinationParentId: 'dest1',
        newName: 'copy.docx',
      });
      const [, body] = args(graph.post);
      expect(body.name).toBe('copy.docx');
    });
  });

  describe('search_drive', () => {
    it('encodes search query in URL', async () => {
      graph.get.mockResolvedValue({ value: [] });
      await server.call('search_drive', { query: 'budget report' });
      const [url] = args(graph.get);
      expect(url).toContain('budget%20report');
    });
  });
});
