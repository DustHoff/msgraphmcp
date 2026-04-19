import { MockMcpServer } from '../helpers/MockMcpServer';
import { args, createMockGraphClient } from '../helpers/mockGraphClient';
import { registerMailTools } from '../../src/tools/mail';

describe('Mail Tools', () => {
  let server: MockMcpServer;
  let graph: ReturnType<typeof createMockGraphClient>;

  beforeEach(() => {
    server = new MockMcpServer();
    graph = createMockGraphClient();
    registerMailTools(server as never, graph as never);
  });

  afterEach(() => jest.clearAllMocks());

  const TOOLS = [
    'list_messages', 'get_message', 'send_mail', 'reply_to_message',
    'forward_message', 'delete_message', 'move_message',
    'list_mail_folders', 'create_mail_folder',
  ];

  it('registers all mail tools', () => {
    TOOLS.forEach(name => expect(server.isRegistered(name)).toBe(true));
  });

  describe('list_messages', () => {
    it('uses inbox folder by default', async () => {
      graph.get.mockResolvedValue({ value: [] });
      await server.call('list_messages', {});
      const [url, params] = args(graph.get);
      expect(url).toContain('inbox/messages');
      expect(params.$top).toBe(25);
    });

    it('supports custom folderId', async () => {
      graph.get.mockResolvedValue({ value: [] });
      await server.call('list_messages', { folderId: 'sentitems' });
      const [url] = args(graph.get);
      expect(url).toContain('sentitems');
    });
  });

  describe('send_mail', () => {
    it('posts to sendMail endpoint with correct structure', async () => {
      graph.post.mockResolvedValue(undefined);
      await server.call('send_mail', {
        subject: 'Hello',
        body: 'World',
        toRecipients: [{ address: 'bob@contoso.com' }],
      });
      const [url, payload] = args(graph.post);
      expect(url).toBe('/users/me/sendMail');
      expect(payload.message.subject).toBe('Hello');
      expect(payload.message.toRecipients[0].emailAddress.address).toBe('bob@contoso.com');
    });

    it('includes CC and BCC when provided', async () => {
      graph.post.mockResolvedValue(undefined);
      await server.call('send_mail', {
        subject: 'Test',
        body: 'Body',
        toRecipients: [{ address: 'to@x.com' }],
        ccRecipients: [{ address: 'cc@x.com' }],
        bccRecipients: [{ address: 'bcc@x.com' }],
      });
      const [, payload] = args(graph.post);
      expect(payload.message.ccRecipients).toHaveLength(1);
      expect(payload.message.bccRecipients).toHaveLength(1);
    });

    it('rejects empty toRecipients array', async () => {
      await expect(server.call('send_mail', {
        subject: 'Test', body: 'Body', toRecipients: [],
      })).rejects.toThrow();
    });
  });

  describe('reply_to_message', () => {
    it('posts to reply endpoint', async () => {
      graph.post.mockResolvedValue(undefined);
      await server.call('reply_to_message', { messageId: 'msg1', comment: 'Thanks!' });
      expect(graph.post).toHaveBeenCalledWith(
        '/users/me/messages/msg1/reply',
        { comment: 'Thanks!' },
      );
    });
  });

  describe('forward_message', () => {
    it('posts to forward endpoint with recipients', async () => {
      graph.post.mockResolvedValue(undefined);
      await server.call('forward_message', {
        messageId: 'msg1',
        toRecipients: [{ address: 'fwd@x.com' }],
      });
      const [url, body] = args(graph.post);
      expect(url).toContain('forward');
      expect(body.toRecipients[0].emailAddress.address).toBe('fwd@x.com');
    });
  });

  describe('delete_message', () => {
    it('calls DELETE on the message', async () => {
      graph.delete.mockResolvedValue(undefined);
      await server.call('delete_message', { messageId: 'msg1' });
      expect(graph.delete).toHaveBeenCalledWith('/users/me/messages/msg1');
    });
  });

  describe('move_message', () => {
    it('posts move with destinationId', async () => {
      graph.post.mockResolvedValue({ id: 'msg1', parentFolderId: 'archive' });
      await server.call('move_message', { messageId: 'msg1', destinationFolderId: 'archive' });
      const [, body] = args(graph.post);
      expect(body.destinationId).toBe('archive');
    });
  });

  describe('create_mail_folder', () => {
    it('creates top-level folder', async () => {
      graph.post.mockResolvedValue({ id: 'f1', displayName: 'Projects' });
      await server.call('create_mail_folder', { displayName: 'Projects' });
      const [url] = args(graph.post);
      expect(url).toBe('/users/me/mailFolders');
    });

    it('creates child folder when parentFolderId given', async () => {
      graph.post.mockResolvedValue({ id: 'f2' });
      await server.call('create_mail_folder', { displayName: 'Sub', parentFolderId: 'inbox' });
      const [url] = args(graph.post);
      expect(url).toContain('inbox/childFolders');
    });
  });
});
