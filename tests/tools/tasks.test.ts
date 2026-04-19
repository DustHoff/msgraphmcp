import { MockMcpServer } from '../helpers/MockMcpServer';
import { args, createMockGraphClient } from '../helpers/mockGraphClient';
import { registerTaskTools } from '../../src/tools/tasks';

describe('Task / To Do Tools', () => {
  let server: MockMcpServer;
  let graph: ReturnType<typeof createMockGraphClient>;

  beforeEach(() => {
    server = new MockMcpServer();
    graph = createMockGraphClient();
    registerTaskTools(server as never, graph as never);
  });

  afterEach(() => jest.clearAllMocks());

  const TOOLS = [
    'list_todo_lists', 'create_todo_list', 'delete_todo_list',
    'list_tasks', 'create_task', 'update_task', 'complete_task', 'delete_task',
  ];

  it('registers all task tools', () => {
    TOOLS.forEach(name => expect(server.isRegistered(name)).toBe(true));
  });

  describe('create_task', () => {
    it('creates task with title and default importance', async () => {
      graph.post.mockResolvedValue({ id: 'task1' });
      await server.call('create_task', { listId: 'list1', title: 'Buy milk' });
      const [url, body] = args(graph.post);
      expect(url).toContain('list1/tasks');
      expect(body.title).toBe('Buy milk');
      expect(body.importance).toBe('normal');
    });

    it('sets dueDateTime as structured object', async () => {
      graph.post.mockResolvedValue({ id: 'task1' });
      await server.call('create_task', {
        listId: 'list1',
        title: 'Deadline task',
        dueDateTime: '2024-12-31T23:59:00',
      });
      const [, body] = args(graph.post);
      expect(body.dueDateTime.dateTime).toBe('2024-12-31T23:59:00');
      expect(body.dueDateTime.timeZone).toBe('UTC');
    });

    it('enables reminder when reminderDateTime is provided', async () => {
      graph.post.mockResolvedValue({ id: 'task1' });
      await server.call('create_task', {
        listId: 'list1',
        title: 'Remind me',
        reminderDateTime: '2024-06-01T09:00:00',
      });
      const [, body] = args(graph.post);
      expect(body.isReminderOn).toBe(true);
      expect(body.reminderDateTime.dateTime).toBe('2024-06-01T09:00:00');
    });
  });

  describe('complete_task', () => {
    it('patches status to completed', async () => {
      graph.patch.mockResolvedValue({ id: 'task1', status: 'completed' });
      await server.call('complete_task', { listId: 'list1', taskId: 'task1' });
      const [, body] = args(graph.patch);
      expect(body.status).toBe('completed');
    });
  });

  describe('update_task', () => {
    it('sends status update', async () => {
      graph.patch.mockResolvedValue({ id: 't1' });
      await server.call('update_task', { listId: 'l1', taskId: 't1', status: 'inProgress' });
      const [, body] = args(graph.patch);
      expect(body.status).toBe('inProgress');
    });
  });
});
