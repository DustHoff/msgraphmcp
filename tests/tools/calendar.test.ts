import { MockMcpServer } from '../helpers/MockMcpServer';
import { args, createMockGraphClient } from '../helpers/mockGraphClient';
import { registerCalendarTools } from '../../src/tools/calendar';

describe('Calendar Tools', () => {
  let server: MockMcpServer;
  let graph: ReturnType<typeof createMockGraphClient>;

  beforeEach(() => {
    server = new MockMcpServer();
    graph = createMockGraphClient();
    registerCalendarTools(server as never, graph as never);
  });

  afterEach(() => jest.clearAllMocks());

  const TOOLS = ['list_calendars', 'create_calendar', 'list_events', 'get_event', 'create_event', 'update_event', 'delete_event'];

  it('registers all calendar tools', () => {
    TOOLS.forEach(name => expect(server.isRegistered(name)).toBe(true));
  });

  describe('list_events', () => {
    it('uses /events by default', async () => {
      graph.get.mockResolvedValue({ value: [] });
      await server.call('list_events', {});
      const [url] = args(graph.get);
      expect(url).toContain('/events');
    });

    it('uses calendarView when start/end provided', async () => {
      graph.get.mockResolvedValue({ value: [] });
      await server.call('list_events', {
        startDateTime: '2024-01-01T00:00:00',
        endDateTime: '2024-01-31T23:59:59',
      });
      const [url] = args(graph.get);
      expect(url).toContain('calendarView');
    });

    it('uses specific calendar when calendarId provided', async () => {
      graph.get.mockResolvedValue({ value: [] });
      await server.call('list_events', { calendarId: 'cal1' });
      const [url] = args(graph.get);
      expect(url).toContain('/calendars/cal1/events');
    });
  });

  describe('create_event', () => {
    const baseEvent = {
      subject: 'Team Meeting',
      startDateTime: '2024-06-01T10:00:00',
      endDateTime: '2024-06-01T11:00:00',
    };

    it('creates event in primary calendar', async () => {
      graph.post.mockResolvedValue({ id: 'ev1' });
      await server.call('create_event', baseEvent);
      const [url, body] = args(graph.post);
      expect(url).toContain('/users/me/events');
      expect(body.subject).toBe('Team Meeting');
      expect(body.start.dateTime).toBe('2024-06-01T10:00:00');
    });

    it('attaches attendees when provided', async () => {
      graph.post.mockResolvedValue({ id: 'ev1' });
      await server.call('create_event', {
        ...baseEvent,
        attendees: [{ address: 'bob@x.com', name: 'Bob', type: 'required' }],
      });
      const [, body] = args(graph.post);
      expect(body.attendees).toHaveLength(1);
      expect(body.attendees[0].emailAddress.address).toBe('bob@x.com');
    });

    it('creates event in specific calendar', async () => {
      graph.post.mockResolvedValue({ id: 'ev1' });
      await server.call('create_event', { ...baseEvent, calendarId: 'cal1' });
      const [url] = args(graph.post);
      expect(url).toContain('/calendars/cal1/events');
    });
  });

  describe('update_event', () => {
    it('sends only provided fields', async () => {
      graph.patch.mockResolvedValue({ id: 'ev1' });
      await server.call('update_event', { eventId: 'ev1', subject: 'New Title' });
      const [, body] = args(graph.patch);
      expect(body.subject).toBe('New Title');
      expect(body.start).toBeUndefined();
    });
  });

  describe('delete_event', () => {
    it('calls DELETE on the event', async () => {
      graph.delete.mockResolvedValue(undefined);
      await server.call('delete_event', { eventId: 'ev1' });
      expect(graph.delete).toHaveBeenCalledWith('/users/me/events/ev1');
    });
  });
});
