import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { GraphClient } from '../graph/GraphClient';
import { userPath } from './shared';

export function registerCalendarTools(server: McpServer, graph: GraphClient) {
  server.tool(
    'list_calendars',
    'List calendars for a user.',
    { userId: z.string().default('me') },
    async ({ userId }) => {
      const calendars = await graph.getAll(`${userPath(userId)}/calendars`);
      return { content: [{ type: 'text', text: JSON.stringify(calendars, null, 2) }] };
    }
  );

  server.tool(
    'create_calendar',
    'Create a new calendar for a user.',
    {
      userId: z.string().default('me'),
      name: z.string().describe('Calendar name'),
      color: z.enum(['auto', 'lightBlue', 'lightGreen', 'lightOrange', 'lightGray', 'lightYellow', 'lightTeal', 'lightPink', 'lightBrown', 'lightRed', 'maxColor']).default('auto'),
    },
    async ({ userId, name, color }) => {
      const calendar = await graph.post(`${userPath(userId)}/calendars`, { name, color });
      return { content: [{ type: 'text', text: JSON.stringify(calendar, null, 2) }] };
    }
  );

  server.tool(
    'list_events',
    'List calendar events. Defaults to the primary calendar.',
    {
      userId: z.string().default('me'),
      calendarId: z.string().optional().describe('Calendar id (omit for primary)'),
      filter: z.string().optional().describe("OData filter, e.g. \"start/dateTime ge '2024-01-01T00:00:00'\""),
      select: z.string().optional(),
      top: z.number().int().min(1).max(999).default(25),
      startDateTime: z.string().optional().describe('ISO 8601 start for calendar view (requires endDateTime)'),
      endDateTime: z.string().optional().describe('ISO 8601 end for calendar view'),
    },
    async ({ userId, calendarId, filter, select, top, startDateTime, endDateTime }) => {
      const base = calendarId
        ? `${userPath(userId)}/calendars/${encodeURIComponent(calendarId)}`
        : userPath(userId);

      let url: string;
      const params: Record<string, unknown> = { $top: top };

      if (startDateTime && endDateTime) {
        url = `${base}/calendarView`;
        params.startDateTime = startDateTime;
        params.endDateTime = endDateTime;
      } else {
        url = `${base}/events`;
      }

      if (filter) params['$filter'] = filter;
      if (select) params['$select'] = select;

      const events = await graph.get(url, params);
      return { content: [{ type: 'text', text: JSON.stringify(events, null, 2) }] };
    }
  );

  server.tool(
    'get_event',
    'Get a specific calendar event.',
    {
      userId: z.string().default('me'),
      eventId: z.string(),
    },
    async ({ userId, eventId }) => {
      const event = await graph.get(`${userPath(userId)}/events/${encodeURIComponent(eventId)}`);
      return { content: [{ type: 'text', text: JSON.stringify(event, null, 2) }] };
    }
  );

  server.tool(
    'create_event',
    'Create a calendar event.',
    {
      userId: z.string().default('me'),
      calendarId: z.string().optional(),
      subject: z.string(),
      body: z.string().optional().describe('Event description / body'),
      bodyContentType: z.enum(['Text', 'HTML']).default('Text'),
      startDateTime: z.string().describe('ISO 8601 date-time, e.g. 2024-06-01T10:00:00'),
      startTimeZone: z.string().default('UTC'),
      endDateTime: z.string().describe('ISO 8601 date-time'),
      endTimeZone: z.string().default('UTC'),
      location: z.string().optional(),
      attendees: z.array(z.object({
        address: z.string().email(),
        name: z.string().optional(),
        type: z.enum(['required', 'optional', 'resource']).default('required'),
      })).optional(),
      isOnlineMeeting: z.boolean().default(false),
      isAllDay: z.boolean().default(false),
    },
    async ({ userId, calendarId, subject, body, bodyContentType, startDateTime, startTimeZone, endDateTime, endTimeZone, location, attendees, isOnlineMeeting, isAllDay }) => {
      const eventBody: Record<string, unknown> = {
        subject,
        start: { dateTime: startDateTime, timeZone: startTimeZone },
        end: { dateTime: endDateTime, timeZone: endTimeZone },
        isOnlineMeeting,
        isAllDay,
      };
      if (body) eventBody.body = { contentType: bodyContentType, content: body };
      if (location) eventBody.location = { displayName: location };
      if (attendees?.length) {
        eventBody.attendees = attendees.map((a) => ({
          emailAddress: { address: a.address, name: a.name },
          type: a.type,
        }));
      }

      const base = calendarId
        ? `${userPath(userId)}/calendars/${encodeURIComponent(calendarId)}`
        : userPath(userId);
      const event = await graph.post(`${base}/events`, eventBody);
      return { content: [{ type: 'text', text: JSON.stringify(event, null, 2) }] };
    }
  );

  server.tool(
    'update_event',
    'Update an existing calendar event.',
    {
      userId: z.string().default('me'),
      eventId: z.string(),
      subject: z.string().optional(),
      body: z.string().optional(),
      bodyContentType: z.enum(['Text', 'HTML']).default('Text'),
      startDateTime: z.string().optional(),
      startTimeZone: z.string().optional(),
      endDateTime: z.string().optional(),
      endTimeZone: z.string().optional(),
      location: z.string().optional(),
      isOnlineMeeting: z.boolean().optional(),
    },
    async ({ userId, eventId, subject, body, bodyContentType, startDateTime, startTimeZone, endDateTime, endTimeZone, location, isOnlineMeeting }) => {
      const patch: Record<string, unknown> = {};
      if (subject) patch.subject = subject;
      if (body) patch.body = { contentType: bodyContentType, content: body };
      if (startDateTime) patch.start = { dateTime: startDateTime, timeZone: startTimeZone ?? 'UTC' };
      if (endDateTime) patch.end = { dateTime: endDateTime, timeZone: endTimeZone ?? 'UTC' };
      if (location) patch.location = { displayName: location };
      if (isOnlineMeeting !== undefined) patch.isOnlineMeeting = isOnlineMeeting;

      const event = await graph.patch(`${userPath(userId)}/events/${encodeURIComponent(eventId)}`, patch);
      return { content: [{ type: 'text', text: JSON.stringify(event, null, 2) }] };
    }
  );

  server.tool(
    'delete_event',
    'Delete a calendar event.',
    { userId: z.string().default('me'), eventId: z.string() },
    async ({ userId, eventId }) => {
      await graph.delete(`${userPath(userId)}/events/${encodeURIComponent(eventId)}`);
      return { content: [{ type: 'text', text: 'Event deleted.' }] };
    }
  );
}
