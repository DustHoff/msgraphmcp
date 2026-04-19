import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { GraphClient } from '../graph/GraphClient';

export function registerTaskTools(server: McpServer, graph: GraphClient) {
  server.tool(
    'list_todo_lists',
    'List Microsoft To Do task lists for a user.',
    { userId: z.string().default('me') },
    async ({ userId }) => {
      const lists = await graph.getAll(`/users/${encodeURIComponent(userId)}/todo/lists`);
      return { content: [{ type: 'text', text: JSON.stringify(lists, null, 2) }] };
    }
  );

  server.tool(
    'create_todo_list',
    'Create a new To Do task list.',
    {
      userId: z.string().default('me'),
      displayName: z.string(),
    },
    async ({ userId, displayName }) => {
      const list = await graph.post(`/users/${encodeURIComponent(userId)}/todo/lists`, { displayName });
      return { content: [{ type: 'text', text: JSON.stringify(list, null, 2) }] };
    }
  );

  server.tool(
    'delete_todo_list',
    'Delete a To Do task list.',
    {
      userId: z.string().default('me'),
      listId: z.string(),
    },
    async ({ userId, listId }) => {
      await graph.delete(`/users/${encodeURIComponent(userId)}/todo/lists/${listId}`);
      return { content: [{ type: 'text', text: 'Task list deleted.' }] };
    }
  );

  server.tool(
    'list_tasks',
    'List tasks in a To Do task list.',
    {
      userId: z.string().default('me'),
      listId: z.string(),
      filter: z.string().optional().describe("OData filter, e.g. \"status eq 'notStarted'\""),
      top: z.number().int().min(1).max(999).default(50),
    },
    async ({ userId, listId, filter, top }) => {
      const params: Record<string, unknown> = { $top: top };
      if (filter) params['$filter'] = filter;
      const tasks = await graph.get(
        `/users/${encodeURIComponent(userId)}/todo/lists/${listId}/tasks`,
        params
      );
      return { content: [{ type: 'text', text: JSON.stringify(tasks, null, 2) }] };
    }
  );

  server.tool(
    'create_task',
    'Create a task in a To Do list.',
    {
      userId: z.string().default('me'),
      listId: z.string(),
      title: z.string(),
      body: z.string().optional().describe('Task body / notes'),
      dueDateTime: z.string().optional().describe('ISO 8601 due date-time, e.g. 2024-12-31T23:59:00'),
      importance: z.enum(['low', 'normal', 'high']).default('normal'),
      reminderDateTime: z.string().optional(),
    },
    async ({ userId, listId, title, body, dueDateTime, importance, reminderDateTime }) => {
      const taskBody: Record<string, unknown> = { title, importance };
      if (body) taskBody.body = { content: body, contentType: 'text' };
      if (dueDateTime) taskBody.dueDateTime = { dateTime: dueDateTime, timeZone: 'UTC' };
      if (reminderDateTime) {
        taskBody.isReminderOn = true;
        taskBody.reminderDateTime = { dateTime: reminderDateTime, timeZone: 'UTC' };
      }

      const task = await graph.post(
        `/users/${encodeURIComponent(userId)}/todo/lists/${listId}/tasks`,
        taskBody
      );
      return { content: [{ type: 'text', text: JSON.stringify(task, null, 2) }] };
    }
  );

  server.tool(
    'update_task',
    'Update a task.',
    {
      userId: z.string().default('me'),
      listId: z.string(),
      taskId: z.string(),
      title: z.string().optional(),
      body: z.string().optional(),
      dueDateTime: z.string().optional(),
      importance: z.enum(['low', 'normal', 'high']).optional(),
      status: z.enum(['notStarted', 'inProgress', 'completed', 'waitingOnOthers', 'deferred']).optional(),
    },
    async ({ userId, listId, taskId, title, body, dueDateTime, importance, status }) => {
      const patch: Record<string, unknown> = {};
      if (title) patch.title = title;
      if (body) patch.body = { content: body, contentType: 'text' };
      if (dueDateTime) patch.dueDateTime = { dateTime: dueDateTime, timeZone: 'UTC' };
      if (importance) patch.importance = importance;
      if (status) patch.status = status;

      const task = await graph.patch(
        `/users/${encodeURIComponent(userId)}/todo/lists/${listId}/tasks/${taskId}`,
        patch
      );
      return { content: [{ type: 'text', text: JSON.stringify(task, null, 2) }] };
    }
  );

  server.tool(
    'complete_task',
    'Mark a task as completed.',
    {
      userId: z.string().default('me'),
      listId: z.string(),
      taskId: z.string(),
    },
    async ({ userId, listId, taskId }) => {
      const task = await graph.patch(
        `/users/${encodeURIComponent(userId)}/todo/lists/${listId}/tasks/${taskId}`,
        { status: 'completed' }
      );
      return { content: [{ type: 'text', text: JSON.stringify(task, null, 2) }] };
    }
  );

  server.tool(
    'delete_task',
    'Delete a task.',
    {
      userId: z.string().default('me'),
      listId: z.string(),
      taskId: z.string(),
    },
    async ({ userId, listId, taskId }) => {
      await graph.delete(`/users/${encodeURIComponent(userId)}/todo/lists/${listId}/tasks/${taskId}`);
      return { content: [{ type: 'text', text: 'Task deleted.' }] };
    }
  );
}
