import { describe, it, expect } from 'vitest';
import {
  sanitizeTask,
  sanitizeTasks,
  sanitizeTaskList,
  sanitizeTaskLists,
  sanitizeError,
} from '../../src/security/sanitizers.js';
import { ToDoTask, ToDoTaskList } from '../../src/graph/todo-service.js';
import { ValidationError, RateLimitError } from '../../src/security/errors.js';

describe('sanitizers', () => {
  describe('sanitizeTask', () => {
    it('should sanitize a complete task', () => {
      const task: ToDoTask = {
        id: 'task-123',
        title: 'Test Task',
        status: 'notStarted',
        importance: 'high',
        createdDateTime: '2025-01-01T00:00:00Z',
        lastModifiedDateTime: '2025-01-02T00:00:00Z',
        dueDateTime: {
          dateTime: '2025-01-10T00:00:00Z',
          timeZone: 'UTC',
        },
        reminderDateTime: {
          dateTime: '2025-01-09T00:00:00Z',
          timeZone: 'UTC',
        },
        isReminderOn: true,
        body: {
          content: 'Task description',
          contentType: 'text',
        },
        categories: ['Work', 'Important'],
      };

      const sanitized = sanitizeTask(task);

      expect(sanitized).toEqual({
        id: 'task-123',
        title: 'Test Task',
        status: 'notStarted',
        importance: 'high',
        createdDateTime: '2025-01-01T00:00:00Z',
        lastModifiedDateTime: '2025-01-02T00:00:00Z',
        dueDateTime: '2025-01-10T00:00:00Z',
        isReminderOn: true,
        body: 'Task description',
        categories: ['Work', 'Important'],
      });
    });

    it('should handle task without optional fields', () => {
      const task: ToDoTask = {
        id: 'task-456',
        title: 'Minimal Task',
        status: 'completed',
        importance: 'normal',
        createdDateTime: '2025-01-01T00:00:00Z',
        lastModifiedDateTime: '2025-01-02T00:00:00Z',
        isReminderOn: false,
      };

      const sanitized = sanitizeTask(task);

      expect(sanitized).toEqual({
        id: 'task-456',
        title: 'Minimal Task',
        status: 'completed',
        importance: 'normal',
        createdDateTime: '2025-01-01T00:00:00Z',
        lastModifiedDateTime: '2025-01-02T00:00:00Z',
        dueDateTime: undefined,
        isReminderOn: false,
        body: undefined,
        categories: undefined,
      });
    });

    it('should extract dateTime from dueDateTime object', () => {
      const task: ToDoTask = {
        id: 'task-789',
        title: 'Task with Due Date',
        status: 'inProgress',
        importance: 'low',
        createdDateTime: '2025-01-01T00:00:00Z',
        lastModifiedDateTime: '2025-01-02T00:00:00Z',
        dueDateTime: {
          dateTime: '2025-01-15T12:00:00Z',
          timeZone: 'America/New_York',
        },
        isReminderOn: false,
      };

      const sanitized = sanitizeTask(task);

      expect(sanitized.dueDateTime).toBe('2025-01-15T12:00:00Z');
      // Timezone should not be included in sanitized output
      expect(sanitized).not.toHaveProperty('timeZone');
    });
  });

  describe('sanitizeTasks', () => {
    it('should sanitize multiple tasks', () => {
      const tasks: ToDoTask[] = [
        {
          id: 'task-1',
          title: 'Task 1',
          status: 'notStarted',
          importance: 'high',
          createdDateTime: '2025-01-01T00:00:00Z',
          lastModifiedDateTime: '2025-01-02T00:00:00Z',
          isReminderOn: false,
        },
        {
          id: 'task-2',
          title: 'Task 2',
          status: 'completed',
          importance: 'normal',
          createdDateTime: '2025-01-03T00:00:00Z',
          lastModifiedDateTime: '2025-01-04T00:00:00Z',
          isReminderOn: true,
        },
      ];

      const sanitized = sanitizeTasks(tasks);

      expect(sanitized).toHaveLength(2);
      expect(sanitized[0].id).toBe('task-1');
      expect(sanitized[1].id).toBe('task-2');
    });

    it('should handle empty array', () => {
      const sanitized = sanitizeTasks([]);
      expect(sanitized).toEqual([]);
    });
  });

  describe('sanitizeTaskList', () => {
    it('should sanitize a task list', () => {
      const list: ToDoTaskList = {
        id: 'list-123',
        displayName: 'My Tasks',
        isOwner: true,
        isShared: false,
        wellknownListName: 'defaultList',
      };

      const sanitized = sanitizeTaskList(list);

      expect(sanitized).toEqual({
        id: 'list-123',
        displayName: 'My Tasks',
        isOwner: true,
        isShared: false,
        wellknownListName: 'defaultList',
      });
    });

    it('should handle list without wellknownListName', () => {
      const list: ToDoTaskList = {
        id: 'list-456',
        displayName: 'Work Tasks',
        isOwner: true,
        isShared: true,
      };

      const sanitized = sanitizeTaskList(list);

      expect(sanitized).toEqual({
        id: 'list-456',
        displayName: 'Work Tasks',
        isOwner: true,
        isShared: true,
        wellknownListName: undefined,
      });
    });
  });

  describe('sanitizeTaskLists', () => {
    it('should sanitize multiple task lists', () => {
      const lists: ToDoTaskList[] = [
        {
          id: 'list-1',
          displayName: 'Personal',
          isOwner: true,
          isShared: false,
        },
        {
          id: 'list-2',
          displayName: 'Work',
          isOwner: true,
          isShared: true,
        },
      ];

      const sanitized = sanitizeTaskLists(lists);

      expect(sanitized).toHaveLength(2);
      expect(sanitized[0].id).toBe('list-1');
      expect(sanitized[1].id).toBe('list-2');
    });

    it('should handle empty array', () => {
      const sanitized = sanitizeTaskLists([]);
      expect(sanitized).toEqual([]);
    });
  });

  describe('sanitizeError', () => {
    it('should sanitize ValidationError', () => {
      const error = new ValidationError('Invalid input provided');
      const sanitized = sanitizeError(error);

      expect(sanitized).toEqual({
        error: 'Invalid input provided',
        code: 'ValidationError',
      });
    });

    it('should sanitize RateLimitError with retryAfter', () => {
      const error = new RateLimitError('Rate limit exceeded', 5000);
      const sanitized = sanitizeError(error);

      expect(sanitized).toEqual({
        error: 'Rate limit exceeded',
        code: 'RateLimitError',
      });
    });

    it('should sanitize generic Error', () => {
      const error = new Error('Something went wrong');
      const sanitized = sanitizeError(error);

      expect(sanitized).toEqual({
        error: 'Something went wrong',
        // Generic Error doesn't have a code
      });
      expect(sanitized.code).toBeUndefined();
    });

    it('should handle non-Error objects', () => {
      const sanitized = sanitizeError('string error');

      expect(sanitized).toEqual({
        error: 'An unexpected error occurred',
      });
    });

    it('should handle null/undefined', () => {
      expect(sanitizeError(null)).toEqual({
        error: 'An unexpected error occurred',
      });

      expect(sanitizeError(undefined)).toEqual({
        error: 'An unexpected error occurred',
      });
    });

    it('should never expose stack traces', () => {
      const error = new Error('Test error');
      const sanitized = sanitizeError(error);

      expect(sanitized).not.toHaveProperty('stack');
      expect(JSON.stringify(sanitized)).not.toContain('stack');
    });
  });
});
