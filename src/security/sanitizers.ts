/**
 * Response sanitizers to ensure no sensitive data is exposed to clients.
 * All Graph API responses should pass through sanitizers before returning.
 */

import { ToDoTask, ToDoTaskList } from '../graph/todo-service.js';
import { SanitizedTask, SanitizedTaskList } from '../mcp/types.js';
import { logger } from './logger.js';

/**
 * Sanitize a single task for client consumption.
 * Removes internal Microsoft Graph fields and sensitive data.
 *
 * @param task - Raw task from Microsoft Graph API
 * @returns Sanitized task safe for client
 */
export function sanitizeTask(task: ToDoTask): SanitizedTask {
  return {
    id: task.id,
    title: task.title,
    status: task.status,
    importance: task.importance,
    createdDateTime: task.createdDateTime,
    lastModifiedDateTime: task.lastModifiedDateTime,
    dueDateTime: task.dueDateTime?.dateTime,
    isReminderOn: task.isReminderOn,
    body: task.body?.content,
    categories: task.categories,
  };
}

/**
 * Sanitize multiple tasks.
 *
 * @param tasks - Array of raw tasks
 * @returns Array of sanitized tasks
 */
export function sanitizeTasks(tasks: ToDoTask[]): SanitizedTask[] {
  logger.debug('Sanitizing tasks', { count: tasks.length });
  return tasks.map(sanitizeTask);
}

/**
 * Sanitize a task list for client consumption.
 *
 * @param list - Raw task list from Microsoft Graph API
 * @returns Sanitized task list
 */
export function sanitizeTaskList(list: ToDoTaskList): SanitizedTaskList {
  return {
    id: list.id,
    displayName: list.displayName,
    isOwner: list.isOwner,
    isShared: list.isShared,
    wellknownListName: list.wellknownListName,
  };
}

/**
 * Sanitize multiple task lists.
 *
 * @param lists - Array of raw task lists
 * @returns Array of sanitized task lists
 */
export function sanitizeTaskLists(lists: ToDoTaskList[]): SanitizedTaskList[] {
  logger.debug('Sanitizing task lists', { count: lists.length });
  return lists.map(sanitizeTaskList);
}

/**
 * Sanitize error for client response.
 * Never expose stack traces, internal paths, or sensitive details.
 *
 * @param error - Error object
 * @returns Safe error message
 */
export function sanitizeError(error: unknown): { error: string; code?: string } {
  if (error instanceof Error) {
    // Use error name as code (e.g., "ValidationError", "RateLimitError")
    const code = error.name !== 'Error' ? error.name : undefined;

    logger.debug('Sanitizing error', {
      errorName: error.name,
      // Don't log error message - might contain sensitive data
    });

    return {
      error: error.message, // Error messages are already sanitized by our error classes
      code,
    };
  }

  // Unknown error type
  logger.warn('Unknown error type during sanitization', {
    errorType: typeof error,
  });

  return {
    error: 'An unexpected error occurred',
  };
}
