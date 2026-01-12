/**
 * MCP server types and Zod schemas for input validation.
 * All tool inputs must be validated before processing.
 */

import { z } from 'zod';

/**
 * Schema for get_auth_status tool (no inputs required).
 */
export const GetAuthStatusSchema = z.object({});

export type GetAuthStatusInput = z.infer<typeof GetAuthStatusSchema>;

/**
 * Schema for list_task_lists tool (no inputs required).
 */
export const ListTaskListsSchema = z.object({});

export type ListTaskListsInput = z.infer<typeof ListTaskListsSchema>;

/**
 * Schema for list_tasks tool.
 */
export const ListTasksSchema = z.object({
  listId: z.string().optional().describe('Task list ID (optional, uses default list if not provided)'),
  filter: z.enum(['all', 'completed', 'incomplete', 'high-priority', 'today', 'overdue', 'this-week', 'later'])
    .optional()
    .describe('Filter tasks by status, priority, or due date'),
  limit: z.number().int().min(1).max(100).optional().describe('Maximum number of tasks to return (1-100)'),
});

export type ListTasksInput = z.infer<typeof ListTasksSchema>;

/**
 * Schema for get_task tool.
 */
export const GetTaskSchema = z.object({
  listId: z.string().describe('Task list ID'),
  taskId: z.string().describe('Task ID'),
});

export type GetTaskInput = z.infer<typeof GetTaskSchema>;

/**
 * Schema for search_tasks tool.
 */
export const SearchTasksSchema = z.object({
  query: z.string().min(1).max(100).describe('Search query (1-100 characters)'),
  limit: z.number().int().min(1).max(50).optional().describe('Maximum number of results (1-50)'),
});

export type SearchTasksInput = z.infer<typeof SearchTasksSchema>;

/**
 * Sanitized task response (safe to return to clients).
 */
export interface SanitizedTask {
  id: string;
  title: string;
  status: string;
  importance: string;
  createdDateTime: string;
  lastModifiedDateTime: string;
  dueDateTime?: string;
  isReminderOn: boolean;
  body?: string;
  categories?: string[];
}

/**
 * Sanitized task list response.
 */
export interface SanitizedTaskList {
  id: string;
  displayName: string;
  isOwner: boolean;
  isShared: boolean;
  wellknownListName?: string;
}

/**
 * Authentication status response.
 */
export interface AuthStatus {
  authenticated: boolean;
  message: string;
}

/**
 * Request context for logging and tracking.
 */
export interface RequestContext {
  requestId: string;
  tool: string;
  timestamp: number;
}

// ============================================
// Calendar schemas and types
// ============================================

/**
 * Schema for list_calendar_events tool.
 */
export const ListCalendarEventsSchema = z.object({
  filter: z.enum(['today', 'this-week', 'this-month'])
    .optional()
    .default('today')
    .describe('Filter events by time range (default: today)'),
  limit: z.number().int().min(1).max(100).optional().describe('Maximum number of events to return (1-100)'),
});

export type ListCalendarEventsInput = z.infer<typeof ListCalendarEventsSchema>;

/**
 * Schema for get_calendar_view tool.
 */
export const GetCalendarViewSchema = z.object({
  startDate: z.string().describe('Start date in ISO format (e.g., 2025-01-15)'),
  endDate: z.string().describe('End date in ISO format (e.g., 2025-01-20)'),
  limit: z.number().int().min(1).max(100).optional().describe('Maximum number of events to return (1-100)'),
});

export type GetCalendarViewInput = z.infer<typeof GetCalendarViewSchema>;

/**
 * Sanitized calendar event response.
 */
export interface SanitizedCalendarEvent {
  id: string;
  subject: string;
  start: string;
  end: string;
  startTimeZone: string;
  endTimeZone: string;
  location?: string;
  organizer?: {
    name: string;
    email: string;
  };
  attendees?: Array<{
    name: string;
    email: string;
    type: string;
    response?: string;
  }>;
  isAllDay: boolean;
  webLink?: string;
}
