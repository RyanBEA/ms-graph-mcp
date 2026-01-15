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

// ============================================
// Planner schemas
// ============================================

/**
 * Schema for list_my_planner_tasks tool.
 */
export const ListMyPlannerTasksSchema = z.object({
  filter: z.enum(['all', 'incomplete', 'completed'])
    .optional()
    .describe('Filter tasks by completion status'),
  limit: z.number().int().min(1).max(100).optional()
    .describe('Maximum number of tasks to return (1-100)'),
});

export type ListMyPlannerTasksInput = z.infer<typeof ListMyPlannerTasksSchema>;

/**
 * Schema for get_planner_plan tool.
 */
export const GetPlannerPlanSchema = z.object({
  planId: z.string().describe('Planner plan ID'),
});

export type GetPlannerPlanInput = z.infer<typeof GetPlannerPlanSchema>;

/**
 * Schema for list_plan_tasks tool.
 */
export const ListPlanTasksSchema = z.object({
  planId: z.string().describe('Planner plan ID'),
  bucketId: z.string().optional().describe('Filter by bucket ID'),
  limit: z.number().int().min(1).max(100).optional()
    .describe('Maximum number of tasks to return (1-100)'),
});

export type ListPlanTasksInput = z.infer<typeof ListPlanTasksSchema>;

/**
 * Schema for list_plan_buckets tool.
 */
export const ListPlanBucketsSchema = z.object({
  planId: z.string().describe('Planner plan ID'),
});

export type ListPlanBucketsInput = z.infer<typeof ListPlanBucketsSchema>;

/**
 * Schema for get_planner_task tool.
 */
export const GetPlannerTaskSchema = z.object({
  taskId: z.string().describe('Planner task ID'),
  includeDetails: z.boolean().optional()
    .describe('Include task details (description, checklist)'),
});

export type GetPlannerTaskInput = z.infer<typeof GetPlannerTaskSchema>;

// ============================================
// To Do Write schemas
// ============================================

/**
 * Schema for create_task tool.
 */
export const CreateTaskSchema = z.object({
  listId: z.string().optional().describe('Task list ID (uses default list if not provided)'),
  title: z.string().min(1).max(400).describe('Task title'),
  dueDateTime: z.string().optional().describe('Due date in ISO format (e.g., 2025-01-20)'),
  importance: z.enum(['low', 'normal', 'high']).optional().describe('Task importance'),
  body: z.string().max(10000).optional().describe('Task body/description'),
});

export type CreateTaskInput = z.infer<typeof CreateTaskSchema>;

/**
 * Schema for update_task tool.
 */
export const UpdateTaskSchema = z.object({
  listId: z.string().describe('Task list ID'),
  taskId: z.string().describe('Task ID'),
  title: z.string().min(1).max(400).optional().describe('New title'),
  dueDateTime: z.string().nullable().optional().describe('New due date (null to clear)'),
  importance: z.enum(['low', 'normal', 'high']).optional().describe('New importance'),
  status: z.enum(['notStarted', 'inProgress', 'completed']).optional().describe('New status'),
  body: z.string().max(10000).optional().describe('New body/description'),
});

export type UpdateTaskInput = z.infer<typeof UpdateTaskSchema>;

/**
 * Schema for complete_task / uncomplete_task tools.
 */
export const CompleteTaskSchema = z.object({
  listId: z.string().describe('Task list ID'),
  taskId: z.string().describe('Task ID'),
});

export type CompleteTaskInput = z.infer<typeof CompleteTaskSchema>;

/**
 * Schema for delete_task tool.
 */
export const DeleteTaskSchema = z.object({
  listId: z.string().describe('Task list ID'),
  taskId: z.string().describe('Task ID'),
  confirm: z.literal(true).describe('Must be true to confirm deletion'),
});

export type DeleteTaskInput = z.infer<typeof DeleteTaskSchema>;

/**
 * Schema for create_task_list tool.
 */
export const CreateTaskListSchema = z.object({
  displayName: z.string().min(1).max(400).describe('List name'),
});

export type CreateTaskListInput = z.infer<typeof CreateTaskListSchema>;

/**
 * Schema for delete_task_list tool.
 */
export const DeleteTaskListSchema = z.object({
  listId: z.string().describe('Task list ID'),
  confirm: z.literal(true).describe('Must be true to confirm deletion'),
});

export type DeleteTaskListInput = z.infer<typeof DeleteTaskListSchema>;

// ============================================
// Planner Write schemas
// ============================================

/**
 * Schema for create_planner_task tool.
 */
export const CreatePlannerTaskSchema = z.object({
  planId: z.string().describe('Planner plan ID'),
  bucketId: z.string().describe('Bucket ID'),
  title: z.string().min(1).max(400).describe('Task title'),
  dueDateTime: z.string().optional().describe('Due date in ISO format'),
  priority: z.enum(['urgent', 'high', 'normal', 'low']).optional().describe('Task priority'),
});

export type CreatePlannerTaskInput = z.infer<typeof CreatePlannerTaskSchema>;

/**
 * Schema for update_planner_task tool.
 */
export const UpdatePlannerTaskSchema = z.object({
  taskId: z.string().describe('Planner task ID'),
  title: z.string().min(1).max(400).optional().describe('New title'),
  dueDateTime: z.string().nullable().optional().describe('New due date (null to clear)'),
  priority: z.enum(['urgent', 'high', 'normal', 'low']).optional().describe('New priority'),
  percentComplete: z.union([z.literal(0), z.literal(50), z.literal(100)]).optional()
    .describe('Completion percentage (0, 50, or 100)'),
  bucketId: z.string().optional().describe('Move to different bucket'),
});

export type UpdatePlannerTaskInput = z.infer<typeof UpdatePlannerTaskSchema>;

/**
 * Schema for complete_planner_task / uncomplete_planner_task tools.
 */
export const CompletePlannerTaskSchema = z.object({
  taskId: z.string().describe('Planner task ID'),
});

export type CompletePlannerTaskInput = z.infer<typeof CompletePlannerTaskSchema>;

/**
 * Schema for delete_planner_task tool.
 */
export const DeletePlannerTaskSchema = z.object({
  taskId: z.string().describe('Planner task ID'),
  confirm: z.literal(true).describe('Must be true to confirm deletion'),
});

export type DeletePlannerTaskInput = z.infer<typeof DeletePlannerTaskSchema>;

// ============================================
// Sanitized Planner types
// ============================================

/**
 * Sanitized Planner plan response.
 */
export interface SanitizedPlannerPlan {
  id: string;
  title: string;
  createdDateTime: string;
  owner: string;
}

/**
 * Sanitized Planner task response.
 */
export interface SanitizedPlannerTask {
  id: string;
  planId: string;
  bucketId: string;
  title: string;
  percentComplete: number;
  priority: string;  // Converted from number to human-readable
  dueDateTime?: string;
  startDateTime?: string;
  createdDateTime: string;
  assigneeCount: number;  // Don't expose full assignment details
}

/**
 * Sanitized Planner bucket response.
 */
export interface SanitizedPlannerBucket {
  id: string;
  planId: string;
  name: string;
}

/**
 * Sanitized Planner task details response.
 */
export interface SanitizedPlannerTaskDetails {
  id: string;
  description: string;
  checklist: Array<{ title: string; isChecked: boolean }>;
}

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
