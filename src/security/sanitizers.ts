/**
 * Response sanitizers to ensure no sensitive data is exposed to clients.
 * All Graph API responses should pass through sanitizers before returning.
 */

import { ToDoTask, ToDoTaskList } from '../graph/todo-service.js';
import { CalendarEvent } from '../graph/calendar-service.js';
import {
  PlannerPlan,
  PlannerTask,
  PlannerBucket,
  PlannerTaskDetails,
} from '../graph/planner-service.js';
import {
  SanitizedTask,
  SanitizedTaskList,
  SanitizedCalendarEvent,
  SanitizedPlannerPlan,
  SanitizedPlannerTask,
  SanitizedPlannerBucket,
  SanitizedPlannerTaskDetails,
} from '../mcp/types.js';
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

// ============================================
// Calendar sanitizers
// ============================================

/**
 * Sanitize a single calendar event for client consumption.
 * Removes internal Microsoft Graph fields and sensitive data.
 *
 * @param event - Raw event from Microsoft Graph API
 * @returns Sanitized event safe for client
 */
export function sanitizeCalendarEvent(event: CalendarEvent): SanitizedCalendarEvent {
  return {
    id: event.id,
    subject: event.subject,
    start: event.start.dateTime,
    end: event.end.dateTime,
    startTimeZone: event.start.timeZone,
    endTimeZone: event.end.timeZone,
    location: event.location?.displayName,
    organizer: event.organizer ? {
      name: event.organizer.emailAddress.name,
      email: event.organizer.emailAddress.address,
    } : undefined,
    attendees: event.attendees?.map(a => ({
      name: a.emailAddress.name,
      email: a.emailAddress.address,
      type: a.type,
      response: a.status?.response,
    })),
    isAllDay: event.isAllDay,
    webLink: event.webLink,
  };
}

/**
 * Sanitize multiple calendar events.
 *
 * @param events - Array of raw events
 * @returns Array of sanitized events
 */
export function sanitizeCalendarEvents(events: CalendarEvent[]): SanitizedCalendarEvent[] {
  logger.debug('Sanitizing calendar events', { count: events.length });
  return events.map(sanitizeCalendarEvent);
}

// ============================================
// Planner sanitizers
// ============================================

/**
 * Convert Planner priority number to human-readable string.
 * Planner uses 0-10, where lower = higher priority.
 */
function priorityToString(priority: number): string {
  if (priority <= 1) return 'urgent';
  if (priority <= 3) return 'high';
  if (priority <= 5) return 'normal';
  return 'low';
}

/**
 * Sanitize a single Planner plan for client consumption.
 *
 * @param plan - Raw plan from Microsoft Graph API
 * @returns Sanitized plan
 */
export function sanitizePlannerPlan(plan: PlannerPlan): SanitizedPlannerPlan {
  return {
    id: plan.id,
    title: plan.title,
    createdDateTime: plan.createdDateTime,
    owner: plan.owner,
  };
}

/**
 * Sanitize multiple Planner plans.
 *
 * @param plans - Array of raw plans
 * @returns Array of sanitized plans
 */
export function sanitizePlannerPlans(plans: PlannerPlan[]): SanitizedPlannerPlan[] {
  logger.debug('Sanitizing Planner plans', { count: plans.length });
  return plans.map(sanitizePlannerPlan);
}

/**
 * Sanitize a single Planner task for client consumption.
 * Converts priority number to string and counts assignees.
 *
 * @param task - Raw task from Microsoft Graph API
 * @returns Sanitized task
 */
export function sanitizePlannerTask(task: PlannerTask): SanitizedPlannerTask {
  return {
    id: task.id,
    planId: task.planId,
    bucketId: task.bucketId,
    title: task.title,
    percentComplete: task.percentComplete,
    priority: priorityToString(task.priority),
    dueDateTime: task.dueDateTime,
    startDateTime: task.startDateTime,
    createdDateTime: task.createdDateTime,
    assigneeCount: Object.keys(task.assignments || {}).length,
  };
}

/**
 * Sanitize multiple Planner tasks.
 *
 * @param tasks - Array of raw tasks
 * @returns Array of sanitized tasks
 */
export function sanitizePlannerTasks(tasks: PlannerTask[]): SanitizedPlannerTask[] {
  logger.debug('Sanitizing Planner tasks', { count: tasks.length });
  return tasks.map(sanitizePlannerTask);
}

/**
 * Sanitize a single Planner bucket for client consumption.
 *
 * @param bucket - Raw bucket from Microsoft Graph API
 * @returns Sanitized bucket
 */
export function sanitizePlannerBucket(bucket: PlannerBucket): SanitizedPlannerBucket {
  return {
    id: bucket.id,
    planId: bucket.planId,
    name: bucket.name,
  };
}

/**
 * Sanitize multiple Planner buckets.
 *
 * @param buckets - Array of raw buckets
 * @returns Array of sanitized buckets
 */
export function sanitizePlannerBuckets(buckets: PlannerBucket[]): SanitizedPlannerBucket[] {
  logger.debug('Sanitizing Planner buckets', { count: buckets.length });
  return buckets.map(sanitizePlannerBucket);
}

/**
 * Sanitize Planner task details for client consumption.
 * Converts checklist to simple array format.
 *
 * @param details - Raw task details from Microsoft Graph API
 * @returns Sanitized task details
 */
export function sanitizePlannerTaskDetails(details: PlannerTaskDetails): SanitizedPlannerTaskDetails {
  return {
    id: details.id,
    description: details.description || '',
    checklist: Object.values(details.checklist || {}).map(item => ({
      title: item.title,
      isChecked: item.isChecked,
    })),
  };
}
