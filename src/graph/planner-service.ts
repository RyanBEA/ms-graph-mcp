/**
 * Microsoft Planner service layer.
 * Provides high-level operations for plans, buckets, and tasks.
 */

import { SecureGraphClient } from './client.js';
import { logger } from '../security/logger.js';
import {
  validateListId,
  validateTaskId,
  validateTaskTitle,
  validateDueDate,
  validatePercentComplete,
} from '../security/validators.js';

// ============================================
// Interfaces
// ============================================

/**
 * Microsoft Planner plan.
 */
export interface PlannerPlan {
  id: string;
  title: string;
  createdDateTime: string;
  owner: string;
  createdBy?: {
    user: {
      id: string;
      displayName?: string;
    };
  };
  '@odata.etag'?: string;
}

/**
 * Microsoft Planner task.
 */
export interface PlannerTask {
  id: string;
  planId: string;
  bucketId: string;
  title: string;
  percentComplete: number;  // 0, 50, or 100
  priority: number;         // 0-10 (lower = higher priority)
  dueDateTime?: string;
  startDateTime?: string;
  createdDateTime: string;
  assignments: Record<string, {
    assignedBy: {
      user: { id: string };
    };
    assignedDateTime?: string;
    orderHint?: string;
  }>;
  orderHint?: string;
  '@odata.etag'?: string;
}

/**
 * Microsoft Planner bucket (column in a plan board).
 */
export interface PlannerBucket {
  id: string;
  planId: string;
  name: string;
  orderHint: string;
  '@odata.etag'?: string;
}

/**
 * Microsoft Planner task details (description, checklist, etc.).
 */
export interface PlannerTaskDetails {
  id: string;
  description: string;
  checklist: Record<string, {
    title: string;
    isChecked: boolean;
    orderHint?: string;
  }>;
  references: Record<string, {
    alias: string;
    type: string;
    previewPriority?: string;
  }>;
  '@odata.etag'?: string;
}

// Reuse list/task ID validators (same format)
export const validatePlanId = validateListId;
export const validateBucketId = validateListId;

// ============================================
// Service
// ============================================

/**
 * Service for Microsoft Planner operations.
 */
export class PlannerService {
  constructor(private readonly graphClient: SecureGraphClient) {
    logger.info('PlannerService initialized');
  }

  /**
   * Get all Planner tasks assigned to the authenticated user.
   *
   * @param options - Query options
   * @returns Array of tasks
   */
  async getMyTasks(options?: { top?: number }): Promise<PlannerTask[]> {
    logger.debug('Fetching user Planner tasks');

    const queryParams: Record<string, string> = {};
    if (options?.top) {
      queryParams.$top = options.top.toString();
    }

    const response = await this.graphClient.get<PlannerTask>('/me/planner/tasks', { queryParams });

    const tasks = response.value ?? [];
    logger.info('User Planner tasks retrieved', { count: tasks.length });

    return tasks;
  }

  /**
   * Get a specific Planner plan by ID.
   *
   * @param planId - Planner plan identifier
   * @returns Plan details
   */
  async getPlan(planId: string): Promise<PlannerPlan> {
    validatePlanId(planId);

    logger.debug('Fetching Planner plan', { planId });

    const response = await this.graphClient.get<PlannerPlan>(`/planner/plans/${planId}`);

    // Single resource response
    return response as unknown as PlannerPlan;
  }

  /**
   * Get all tasks in a Planner plan.
   *
   * @param planId - Planner plan identifier
   * @param options - Query options
   * @returns Array of tasks
   */
  async getPlanTasks(
    planId: string,
    options?: { bucketId?: string; top?: number }
  ): Promise<PlannerTask[]> {
    validatePlanId(planId);
    if (options?.bucketId) {
      validateBucketId(options.bucketId);
    }

    logger.debug('Fetching plan tasks', { planId, bucketId: options?.bucketId });

    const queryParams: Record<string, string> = {};
    if (options?.top) {
      queryParams.$top = options.top.toString();
    }

    const response = await this.graphClient.get<PlannerTask>(
      `/planner/plans/${planId}/tasks`,
      { queryParams }
    );

    let tasks = response.value ?? [];

    // Filter by bucket if specified (Graph API doesn't support $filter on bucketId)
    if (options?.bucketId) {
      tasks = tasks.filter(t => t.bucketId === options.bucketId);
    }

    logger.info('Plan tasks retrieved', { planId, count: tasks.length });

    return tasks;
  }

  /**
   * Get all buckets (columns) in a Planner plan.
   *
   * @param planId - Planner plan identifier
   * @returns Array of buckets
   */
  async getPlanBuckets(planId: string): Promise<PlannerBucket[]> {
    validatePlanId(planId);

    logger.debug('Fetching plan buckets', { planId });

    const response = await this.graphClient.get<PlannerBucket>(`/planner/plans/${planId}/buckets`);

    const buckets = response.value ?? [];
    logger.info('Plan buckets retrieved', { planId, count: buckets.length });

    return buckets;
  }

  /**
   * Get a specific Planner task by ID.
   *
   * @param taskId - Planner task identifier
   * @returns Task details
   */
  async getTask(taskId: string): Promise<PlannerTask> {
    validateTaskId(taskId);

    logger.debug('Fetching Planner task', { taskId });

    const response = await this.graphClient.get<PlannerTask>(`/planner/tasks/${taskId}`);

    return response as unknown as PlannerTask;
  }

  /**
   * Get details (description, checklist) for a Planner task.
   *
   * @param taskId - Planner task identifier
   * @returns Task details including description and checklist
   */
  async getTaskDetails(taskId: string): Promise<PlannerTaskDetails> {
    validateTaskId(taskId);

    logger.debug('Fetching Planner task details', { taskId });

    const response = await this.graphClient.get<PlannerTaskDetails>(`/planner/tasks/${taskId}/details`);

    return response as unknown as PlannerTaskDetails;
  }

  /**
   * Get plans for a specific group.
   *
   * @param groupId - Microsoft 365 group identifier
   * @returns Array of plans
   */
  async getGroupPlans(groupId: string): Promise<PlannerPlan[]> {
    validatePlanId(groupId); // Group IDs have same format

    logger.debug('Fetching group plans', { groupId });

    const response = await this.graphClient.get<PlannerPlan>(`/groups/${groupId}/planner/plans`);

    const plans = response.value ?? [];
    logger.info('Group plans retrieved', { groupId, count: plans.length });

    return plans;
  }

  // ============================================
  // Helper methods
  // ============================================

  /**
   * Get incomplete tasks assigned to the user.
   *
   * @param top - Maximum number of tasks
   * @returns Array of incomplete tasks
   */
  async getIncompleteTasks(top: number = 50): Promise<PlannerTask[]> {
    const tasks = await this.getMyTasks({ top });
    return tasks.filter(t => t.percentComplete < 100);
  }

  /**
   * Get completed tasks assigned to the user.
   *
   * @param top - Maximum number of tasks
   * @returns Array of completed tasks
   */
  async getCompletedTasks(top: number = 50): Promise<PlannerTask[]> {
    const tasks = await this.getMyTasks({ top });
    return tasks.filter(t => t.percentComplete === 100);
  }

  // ============================================
  // Write Operations
  // ============================================

  /**
   * Create a new task in a Planner plan.
   *
   * @param planId - Planner plan ID
   * @param bucketId - Bucket ID
   * @param data - Task data
   * @returns Created task
   */
  async createTask(
    planId: string,
    bucketId: string,
    data: {
      title: string;
      dueDateTime?: string;
      priority?: number;  // 0-10
      assignments?: string[];  // User IDs
    }
  ): Promise<PlannerTask> {
    validatePlanId(planId);
    validateBucketId(bucketId);
    validateTaskTitle(data.title);
    if (data.dueDateTime) validateDueDate(data.dueDateTime);

    logger.info('Creating Planner task', { planId, bucketId, titleLength: data.title.length });

    const body: Record<string, any> = {
      planId,
      bucketId,
      title: data.title,
    };

    if (data.dueDateTime) {
      body.dueDateTime = new Date(data.dueDateTime).toISOString();
    }

    if (data.priority !== undefined) {
      body.priority = data.priority;
    }

    if (data.assignments?.length) {
      body.assignments = {};
      for (const userId of data.assignments) {
        body.assignments[userId] = {
          '@odata.type': '#microsoft.graph.plannerAssignment',
          orderHint: ' !',
        };
      }
    }

    const response = await this.graphClient.post<PlannerTask>(
      '/planner/tasks',
      body
    );

    logger.info('Planner task created', { taskId: response.id });
    return response;
  }

  /**
   * Update a Planner task.
   * Note: Planner requires an If-Match header with the task's @odata.etag.
   *
   * @param taskId - Task ID
   * @param etag - Task's etag from previous fetch
   * @param data - Fields to update
   * @returns Updated task
   */
  async updateTask(
    taskId: string,
    etag: string,
    data: {
      title?: string;
      dueDateTime?: string | null;
      priority?: number;
      percentComplete?: number;
      bucketId?: string;
    }
  ): Promise<PlannerTask> {
    validateTaskId(taskId);
    if (data.title) validateTaskTitle(data.title);
    if (data.dueDateTime) validateDueDate(data.dueDateTime);
    if (data.percentComplete !== undefined) validatePercentComplete(data.percentComplete);
    if (data.bucketId) validateBucketId(data.bucketId);

    logger.info('Updating Planner task', { taskId });

    const body: Record<string, any> = {};

    if (data.title !== undefined) body.title = data.title;
    if (data.priority !== undefined) body.priority = data.priority;
    if (data.percentComplete !== undefined) body.percentComplete = data.percentComplete;
    if (data.bucketId !== undefined) body.bucketId = data.bucketId;

    if (data.dueDateTime === null) {
      body.dueDateTime = null;
    } else if (data.dueDateTime) {
      body.dueDateTime = new Date(data.dueDateTime).toISOString();
    }

    await this.graphClient.patch<PlannerTask>(
      `/planner/tasks/${taskId}`,
      body,
      { headers: { 'If-Match': etag } }
    );

    // PATCH returns 204 No Content, so re-fetch to get updated task
    const updatedTask = await this.getTask(taskId);

    logger.info('Planner task updated', { taskId });
    return updatedTask;
  }

  /**
   * Complete a Planner task (set percentComplete to 100).
   */
  async completeTask(taskId: string, etag: string): Promise<PlannerTask> {
    return this.updateTask(taskId, etag, { percentComplete: 100 });
  }

  /**
   * Uncomplete a Planner task (set percentComplete to 0).
   */
  async uncompleteTask(taskId: string, etag: string): Promise<PlannerTask> {
    return this.updateTask(taskId, etag, { percentComplete: 0 });
  }

  /**
   * Move a task to a different bucket.
   */
  async moveTask(taskId: string, etag: string, bucketId: string): Promise<PlannerTask> {
    return this.updateTask(taskId, etag, { bucketId });
  }

  /**
   * Delete a Planner task.
   */
  async deleteTask(taskId: string, etag: string): Promise<void> {
    validateTaskId(taskId);

    logger.info('Deleting Planner task', { taskId });

    await this.graphClient.delete(
      `/planner/tasks/${taskId}`,
      { headers: { 'If-Match': etag } }
    );

    logger.info('Planner task deleted', { taskId });
  }
}
