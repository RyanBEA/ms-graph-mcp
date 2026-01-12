/**
 * Microsoft To Do service layer.
 * Provides high-level operations for task lists and tasks.
 */

import { SecureGraphClient } from './client.js';
import { logger } from '../security/logger.js';
import {
  validateListId,
  validateTaskId,
  validateODataFilter,
  validateSearchQuery,
} from '../security/validators.js';

/**
 * Microsoft To Do task list.
 */
export interface ToDoTaskList {
  id: string;
  displayName: string;
  isOwner: boolean;
  isShared: boolean;
  wellknownListName?: string;
}

/**
 * Microsoft To Do task.
 */
export interface ToDoTask {
  id: string;
  title: string;
  status: 'notStarted' | 'inProgress' | 'completed' | 'waitingOnOthers' | 'deferred';
  importance: 'low' | 'normal' | 'high';
  createdDateTime: string;
  lastModifiedDateTime: string;
  dueDateTime?: {
    dateTime: string;
    timeZone: string;
  };
  reminderDateTime?: {
    dateTime: string;
    timeZone: string;
  };
  isReminderOn: boolean;
  body?: {
    content: string;
    contentType: string;
  };
  categories?: string[];
}

/**
 * Service for Microsoft To Do operations.
 */
export class ToDoService {
  constructor(private readonly graphClient: SecureGraphClient) {
    logger.info('ToDoService initialized');
  }

  /**
   * Get all task lists for the authenticated user.
   *
   * @returns Array of task lists
   */
  async getTaskLists(): Promise<ToDoTaskList[]> {
    logger.debug('Fetching task lists');

    const response = await this.graphClient.get<ToDoTaskList>('/me/todo/lists', {
      // TEMP: Remove $select to test if that's causing 400 errors
      // queryParams: {
      //   $select: 'id,displayName,isOwner,isShared,wellknownListName',
      // },
    });

    const lists = response.value ?? [];
    logger.info('Task lists retrieved', {
      count: lists.length,
    });

    return lists;
  }

  /**
   * Get a specific task list by ID.
   *
   * @param listId - Task list identifier
   * @returns Task list details
   */
  async getTaskList(listId: string): Promise<ToDoTaskList> {
    validateListId(listId);

    logger.debug('Fetching task list', { listId });

    const response = await this.graphClient.get<ToDoTaskList>(
      `/me/todo/lists/${listId}`
      // Removed $select - causes 400 errors with current token
    );

    // Single resource response doesn't have 'value' array
    return response as unknown as ToDoTaskList;
  }

  /**
   * Get tasks from a specific list.
   *
   * @param listId - Task list identifier (optional, defaults to default list)
   * @param options - Query options
   * @returns Array of tasks
   */
  async getTasks(
    listId?: string,
    options?: {
      filter?: string;
      top?: number;
      orderBy?: string;
    }
  ): Promise<ToDoTask[]> {
    // Use default list if not specified
    const effectiveListId = listId ?? await this.getDefaultListId();
    validateListId(effectiveListId);

    // Validate filter if provided
    if (options?.filter) {
      validateODataFilter(options.filter);
    }

    logger.debug('Fetching tasks', {
      listId: effectiveListId,
      hasFilter: !!options?.filter,
    });

    const queryParams: Record<string, string> = {
      // Removed $select - causes 400 errors with current token
    };

    if (options?.filter) {
      queryParams.$filter = options.filter;
    }

    if (options?.top) {
      queryParams.$top = options.top.toString();
    }

    if (options?.orderBy) {
      queryParams.$orderby = options.orderBy;
    }

    const response = await this.graphClient.get<ToDoTask>(
      `/me/todo/lists/${effectiveListId}/tasks`,
      { queryParams }
    );

    const tasks = response.value ?? [];
    logger.info('Tasks retrieved', {
      listId: effectiveListId,
      count: tasks.length,
    });

    return tasks;
  }

  /**
   * Get a specific task by ID.
   *
   * @param listId - Task list identifier
   * @param taskId - Task identifier
   * @returns Task details
   */
  async getTask(listId: string, taskId: string): Promise<ToDoTask> {
    validateListId(listId);
    validateTaskId(taskId);

    logger.debug('Fetching task', { listId, taskId });

    const response = await this.graphClient.get<ToDoTask>(
      `/me/todo/lists/${listId}/tasks/${taskId}`
      // Removed $select - causes 400 errors with current token
    );

    // Single resource response
    return response as unknown as ToDoTask;
  }

  /**
   * Search tasks across all lists.
   *
   * @param query - Search query string
   * @param options - Search options
   * @returns Array of tasks matching the query
   */
  async searchTasks(
    query: string,
    options?: {
      top?: number;
      filter?: string;
    }
  ): Promise<ToDoTask[]> {
    validateSearchQuery(query);

    if (options?.filter) {
      validateODataFilter(options.filter);
    }

    logger.debug('Searching tasks', {
      queryLength: query.length,
      hasFilter: !!options?.filter,
    });

    // Get all task lists
    const lists = await this.getTaskLists();

    // Search in each list
    const allTasks: ToDoTask[] = [];
    for (const list of lists) {
      try {
        const tasks = await this.getTasks(list.id, {
          top: options?.top,
          filter: options?.filter,
        });

        // Filter by query (case-insensitive title match)
        const matchingTasks = tasks.filter((task) =>
          task.title.toLowerCase().includes(query.toLowerCase())
        );

        allTasks.push(...matchingTasks);
      } catch (error) {
        // Log error but continue with other lists
        logger.warn('Failed to search tasks in list', {
          listId: list.id,
          error: error instanceof Error ? error.message : 'Unknown error',
        });
      }
    }

    logger.info('Task search completed', {
      query: query.substring(0, 50), // Log truncated query
      resultsCount: allTasks.length,
      listsSearched: lists.length,
    });

    return allTasks;
  }

  /**
   * Get completed tasks from a list.
   *
   * @param listId - Task list identifier (optional)
   * @param top - Maximum number of tasks to return
   * @returns Array of completed tasks
   */
  async getCompletedTasks(listId?: string, top: number = 50): Promise<ToDoTask[]> {
    return this.getTasks(listId, {
      filter: "status eq 'completed'",
      top,
      orderBy: 'lastModifiedDateTime desc',
    });
  }

  /**
   * Get high-priority tasks from a list.
   *
   * @param listId - Task list identifier (optional)
   * @param top - Maximum number of tasks to return
   * @returns Array of high-priority tasks
   */
  async getHighPriorityTasks(listId?: string, top: number = 50): Promise<ToDoTask[]> {
    return this.getTasks(listId, {
      filter: "importance eq 'high' and status ne 'completed'",
      top,
      orderBy: 'dueDateTime asc',
    });
  }

  /**
   * Get tasks due today.
   *
   * @param listId - Task list identifier (optional)
   * @returns Array of tasks due today
   */
  async getTasksDueToday(listId?: string): Promise<ToDoTask[]> {
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const tomorrow = new Date(today);
    tomorrow.setDate(tomorrow.getDate() + 1);

    // If listId provided, search that list only; otherwise search all lists
    const tasks = listId
      ? await this.getTasks(listId, { filter: "status ne 'completed'" })
      : await this.getAllIncompleteTasks();

    return tasks
      .filter((task) => {
        if (!task.dueDateTime?.dateTime) return false;
        const dueDate = new Date(task.dueDateTime.dateTime);
        return dueDate >= today && dueDate < tomorrow;
      })
      .sort((a, b) => {
        const aDate = new Date(a.dueDateTime!.dateTime);
        const bDate = new Date(b.dueDateTime!.dateTime);
        return aDate.getTime() - bDate.getTime();
      });
  }

  /**
   * Get overdue tasks (due date in the past).
   *
   * @param listId - Task list identifier (optional)
   * @returns Array of overdue tasks
   */
  async getTasksDueOverdue(listId?: string): Promise<ToDoTask[]> {
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    // If listId provided, search that list only; otherwise search all lists
    const tasks = listId
      ? await this.getTasks(listId, { filter: "status ne 'completed'" })
      : await this.getAllIncompleteTasks();

    return tasks
      .filter((task) => {
        if (!task.dueDateTime?.dateTime) return false;
        const dueDate = new Date(task.dueDateTime.dateTime);
        return dueDate < today;
      })
      .sort((a, b) => {
        const aDate = new Date(a.dueDateTime!.dateTime);
        const bDate = new Date(b.dueDateTime!.dateTime);
        return aDate.getTime() - bDate.getTime();
      });
  }

  /**
   * Get tasks due this week (next 7 days).
   *
   * @param listId - Task list identifier (optional)
   * @returns Array of tasks due in the next 7 days
   */
  async getTasksDueThisWeek(listId?: string): Promise<ToDoTask[]> {
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const nextWeek = new Date(today);
    nextWeek.setDate(today.getDate() + 7);

    // If listId provided, search that list only; otherwise search all lists
    const tasks = listId
      ? await this.getTasks(listId, { filter: "status ne 'completed'" })
      : await this.getAllIncompleteTasks();

    return tasks
      .filter((task) => {
        if (!task.dueDateTime?.dateTime) return false;
        const dueDate = new Date(task.dueDateTime.dateTime);
        return dueDate >= today && dueDate < nextWeek;
      })
      .sort((a, b) => {
        const aDate = new Date(a.dueDateTime!.dateTime);
        const bDate = new Date(b.dueDateTime!.dateTime);
        return aDate.getTime() - bDate.getTime();
      });
  }

  /**
   * Get tasks due later (more than 7 days from now or no due date).
   *
   * @param listId - Task list identifier (optional)
   * @returns Array of tasks due later or with no due date
   */
  async getTasksDueLater(listId?: string): Promise<ToDoTask[]> {
    const nextWeek = new Date();
    nextWeek.setHours(0, 0, 0, 0);
    nextWeek.setDate(nextWeek.getDate() + 7);

    // If listId provided, search that list only; otherwise search all lists
    const tasks = listId
      ? await this.getTasks(listId, { filter: "status ne 'completed'" })
      : await this.getAllIncompleteTasks();

    return tasks
      .filter((task) => {
        if (!task.dueDateTime?.dateTime) return true; // Include tasks with no due date
        const dueDate = new Date(task.dueDateTime.dateTime);
        return dueDate >= nextWeek;
      })
      .sort((a, b) => {
        // Tasks without due dates come last
        if (!a.dueDateTime?.dateTime) return 1;
        if (!b.dueDateTime?.dateTime) return -1;
        const aDate = new Date(a.dueDateTime.dateTime);
        const bDate = new Date(b.dueDateTime.dateTime);
        return aDate.getTime() - bDate.getTime();
      });
  }

/**
   * Get all incomplete tasks across all lists.
   * Used for global date-based filtering when no listId is provided.
   */
  private async getAllIncompleteTasks(): Promise<ToDoTask[]> {
    logger.debug('Fetching incomplete tasks from all lists');

    const lists = await this.getTaskLists();
    const allTasks: ToDoTask[] = [];

    for (const list of lists) {
      try {
        const tasks = await this.getTasks(list.id, {
          filter: "status ne 'completed'",
        });
        allTasks.push(...tasks);
      } catch (error) {
        logger.warn('Failed to fetch tasks from list', {
          listId: list.id,
          error: error instanceof Error ? error.message : 'Unknown error',
        });
      }
    }

    logger.info('All incomplete tasks retrieved', {
      count: allTasks.length,
      listsSearched: lists.length,
    });

    return allTasks;
  }
  /**
   * Get the default task list ID.
   * Microsoft To Do has a built-in default list.
   */
  private async getDefaultListId(): Promise<string> {
    const lists = await this.getTaskLists();

    // Try to find the well-known default list
    const defaultList = lists.find((list) => list.wellknownListName === 'defaultList');

    if (defaultList) {
      return defaultList.id;
    }

    // Fallback to first list
    if (lists.length > 0) {
      logger.warn('Default list not found, using first list', {
        listId: lists[0].id,
      });
      return lists[0].id;
    }

    throw new Error('No task lists found for user');
  }
}
