/**
 * MCP Server implementation for Microsoft Graph API.
 * Provides access to To Do tasks, Planner tasks, and Calendar events via secure, validated tools.
 */

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
  Tool,
} from '@modelcontextprotocol/sdk/types.js';
import { randomUUID } from 'crypto';

import { logger } from '../security/logger.js';
import {
  sanitizeError,
  sanitizeTasks,
  sanitizeTaskLists,
  sanitizeTask,
  sanitizeCalendarEvents,
  sanitizePlannerTasks,
  sanitizePlannerTask,
  sanitizePlannerPlan,
  sanitizePlannerBuckets,
  sanitizePlannerTaskDetails,
} from '../security/sanitizers.js';
import { ToDoService } from '../graph/todo-service.js';
import { CalendarService } from '../graph/calendar-service.js';
import { PlannerService } from '../graph/planner-service.js';
import { SecureGraphClient } from '../graph/client.js';
import { TokenRefresher } from '../auth/token-refresher.js';
import { getTokenManager } from '../auth/token-manager-factory.js';
import { validateDateString, validateDateRange } from '../security/validators.js';
import {
  GetAuthStatusSchema,
  ListTaskListsSchema,
  ListTasksSchema,
  GetTaskSchema,
  SearchTasksSchema,
  ListCalendarEventsSchema,
  GetCalendarViewSchema,
  // Planner schemas
  ListMyPlannerTasksSchema,
  GetPlannerPlanSchema,
  ListPlanTasksSchema,
  ListPlanBucketsSchema,
  GetPlannerTaskSchema,
  // Write schemas
  CreateTaskSchema,
  UpdateTaskSchema,
  CompleteTaskSchema,
  DeleteTaskSchema,
  CreateTaskListSchema,
  DeleteTaskListSchema,
  CreatePlannerTaskSchema,
  UpdatePlannerTaskSchema,
  CompletePlannerTaskSchema,
  DeletePlannerTaskSchema,
  RequestContext,
  AuthStatus,
} from './types.js';

/**
 * MsGraph MCP Server - Secure Microsoft Graph API integration.
 */
export class MsGraphMCPServer {
  private server: Server;
  private todoService: ToDoService;
  private calendarService: CalendarService;
  private plannerService: PlannerService;
  private tokenRefresher: TokenRefresher;

  private constructor(
    todoService: ToDoService,
    calendarService: CalendarService,
    plannerService: PlannerService,
    tokenRefresher: TokenRefresher
  ) {
    logger.info('Initializing MsGraph MCP Server');

    this.todoService = todoService;
    this.calendarService = calendarService;
    this.plannerService = plannerService;
    this.tokenRefresher = tokenRefresher;

    // Initialize server
    this.server = new Server(
      {
        name: 'ms-graph-mcp',
        version: '1.0.0',
      },
      {
        capabilities: {
          tools: {},
        },
      }
    );

    // Set up handlers
    this.setupHandlers();

    logger.info('MsGraph MCP Server initialized successfully');
  }

  /**
   * Create a new MsGraphMCPServer instance.
   * This is async because token manager initialization is async.
   */
  static async create(): Promise<MsGraphMCPServer> {
    const tokenManager = await getTokenManager();
    const tokenRefresher = new TokenRefresher(tokenManager);
    const graphClient = new SecureGraphClient(tokenRefresher);
    const todoService = new ToDoService(graphClient);
    const calendarService = new CalendarService(graphClient);
    const plannerService = new PlannerService(graphClient);

    return new MsGraphMCPServer(todoService, calendarService, plannerService, tokenRefresher);
  }

  /**
   * Set up MCP protocol handlers.
   */
  private setupHandlers(): void {
    // List available tools
    this.server.setRequestHandler(ListToolsRequestSchema, async () => {
      logger.debug('Handling list_tools request');
      return {
        tools: this.getTools(),
      };
    });

    // Handle tool calls
    this.server.setRequestHandler(CallToolRequestSchema, async (request) => {
      const context = this.createRequestContext(request.params.name);
      logger.info('Tool called', {
        requestId: context.requestId,
        tool: context.tool,
      });

      try {
        const result = await this.handleToolCall(request.params.name, request.params.arguments);

        logger.info('Tool completed successfully', {
          requestId: context.requestId,
          tool: context.tool,
        });

        return result;
      } catch (error) {
        logger.error('Tool execution failed', {
          requestId: context.requestId,
          tool: context.tool,
          error: error instanceof Error ? error.message : 'Unknown error',
        });

        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(sanitizeError(error), null, 2),
            },
          ],
        };
      }
    });
  }

  /**
   * Get list of available tools.
   */
  private getTools(): Tool[] {
    return [
      {
        name: 'get_auth_status',
        description: 'Check if user is authenticated with Microsoft Graph. Returns authentication status without exposing tokens.',
        inputSchema: {
          type: 'object',
          properties: {},
        },
      },
      {
        name: 'list_task_lists',
        description: 'Get all task lists for the authenticated user.',
        inputSchema: {
          type: 'object',
          properties: {},
        },
      },
      {
        name: 'list_tasks',
        description: 'Get tasks from a list with optional filtering. Supports filters: all (default), completed, incomplete, high-priority, today, overdue, this-week, later.',
        inputSchema: {
          type: 'object',
          properties: {
            listId: {
              type: 'string',
              description: 'Task list ID (optional, uses default list if not provided)',
            },
            filter: {
              type: 'string',
              enum: ['all', 'completed', 'incomplete', 'high-priority', 'today', 'overdue', 'this-week', 'later'],
              description: 'Filter tasks by status, priority, or due date',
            },
            limit: {
              type: 'number',
              description: 'Maximum number of tasks to return (1-100)',
              minimum: 1,
              maximum: 100,
            },
          },
        },
      },
      {
        name: 'get_task',
        description: 'Get a specific task by ID.',
        inputSchema: {
          type: 'object',
          properties: {
            listId: {
              type: 'string',
              description: 'Task list ID',
            },
            taskId: {
              type: 'string',
              description: 'Task ID',
            },
          },
          required: ['listId', 'taskId'],
        },
      },
      {
        name: 'search_tasks',
        description: 'Search for tasks across all lists by title.',
        inputSchema: {
          type: 'object',
          properties: {
            query: {
              type: 'string',
              description: 'Search query (1-100 characters)',
              minLength: 1,
              maxLength: 100,
            },
            limit: {
              type: 'number',
              description: 'Maximum number of results (1-50)',
              minimum: 1,
              maximum: 50,
            },
          },
          required: ['query'],
        },
      },
      // Calendar tools
      {
        name: 'list_calendar_events',
        description: 'Get calendar events with optional time range filter. Defaults to today.',
        inputSchema: {
          type: 'object',
          properties: {
            filter: {
              type: 'string',
              enum: ['today', 'this-week', 'this-month'],
              description: 'Filter events by time range (default: today)',
            },
            limit: {
              type: 'number',
              description: 'Maximum number of events to return (1-100)',
              minimum: 1,
              maximum: 100,
            },
          },
        },
      },
      {
        name: 'get_calendar_view',
        description: 'Get calendar events in a custom date range.',
        inputSchema: {
          type: 'object',
          properties: {
            startDate: {
              type: 'string',
              description: 'Start date in ISO format (e.g., 2025-01-15)',
            },
            endDate: {
              type: 'string',
              description: 'End date in ISO format (e.g., 2025-01-20)',
            },
            limit: {
              type: 'number',
              description: 'Maximum number of events to return (1-100)',
              minimum: 1,
              maximum: 100,
            },
          },
          required: ['startDate', 'endDate'],
        },
      },
      // ============================================
      // Planner read tools
      // ============================================
      {
        name: 'list_my_planner_tasks',
        description: 'Get all Planner tasks assigned to the authenticated user.',
        inputSchema: {
          type: 'object',
          properties: {
            filter: {
              type: 'string',
              enum: ['all', 'incomplete', 'completed'],
              description: 'Filter tasks by completion status',
            },
            limit: {
              type: 'number',
              description: 'Maximum number of tasks to return (1-100)',
              minimum: 1,
              maximum: 100,
            },
          },
        },
      },
      {
        name: 'get_planner_plan',
        description: 'Get a specific Planner plan by ID.',
        inputSchema: {
          type: 'object',
          properties: {
            planId: {
              type: 'string',
              description: 'Planner plan ID',
            },
          },
          required: ['planId'],
        },
      },
      {
        name: 'list_plan_tasks',
        description: 'Get all tasks in a Planner plan.',
        inputSchema: {
          type: 'object',
          properties: {
            planId: {
              type: 'string',
              description: 'Planner plan ID',
            },
            bucketId: {
              type: 'string',
              description: 'Filter by bucket ID',
            },
            limit: {
              type: 'number',
              description: 'Maximum number of tasks to return (1-100)',
              minimum: 1,
              maximum: 100,
            },
          },
          required: ['planId'],
        },
      },
      {
        name: 'list_plan_buckets',
        description: 'Get all buckets (columns) in a Planner plan.',
        inputSchema: {
          type: 'object',
          properties: {
            planId: {
              type: 'string',
              description: 'Planner plan ID',
            },
          },
          required: ['planId'],
        },
      },
      {
        name: 'get_planner_task',
        description: 'Get a specific Planner task by ID with optional details.',
        inputSchema: {
          type: 'object',
          properties: {
            taskId: {
              type: 'string',
              description: 'Planner task ID',
            },
            includeDetails: {
              type: 'boolean',
              description: 'Include description and checklist',
            },
          },
          required: ['taskId'],
        },
      },
      // ============================================
      // To Do write tools
      // ============================================
      {
        name: 'create_task',
        description: 'Create a new task in Microsoft To Do.',
        inputSchema: {
          type: 'object',
          properties: {
            listId: {
              type: 'string',
              description: 'Task list ID (uses default list if not provided)',
            },
            title: {
              type: 'string',
              description: 'Task title (max 400 characters)',
              maxLength: 400,
            },
            dueDateTime: {
              type: 'string',
              description: 'Due date in ISO format (e.g., 2025-01-20)',
            },
            importance: {
              type: 'string',
              enum: ['low', 'normal', 'high'],
              description: 'Task importance',
            },
            body: {
              type: 'string',
              description: 'Task body/description',
            },
          },
          required: ['title'],
        },
      },
      {
        name: 'update_task',
        description: 'Update an existing task in Microsoft To Do.',
        inputSchema: {
          type: 'object',
          properties: {
            listId: {
              type: 'string',
              description: 'Task list ID',
            },
            taskId: {
              type: 'string',
              description: 'Task ID',
            },
            title: {
              type: 'string',
              description: 'New title',
            },
            dueDateTime: {
              type: ['string', 'null'],
              description: 'New due date (null to clear)',
            },
            importance: {
              type: 'string',
              enum: ['low', 'normal', 'high'],
              description: 'New importance',
            },
            status: {
              type: 'string',
              enum: ['notStarted', 'inProgress', 'completed'],
              description: 'New status',
            },
            body: {
              type: 'string',
              description: 'New body/description',
            },
          },
          required: ['listId', 'taskId'],
        },
      },
      {
        name: 'complete_task',
        description: 'Mark a To Do task as completed.',
        inputSchema: {
          type: 'object',
          properties: {
            listId: {
              type: 'string',
              description: 'Task list ID',
            },
            taskId: {
              type: 'string',
              description: 'Task ID',
            },
          },
          required: ['listId', 'taskId'],
        },
      },
      {
        name: 'uncomplete_task',
        description: 'Mark a To Do task as not started.',
        inputSchema: {
          type: 'object',
          properties: {
            listId: {
              type: 'string',
              description: 'Task list ID',
            },
            taskId: {
              type: 'string',
              description: 'Task ID',
            },
          },
          required: ['listId', 'taskId'],
        },
      },
      {
        name: 'delete_task',
        description: 'Delete a task from Microsoft To Do. Requires confirmation.',
        inputSchema: {
          type: 'object',
          properties: {
            listId: {
              type: 'string',
              description: 'Task list ID',
            },
            taskId: {
              type: 'string',
              description: 'Task ID',
            },
            confirm: {
              type: 'boolean',
              const: true,
              description: 'Must be true to confirm deletion',
            },
          },
          required: ['listId', 'taskId', 'confirm'],
        },
      },
      {
        name: 'create_task_list',
        description: 'Create a new task list in Microsoft To Do.',
        inputSchema: {
          type: 'object',
          properties: {
            displayName: {
              type: 'string',
              description: 'List name (max 400 characters)',
              maxLength: 400,
            },
          },
          required: ['displayName'],
        },
      },
      {
        name: 'delete_task_list',
        description: 'Delete a task list from Microsoft To Do. Requires confirmation.',
        inputSchema: {
          type: 'object',
          properties: {
            listId: {
              type: 'string',
              description: 'Task list ID',
            },
            confirm: {
              type: 'boolean',
              const: true,
              description: 'Must be true to confirm deletion',
            },
          },
          required: ['listId', 'confirm'],
        },
      },
      // ============================================
      // Planner write tools
      // ============================================
      {
        name: 'create_planner_task',
        description: 'Create a new task in Microsoft Planner.',
        inputSchema: {
          type: 'object',
          properties: {
            planId: {
              type: 'string',
              description: 'Planner plan ID',
            },
            bucketId: {
              type: 'string',
              description: 'Bucket ID',
            },
            title: {
              type: 'string',
              description: 'Task title (max 400 characters)',
              maxLength: 400,
            },
            dueDateTime: {
              type: 'string',
              description: 'Due date in ISO format',
            },
            priority: {
              type: 'string',
              enum: ['urgent', 'high', 'normal', 'low'],
              description: 'Task priority',
            },
          },
          required: ['planId', 'bucketId', 'title'],
        },
      },
      {
        name: 'update_planner_task',
        description: 'Update an existing task in Microsoft Planner.',
        inputSchema: {
          type: 'object',
          properties: {
            taskId: {
              type: 'string',
              description: 'Planner task ID',
            },
            title: {
              type: 'string',
              description: 'New title',
            },
            dueDateTime: {
              type: ['string', 'null'],
              description: 'New due date (null to clear)',
            },
            priority: {
              type: 'string',
              enum: ['urgent', 'high', 'normal', 'low'],
              description: 'New priority',
            },
            percentComplete: {
              type: 'number',
              enum: [0, 50, 100],
              description: 'Completion percentage (0, 50, or 100)',
            },
            bucketId: {
              type: 'string',
              description: 'Move to different bucket',
            },
          },
          required: ['taskId'],
        },
      },
      {
        name: 'complete_planner_task',
        description: 'Mark a Planner task as completed (100%).',
        inputSchema: {
          type: 'object',
          properties: {
            taskId: {
              type: 'string',
              description: 'Planner task ID',
            },
          },
          required: ['taskId'],
        },
      },
      {
        name: 'uncomplete_planner_task',
        description: 'Mark a Planner task as not started (0%).',
        inputSchema: {
          type: 'object',
          properties: {
            taskId: {
              type: 'string',
              description: 'Planner task ID',
            },
          },
          required: ['taskId'],
        },
      },
      {
        name: 'delete_planner_task',
        description: 'Delete a task from Microsoft Planner. Requires confirmation.',
        inputSchema: {
          type: 'object',
          properties: {
            taskId: {
              type: 'string',
              description: 'Planner task ID',
            },
            confirm: {
              type: 'boolean',
              const: true,
              description: 'Must be true to confirm deletion',
            },
          },
          required: ['taskId', 'confirm'],
        },
      },
    ];
  }

  /**
   * Handle tool call execution.
   */
  private async handleToolCall(name: string, args: any) {
    switch (name) {
      case 'get_auth_status':
        return this.handleGetAuthStatus(args);

      case 'list_task_lists':
        return this.handleListTaskLists(args);

      case 'list_tasks':
        return this.handleListTasks(args);

      case 'get_task':
        return this.handleGetTask(args);

      case 'search_tasks':
        return this.handleSearchTasks(args);

      // Calendar tools
      case 'list_calendar_events':
        return this.handleListCalendarEvents(args);

      case 'get_calendar_view':
        return this.handleGetCalendarView(args);

      // Planner read tools
      case 'list_my_planner_tasks':
        return this.handleListMyPlannerTasks(args);

      case 'get_planner_plan':
        return this.handleGetPlannerPlan(args);

      case 'list_plan_tasks':
        return this.handleListPlanTasks(args);

      case 'list_plan_buckets':
        return this.handleListPlanBuckets(args);

      case 'get_planner_task':
        return this.handleGetPlannerTask(args);

      // To Do write tools
      case 'create_task':
        return this.handleCreateTask(args);

      case 'update_task':
        return this.handleUpdateTask(args);

      case 'complete_task':
        return this.handleCompleteTask(args);

      case 'uncomplete_task':
        return this.handleUncompleteTask(args);

      case 'delete_task':
        return this.handleDeleteTask(args);

      case 'create_task_list':
        return this.handleCreateTaskList(args);

      case 'delete_task_list':
        return this.handleDeleteTaskList(args);

      // Planner write tools
      case 'create_planner_task':
        return this.handleCreatePlannerTask(args);

      case 'update_planner_task':
        return this.handleUpdatePlannerTask(args);

      case 'complete_planner_task':
        return this.handleCompletePlannerTask(args);

      case 'uncomplete_planner_task':
        return this.handleUncompletePlannerTask(args);

      case 'delete_planner_task':
        return this.handleDeletePlannerTask(args);

      default:
        throw new Error(`Unknown tool: ${name}`);
    }
  }

  /**
   * Handle get_auth_status tool.
   */
  private async handleGetAuthStatus(args: unknown) {
    // Validate input (no fields required, but validates structure)
    GetAuthStatusSchema.parse(args);

    let authStatus: AuthStatus;
    try {
      // Try to get a valid access token
      await this.tokenRefresher.getValidAccessToken();
      authStatus = {
        authenticated: true,
        message: 'Authenticated with Microsoft Graph',
      };
    } catch {
      authStatus = {
        authenticated: false,
        message: 'Not authenticated. Please run OAuth flow first.',
      };
    }

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(authStatus, null, 2),
        },
      ],
    };
  }

  /**
   * Handle list_task_lists tool.
   */
  private async handleListTaskLists(args: unknown) {
    // Validate input (no fields required, but validates structure)
    ListTaskListsSchema.parse(args);

    const lists = await this.todoService.getTaskLists();
    const sanitized = sanitizeTaskLists(lists);

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(sanitized, null, 2),
        },
      ],
    };
  }

  /**
   * Handle list_tasks tool.
   */
  private async handleListTasks(args: unknown) {
    const validated = ListTasksSchema.parse(args);

    let tasks;
    const limit = validated.limit ?? 50;

    // Apply filter
    switch (validated.filter) {
      case 'completed':
        tasks = await this.todoService.getCompletedTasks(validated.listId, limit);
        break;

      case 'incomplete':
        tasks = await this.todoService.getTasks(validated.listId, {
          filter: "status ne 'completed'",
          top: limit,
        });
        break;

      case 'high-priority':
        tasks = await this.todoService.getHighPriorityTasks(validated.listId, limit);
        break;

      case 'today':
        tasks = await this.todoService.getTasksDueToday(validated.listId);
        // Apply limit after fetching
        tasks = tasks.slice(0, limit);
        break;

      case 'overdue':
        tasks = await this.todoService.getTasksDueOverdue(validated.listId);
        // Apply limit after fetching
        tasks = tasks.slice(0, limit);
        break;

      case 'this-week':
        tasks = await this.todoService.getTasksDueThisWeek(validated.listId);
        // Apply limit after fetching
        tasks = tasks.slice(0, limit);
        break;

      case 'later':
        tasks = await this.todoService.getTasksDueLater(validated.listId);
        // Apply limit after fetching
        tasks = tasks.slice(0, limit);
        break;

      case 'all':
      default:
        tasks = await this.todoService.getTasks(validated.listId, {
          top: limit,
        });
        break;
    }

    const sanitized = sanitizeTasks(tasks);

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(sanitized, null, 2),
        },
      ],
    };
  }

  /**
   * Handle get_task tool.
   */
  private async handleGetTask(args: unknown) {
    const validated = GetTaskSchema.parse(args);

    const task = await this.todoService.getTask(validated.listId, validated.taskId);
    const sanitized = sanitizeTask(task);

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(sanitized, null, 2),
        },
      ],
    };
  }

  /**
   * Handle search_tasks tool.
   */
  private async handleSearchTasks(args: unknown) {
    const validated = SearchTasksSchema.parse(args);

    const limit = validated.limit ?? 20;
    const tasks = await this.todoService.searchTasks(validated.query, {
      top: limit,
    });

    const sanitized = sanitizeTasks(tasks);

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(sanitized, null, 2),
        },
      ],
    };
  }

  // ============================================
  // Calendar tool handlers
  // ============================================

  /**
   * Handle list_calendar_events tool.
   */
  private async handleListCalendarEvents(args: unknown) {
    const validated = ListCalendarEventsSchema.parse(args);

    let events;
    const limit = validated.limit ?? 50;

    // Apply filter
    switch (validated.filter) {
      case 'this-week':
        events = await this.calendarService.getEventsThisWeek();
        break;

      case 'this-month':
        events = await this.calendarService.getEventsThisMonth();
        break;

      case 'today':
      default:
        events = await this.calendarService.getEventsToday();
        break;
    }

    // Apply limit
    events = events.slice(0, limit);

    const sanitized = sanitizeCalendarEvents(events);

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(sanitized, null, 2),
        },
      ],
    };
  }

  /**
   * Handle get_calendar_view tool.
   */
  private async handleGetCalendarView(args: unknown) {
    const validated = GetCalendarViewSchema.parse(args);

    // Validate and parse dates
    const startDate = validateDateString(validated.startDate, 'startDate');
    const endDate = validateDateString(validated.endDate, 'endDate');

    // Validate the date range
    validateDateRange(startDate, endDate);

    const limit = validated.limit ?? 100;

    // Get events in the date range
    let events = await this.calendarService.getCalendarView(startDate, endDate);

    // Apply limit
    events = events.slice(0, limit);

    const sanitized = sanitizeCalendarEvents(events);

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(sanitized, null, 2),
        },
      ],
    };
  }

  // ============================================
  // Planner read tool handlers
  // ============================================

  /**
   * Handle list_my_planner_tasks tool.
   */
  private async handleListMyPlannerTasks(args: unknown) {
    const validated = ListMyPlannerTasksSchema.parse(args);

    let tasks = await this.plannerService.getMyTasks({ top: validated.limit });

    // Apply filter
    if (validated.filter === 'completed') {
      tasks = tasks.filter(t => t.percentComplete === 100);
    } else if (validated.filter === 'incomplete') {
      tasks = tasks.filter(t => t.percentComplete < 100);
    }

    const sanitized = sanitizePlannerTasks(tasks);

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(sanitized, null, 2),
        },
      ],
    };
  }

  /**
   * Handle get_planner_plan tool.
   */
  private async handleGetPlannerPlan(args: unknown) {
    const validated = GetPlannerPlanSchema.parse(args);

    const plan = await this.plannerService.getPlan(validated.planId);
    const sanitized = sanitizePlannerPlan(plan);

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(sanitized, null, 2),
        },
      ],
    };
  }

  /**
   * Handle list_plan_tasks tool.
   */
  private async handleListPlanTasks(args: unknown) {
    const validated = ListPlanTasksSchema.parse(args);

    const tasks = await this.plannerService.getPlanTasks(validated.planId, {
      bucketId: validated.bucketId,
      top: validated.limit,
    });
    const sanitized = sanitizePlannerTasks(tasks);

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(sanitized, null, 2),
        },
      ],
    };
  }

  /**
   * Handle list_plan_buckets tool.
   */
  private async handleListPlanBuckets(args: unknown) {
    const validated = ListPlanBucketsSchema.parse(args);

    const buckets = await this.plannerService.getPlanBuckets(validated.planId);
    const sanitized = sanitizePlannerBuckets(buckets);

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(sanitized, null, 2),
        },
      ],
    };
  }

  /**
   * Handle get_planner_task tool.
   */
  private async handleGetPlannerTask(args: unknown) {
    const validated = GetPlannerTaskSchema.parse(args);

    const task = await this.plannerService.getTask(validated.taskId);
    const sanitized = sanitizePlannerTask(task);

    let result: any = sanitized;

    // Include details if requested
    if (validated.includeDetails) {
      const details = await this.plannerService.getTaskDetails(validated.taskId);
      result = {
        ...sanitized,
        details: sanitizePlannerTaskDetails(details),
      };
    }

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(result, null, 2),
        },
      ],
    };
  }

  // ============================================
  // To Do write tool handlers
  // ============================================

  /**
   * Handle create_task tool.
   */
  private async handleCreateTask(args: unknown) {
    const validated = CreateTaskSchema.parse(args);

    const task = await this.todoService.createTask(validated.listId, {
      title: validated.title,
      dueDateTime: validated.dueDateTime,
      importance: validated.importance,
      body: validated.body,
    });
    const sanitized = sanitizeTask(task);

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({ message: 'Task created successfully', task: sanitized }, null, 2),
        },
      ],
    };
  }

  /**
   * Handle update_task tool.
   */
  private async handleUpdateTask(args: unknown) {
    const validated = UpdateTaskSchema.parse(args);

    const task = await this.todoService.updateTask(validated.listId, validated.taskId, {
      title: validated.title,
      dueDateTime: validated.dueDateTime,
      importance: validated.importance,
      status: validated.status,
      body: validated.body,
    });
    const sanitized = sanitizeTask(task);

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({ message: 'Task updated successfully', task: sanitized }, null, 2),
        },
      ],
    };
  }

  /**
   * Handle complete_task tool.
   */
  private async handleCompleteTask(args: unknown) {
    const validated = CompleteTaskSchema.parse(args);

    const task = await this.todoService.completeTask(validated.listId, validated.taskId);
    const sanitized = sanitizeTask(task);

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({ message: 'Task marked as completed', task: sanitized }, null, 2),
        },
      ],
    };
  }

  /**
   * Handle uncomplete_task tool.
   */
  private async handleUncompleteTask(args: unknown) {
    const validated = CompleteTaskSchema.parse(args);

    const task = await this.todoService.uncompleteTask(validated.listId, validated.taskId);
    const sanitized = sanitizeTask(task);

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({ message: 'Task marked as not started', task: sanitized }, null, 2),
        },
      ],
    };
  }

  /**
   * Handle delete_task tool.
   */
  private async handleDeleteTask(args: unknown) {
    const validated = DeleteTaskSchema.parse(args);

    await this.todoService.deleteTask(validated.listId, validated.taskId);

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({ message: 'Task deleted successfully' }, null, 2),
        },
      ],
    };
  }

  /**
   * Handle create_task_list tool.
   */
  private async handleCreateTaskList(args: unknown) {
    const validated = CreateTaskListSchema.parse(args);

    const list = await this.todoService.createTaskList(validated.displayName);
    const sanitized = {
      id: list.id,
      displayName: list.displayName,
    };

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({ message: 'Task list created successfully', list: sanitized }, null, 2),
        },
      ],
    };
  }

  /**
   * Handle delete_task_list tool.
   */
  private async handleDeleteTaskList(args: unknown) {
    const validated = DeleteTaskListSchema.parse(args);

    await this.todoService.deleteTaskList(validated.listId);

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({ message: 'Task list deleted successfully' }, null, 2),
        },
      ],
    };
  }

  // ============================================
  // Planner write tool handlers
  // ============================================

  /**
   * Handle create_planner_task tool.
   */
  private async handleCreatePlannerTask(args: unknown) {
    const validated = CreatePlannerTaskSchema.parse(args);

    // Convert priority string to number
    const priorityMap: Record<string, number> = {
      urgent: 1,
      high: 3,
      normal: 5,
      low: 9,
    };

    const task = await this.plannerService.createTask(validated.planId, validated.bucketId, {
      title: validated.title,
      dueDateTime: validated.dueDateTime,
      priority: validated.priority ? priorityMap[validated.priority] : undefined,
    });
    const sanitized = sanitizePlannerTask(task);

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({ message: 'Planner task created successfully', task: sanitized }, null, 2),
        },
      ],
    };
  }

  /**
   * Handle update_planner_task tool.
   */
  private async handleUpdatePlannerTask(args: unknown) {
    const validated = UpdatePlannerTaskSchema.parse(args);

    // First, get the current task to obtain its etag
    const currentTask = await this.plannerService.getTask(validated.taskId);
    const etag = currentTask['@odata.etag'];

    if (!etag) {
      throw new Error('Unable to update task: etag not found');
    }

    // Convert priority string to number
    const priorityMap: Record<string, number> = {
      urgent: 1,
      high: 3,
      normal: 5,
      low: 9,
    };

    const task = await this.plannerService.updateTask(validated.taskId, etag, {
      title: validated.title,
      dueDateTime: validated.dueDateTime,
      priority: validated.priority ? priorityMap[validated.priority] : undefined,
      percentComplete: validated.percentComplete,
      bucketId: validated.bucketId,
    });
    const sanitized = sanitizePlannerTask(task);

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({ message: 'Planner task updated successfully', task: sanitized }, null, 2),
        },
      ],
    };
  }

  /**
   * Handle complete_planner_task tool.
   */
  private async handleCompletePlannerTask(args: unknown) {
    const validated = CompletePlannerTaskSchema.parse(args);

    // First, get the current task to obtain its etag
    const currentTask = await this.plannerService.getTask(validated.taskId);
    const etag = currentTask['@odata.etag'];

    if (!etag) {
      throw new Error('Unable to complete task: etag not found');
    }

    const task = await this.plannerService.completeTask(validated.taskId, etag);
    const sanitized = sanitizePlannerTask(task);

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({ message: 'Planner task marked as completed', task: sanitized }, null, 2),
        },
      ],
    };
  }

  /**
   * Handle uncomplete_planner_task tool.
   */
  private async handleUncompletePlannerTask(args: unknown) {
    const validated = CompletePlannerTaskSchema.parse(args);

    // First, get the current task to obtain its etag
    const currentTask = await this.plannerService.getTask(validated.taskId);
    const etag = currentTask['@odata.etag'];

    if (!etag) {
      throw new Error('Unable to uncomplete task: etag not found');
    }

    const task = await this.plannerService.uncompleteTask(validated.taskId, etag);
    const sanitized = sanitizePlannerTask(task);

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({ message: 'Planner task marked as not started', task: sanitized }, null, 2),
        },
      ],
    };
  }

  /**
   * Handle delete_planner_task tool.
   */
  private async handleDeletePlannerTask(args: unknown) {
    const validated = DeletePlannerTaskSchema.parse(args);

    // First, get the current task to obtain its etag
    const currentTask = await this.plannerService.getTask(validated.taskId);
    const etag = currentTask['@odata.etag'];

    if (!etag) {
      throw new Error('Unable to delete task: etag not found');
    }

    await this.plannerService.deleteTask(validated.taskId, etag);

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({ message: 'Planner task deleted successfully' }, null, 2),
        },
      ],
    };
  }

  /**
   * Create request context for logging and tracking.
   */
  private createRequestContext(tool: string): RequestContext {
    return {
      requestId: randomUUID(),
      tool,
      timestamp: Date.now(),
    };
  }

  /**
   * Start the MCP server.
   */
  async start(): Promise<void> {
    logger.info('Starting MsGraph MCP Server');

    const transport = new StdioServerTransport();
    await this.server.connect(transport);

    logger.info('MsGraph MCP Server started and listening on stdio');
  }
}
