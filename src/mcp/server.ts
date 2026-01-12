/**
 * MCP Server implementation for Microsoft Graph API.
 * Provides read-only access to To Do tasks and Calendar events via secure, validated tools.
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
import { sanitizeError, sanitizeTasks, sanitizeTaskLists, sanitizeTask } from '../security/sanitizers.js';
import { ToDoService } from '../graph/todo-service.js';
import { SecureGraphClient } from '../graph/client.js';
import { TokenRefresher } from '../auth/token-refresher.js';
import { getTokenManager } from '../auth/token-manager-factory.js';
import {
  GetAuthStatusSchema,
  ListTaskListsSchema,
  ListTasksSchema,
  GetTaskSchema,
  SearchTasksSchema,
  RequestContext,
  AuthStatus,
} from './types.js';

/**
 * MsGraph MCP Server - Secure Microsoft Graph API integration.
 */
export class MsGraphMCPServer {
  private server: Server;
  private todoService: ToDoService;
  private tokenRefresher: TokenRefresher;

  private constructor(todoService: ToDoService, tokenRefresher: TokenRefresher) {
    logger.info('Initializing MsGraph MCP Server');

    this.todoService = todoService;
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

    return new MsGraphMCPServer(todoService, tokenRefresher);
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
