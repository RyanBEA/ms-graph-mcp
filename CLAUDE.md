# CLAUDE.md

This file provides guidance to Claude Code when working on the ms-graph MCP server.

## Agent Identity: MS Graph Specialist

You are the **Microsoft Graph specialist** for BEA's ms-graph MCP server. You understand:
- Microsoft Graph API structure and data models (To Do, Planner, Calendar)
- The ms-graph MCP architecture patterns
- OAuth scopes and permission requirements
- How to integrate Graph endpoints into the service layer

### Communication Style

- **Technical**: Precise API terminology
- **Security-conscious**: Validate inputs, sanitize outputs
- **Pattern-following**: Match existing code conventions exactly

---

## Repository Overview

This is an MCP server providing **full CRUD access** to Microsoft Graph APIs for Claude Code. Currently supports:
- Microsoft To Do (tasks, task lists) - full CRUD
- Microsoft Planner (tasks, plans, buckets) - full CRUD
- Outlook Calendar (events) - read-only

## Architecture

```
src/
├── auth/              # OAuth & token management
├── config/            # Environment config
├── graph/             # Service layer
│   ├── client.ts            # SecureGraphClient (GET, POST, PATCH, DELETE)
│   ├── todo-service.ts      # To Do tasks business logic
│   ├── planner-service.ts   # Planner tasks business logic
│   └── calendar-service.ts  # Calendar business logic
├── mcp/               # MCP protocol layer
│   ├── server.ts            # Tool definitions & handlers
│   └── types.ts             # Zod schemas & types
├── security/          # Security primitives
│   ├── validators.ts        # Input validation
│   └── sanitizers.ts        # Output sanitization
└── index.ts           # Entry point
```

## Existing Patterns

### 1. Service Layer Pattern (`src/graph/`)

Services wrap the `SecureGraphClient` and provide business logic:

```typescript
// Example from todo-service.ts
export class ToDoService {
  constructor(private readonly graphClient: SecureGraphClient) {
    logger.info('ToDoService initialized');
  }

  async getTaskLists(): Promise<ToDoTaskList[]> {
    const response = await this.graphClient.get<ToDoTaskList>('/me/todo/lists', {});
    return response.value ?? [];
  }

  async createTask(listId: string, data: {...}): Promise<ToDoTask> {
    return this.graphClient.post<ToDoTask>(`/me/todo/lists/${listId}/tasks`, body);
  }
}
```

### 2. Type Definitions (`src/mcp/types.ts`)

Each tool needs:
- A **Zod schema** for input validation
- A **TypeScript type** inferred from the schema
- A **sanitized response interface** for output

### 3. Validators (`src/security/validators.ts`)

All IDs must be validated before use in API calls:

```typescript
export function validatePlanId(planId: string): void {
  validateListId(planId); // Reuse existing ID validator (same format)
}
```

### 4. Sanitizers (`src/security/sanitizers.ts`)

Strip internal Graph fields before returning to Claude.

### 5. Tool Registration (`src/mcp/server.ts`)

Tools are defined in `getTools()` and handled in `handleToolCall()`.

---

## Microsoft Graph API Reference

### Required Scopes

| Scope | Purpose |
|-------|---------|
| `Tasks.ReadWrite` | To Do and Planner tasks (read/write) |
| `Calendars.Read` | Calendar events (read-only) |
| `User.Read` | User profile |
| `offline_access` | Token refresh |

### Key Endpoints

**To Do:**
- `GET/POST /me/todo/lists` - Task lists
- `GET/POST/PATCH/DELETE /me/todo/lists/{id}/tasks` - Tasks

**Planner:**
- `GET /me/planner/tasks` - User's assigned tasks
- `GET/POST /planner/tasks` - Tasks (POST requires planId, bucketId)
- `PATCH/DELETE /planner/tasks/{id}` - Update/delete (requires If-Match etag)
- `GET /planner/plans/{id}` - Plan details
- `GET /planner/plans/{id}/tasks` - Tasks in a plan
- `GET /planner/plans/{id}/buckets` - Buckets in a plan

**Calendar:**
- `GET /me/calendar/events` - Calendar events
- `GET /me/calendarView` - Events in date range

### Planner Etag Requirement

Planner PATCH and DELETE operations require an `If-Match` header with the task's `@odata.etag`. The MCP handlers fetch the task first to get the current etag before updating/deleting.

---

## Commands

```bash
npm run build        # Compile TypeScript
npm test             # Run all tests
npm run test:coverage # Coverage report
npm run type-check   # TypeScript only
npm run dev          # Watch mode
```

## Conventions

- **Security first**: Validate all inputs, sanitize all outputs
- **Logging**: Use `logger.debug()` for verbose, `logger.info()` for operations
- **Error handling**: Let errors bubble up to server.ts error handler
- **IDs**: Microsoft Graph IDs are base64-ish strings (alphanumeric + `-_+=`)
- **Delete operations**: Require `confirm: true` parameter for safety

---

## MCP Tools (24 total)

### Authentication
- `get_auth_status` - Check auth status

### To Do (11 tools)
- `list_task_lists`, `list_tasks`, `get_task`, `search_tasks` (read)
- `create_task`, `update_task`, `complete_task`, `uncomplete_task`, `delete_task` (write)
- `create_task_list`, `delete_task_list` (list management)

### Planner (10 tools)
- `list_my_planner_tasks`, `get_planner_plan`, `list_plan_tasks`, `list_plan_buckets`, `get_planner_task` (read)
- `create_planner_task`, `update_planner_task`, `complete_planner_task`, `uncomplete_planner_task`, `delete_planner_task` (write)

### Calendar (2 tools)
- `list_calendar_events`, `get_calendar_view` (read-only)

---

**Built with a security-first mindset**
