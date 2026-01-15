# MS-Graph MCP Server

A **security-first** Model Context Protocol (MCP) server providing full access to Microsoft To Do, Planner, and Calendar for Claude Code and Claude Desktop.

## Features

- **Security-First Design** - CSRF protection, input validation, automatic token redaction
- **Full CRUD Support** - Create, read, update, and delete tasks in To Do and Planner
- **Microsoft Planner Integration** - Manage Planner tasks, plans, and buckets
- **Secure Token Storage** - MSAL cache with automatic token refresh
- **Rate Limited** - Token bucket algorithm (60 req/min) with circuit breaker
- **Well Tested** - 160+ tests passing
- **Comprehensive Logging** - Winston with automatic sensitive data redaction

## MCP Tools Available

The MS-Graph MCP server provides 24 tools:

### Authentication
1. **`get_auth_status`** - Check authentication status (no sensitive data exposed)

### To Do - Read Tools
2. **`list_task_lists`** - Get all task lists for the authenticated user
3. **`list_tasks`** - Get tasks with filtering:
   - `all` - All tasks (default)
   - `completed` - Only completed tasks
   - `incomplete` - Only incomplete tasks
   - `high-priority` - High-priority incomplete tasks
   - `today` - Tasks due today
   - `overdue` - Tasks with due dates in the past
   - `this-week` - Tasks due in the next 7 days
   - `later` - Tasks due beyond 7 days or with no due date
4. **`get_task`** - Get a specific task by list ID and task ID
5. **`search_tasks`** - Search tasks across all lists (1-100 char query, 1-50 result limit)

### To Do - Write Tools
6. **`create_task`** - Create a new task in a task list
7. **`update_task`** - Update task title, due date, or body content
8. **`complete_task`** - Mark a task as completed
9. **`uncomplete_task`** - Mark a task as not started
10. **`delete_task`** - Delete a task (requires `confirm: true`)
11. **`create_task_list`** - Create a new task list
12. **`delete_task_list`** - Delete a task list (requires `confirm: true`)

### Planner - Read Tools
13. **`list_my_planner_tasks`** - Get all Planner tasks assigned to the user
14. **`get_planner_plan`** - Get a specific Planner plan by ID
15. **`list_plan_tasks`** - Get all tasks in a Planner plan
16. **`list_plan_buckets`** - Get all buckets (columns) in a Planner plan
17. **`get_planner_task`** - Get a specific Planner task with details

### Planner - Write Tools
18. **`create_planner_task`** - Create a new task in a Planner plan/bucket
19. **`update_planner_task`** - Update task title, due date, or priority
20. **`complete_planner_task`** - Mark a Planner task as 100% complete
21. **`uncomplete_planner_task`** - Mark a Planner task as 0% complete
22. **`delete_planner_task`** - Delete a Planner task (requires `confirm: true`)

### Calendar Tools
23. **`list_calendar_events`** - Get calendar events with time range filters:
   - `today` - Events today (default)
   - `this-week` - Events in the next 7 days
   - `this-month` - Events in the next 30 days
24. **`get_calendar_view`** - Get events in a custom date range

## Quick Start

### Step 1: Install and Build

```bash
cd ms-graph
npm install
npm run build
```

### Step 2: Azure App Registration

See **[AUTH.md](./AUTH.md)** for complete Azure Portal setup instructions.

**Quick summary:**
1. Create app registration in [Azure Portal](https://portal.azure.com)
2. Add redirect URI: `http://localhost:3000/callback` (Web platform)
3. Add API permissions: `Tasks.ReadWrite`, `Calendars.Read`, `User.Read`, `offline_access`
4. Create client secret and save the **VALUE** (not ID)
5. Set "Allow public client flows" to **No** (confidential client)

### Step 3: Configure Environment

Create `.env` file:

```env
AZURE_CLIENT_ID=your-application-client-id
AZURE_TENANT_ID=your-directory-tenant-id
AZURE_CLIENT_SECRET=your-client-secret-value
AZURE_REDIRECT_URI=http://localhost:3000/callback
TOKEN_STORAGE=msal
LOG_LEVEL=info
```

Or configure in `~/.claude.json` (see [AUTH.md](./AUTH.md#configuration) for details).

### Step 4: Authenticate

```bash
node simple-auth.mjs
```

Open the URL displayed in an **incognito/private browser window**, sign in with your Microsoft account, and grant permissions. The script saves tokens to MSAL cache files for automatic refresh.

### Step 5: Configure Claude Code

Add to `~/.claude.json`:

```json
{
  "mcpServers": {
    "ms-graph": {
      "type": "stdio",
      "command": "node",
      "args": ["C:/ai/mcp-servers/ms-graph/dist/index.js"],
      "env": {
        "AZURE_CLIENT_ID": "your-client-id",
        "AZURE_TENANT_ID": "your-tenant-id",
        "AZURE_CLIENT_SECRET": "your-secret-value",
        "AZURE_REDIRECT_URI": "http://localhost:3000/callback",
        "TOKEN_STORAGE": "msal",
        "LOG_LEVEL": "info"
      },
      "cwd": "C:/ai/mcp-servers/ms-graph"
    }
  }
}
```

### Step 6: Restart Claude Code

Restart Claude Code to load the MCP server.

## Authentication Guide

For detailed authentication instructions, troubleshooting, and security notes, see:

**[AUTH.md](./AUTH.md)** - Complete Authentication Guide

Topics covered:
- Azure Portal setup (step-by-step)
- OAuth authentication
- Token storage and refresh
- Troubleshooting common errors
- Re-authentication process

## Configuration

### Environment Variables

| Variable | Required | Default | Description |
|----------|----------|---------|-------------|
| `AZURE_CLIENT_ID` | Yes | - | Application (client) ID from Azure |
| `AZURE_TENANT_ID` | Yes | - | Directory (tenant) ID from Azure |
| `AZURE_CLIENT_SECRET` | Yes | - | Client secret VALUE from Azure |
| `AZURE_REDIRECT_URI` | No | `http://localhost:3000/callback` | OAuth callback URL |
| `TOKEN_STORAGE` | No | `msal` | Token storage backend (`msal` recommended for auto-refresh) |
| `LOG_LEVEL` | No | `info` | Logging level (`error`, `warn`, `info`, `debug`) |
| `RATE_LIMIT_PER_MINUTE` | No | `60` | Microsoft Graph API rate limit |

## Development

### Commands

```bash
npm run dev          # Development mode with auto-reload
npm test             # Run all tests
npm run test:coverage # Run with coverage report
npm run test:security # Security-specific tests only
npm run type-check   # TypeScript type checking
npm run lint         # Linting
npm run clean        # Clean build artifacts
npm run build        # Production build
npm start            # Start production server
```

### Project Structure

```
ms-graph/
├── src/
│   ├── auth/              # OAuth & token management
│   │   ├── msal-token-manager.ts    # MSAL cache with auto-refresh
│   │   ├── token-manager-factory.ts # Factory pattern
│   │   ├── token-refresher.ts       # Refresh delegation
│   │   └── secure-oauth-client.ts   # CSRF-protected OAuth
│   ├── config/            # Configuration management
│   ├── graph/             # Microsoft Graph API client
│   │   ├── todo-service.ts          # To Do tasks business logic
│   │   ├── planner-service.ts       # Planner tasks business logic
│   │   └── calendar-service.ts      # Calendar business logic
│   ├── mcp/               # MCP server implementation
│   └── security/          # Security primitives
├── simple-auth.mjs        # MSAL OAuth authentication script (recommended)
├── AUTH.md                # Authentication guide
├── SECURITY.md            # Security documentation
└── README.md              # This file
```

## Security

This project follows **security-first principles**:

### Defense in Depth

1. **Input Validation** - Whitelist-based validation for OData filters, IDs, search queries
2. **Output Sanitization** - Strips Microsoft Graph internal fields from responses
3. **Logging Security** - Automatic redaction of tokens and secrets
4. **Rate Limiting** - Token bucket algorithm prevents API abuse
5. **Circuit Breaker** - Fail-fast pattern prevents cascading failures
6. **OAuth Security** - CSRF protection with cryptographic state tokens
7. **Token Management** - MSAL cache with automatic refresh (never in logs)

### Task Management Access

This server supports full task management for To Do and Planner:
- **To Do**: Create, read, update, complete, and delete tasks and task lists
- **Planner**: Create, read, update, complete, and delete Planner tasks
- **Calendar**: Read-only access to calendar events
- Cannot access mail, files, or other sensitive data

For detailed security information, see [SECURITY.md](./SECURITY.md).

## Troubleshooting

For authentication issues, see **[AUTH.md - Troubleshooting](./AUTH.md#troubleshooting)**.

### Common Issues

| Problem | Solution |
|---------|----------|
| "Not authenticated" | Run `node simple-auth.mjs` and restart Claude Code |
| "Invalid client secret" | Use secret VALUE (not ID); check for Windows env var override |
| "Client is public" error | Set "Allow public client flows" to **No** in Azure |
| Rate limit errors | Default is 60 req/min; check for loops in prompts |
| Build errors | Ensure Node.js v18+, run `npm run clean && npm install && npm run build` |
| MCP not loading | Verify `~/.claude.json` path and `cwd` setting |

## API Documentation

### Tool: `list_calendar_events`

**Description**: Get calendar events with optional time range filter

**Parameters**:
- `filter` (optional): `"today"` | `"this-week"` | `"this-month"` (default: `"today"`)
- `limit` (optional): Number of events (1-100, default: 50)

**Returns**:
```json
[
  {
    "id": "event-123",
    "subject": "Team Meeting",
    "start": "2025-01-12T09:00:00",
    "end": "2025-01-12T10:00:00",
    "location": "Conference Room A",
    "isAllDay": false,
    "organizer": "jane@example.com",
    "attendees": ["bob@example.com"]
  }
]
```

### Tool: `get_calendar_view`

**Description**: Get calendar events in a custom date range

**Parameters**:
- `startDate` (required): Start date in ISO format (e.g., `"2025-01-15"`)
- `endDate` (required): End date in ISO format (e.g., `"2025-01-20"`)
- `limit` (optional): Number of events (1-100, default: 50)

See [AUTH.md](./AUTH.md) for task tool documentation.

## License

ISC

## Acknowledgments

- **Anthropic** - For the Model Context Protocol SDK
- **Microsoft** - For the Graph API and MSAL library

---

**Built with a security-first mindset**
