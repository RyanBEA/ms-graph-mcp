# Toto MCP Server 🐕

A **security-first** Model Context Protocol (MCP) server providing read-only access to Microsoft To Do tasks for Claude Code and Claude Desktop.

Named after Dorothy's faithful companion in The Wizard of Oz - because every great journey needs a trustworthy helper!

## ✨ Features

- **🔒 Security-First Design** - CSRF protection, input validation, automatic token redaction
- **📖 Read-Only Access** - Safe integration with `Tasks.Read` and `User.Read` scopes
- **🔐 Secure Token Storage** - Supports keytar (Windows Credential Manager) and 1Password SDK
- **⚡ Rate Limited** - Token bucket algorithm (60 req/min) with circuit breaker
- **🧪 Well Tested** - 161 tests passing with 91.68% code coverage
- **📝 Comprehensive Logging** - Winston with automatic sensitive data redaction

## 🎯 MCP Tools Available

The Toto MCP server provides 5 read-only tools:

1. **`get_auth_status`** - Check authentication status (no sensitive data exposed)
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

   **Note**: When no `listId` is provided, date-based filters (`today`, `overdue`, `this-week`, `later`) search **all lists**.
4. **`get_task`** - Get a specific task by list ID and task ID
5. **`search_tasks`** - Search tasks across all lists (1-100 char query, 1-50 result limit)

## 📋 Prerequisites

- **Node.js** v18+ (v22.18.0 recommended)
- **npm** v9+ (v11.5.2 recommended)
- **Microsoft account** with Azure AD access
- **Azure app registration** (see setup below)
- **Windows** (for keytar/Credential Manager) or **1Password Teams** subscription

## 🚀 Quick Start

### Step 1: Install Dependencies

```bash
cd toto-mcp
npm install
```

### Step 2: Azure App Registration

You need to create an Azure AD app registration with the following configuration:

1. Go to [Azure Portal](https://portal.azure.com) → Azure Active Directory → App registrations
2. Click "New registration"
3. Name: `toto` (or any name you prefer)
4. Supported account types: "Accounts in any organizational directory and personal Microsoft accounts"
5. Redirect URI: `Web` - `http://localhost:3000/callback`
6. Click "Register"

After registration:
- Copy the **Application (client) ID**
- Copy the **Directory (tenant) ID**
- Go to "Certificates & secrets" → "New client secret" → Copy the **secret value**

**IMPORTANT**: Store these credentials securely! Never commit them to git.

### Step 3: Configure Environment Variables

The Toto MCP server needs Azure credentials to authenticate. You have two options:

#### Option A: Set System Environment Variables (Recommended for Claude Code)

Windows (PowerShell):
```powershell
[System.Environment]::SetEnvironmentVariable('AZURE_CLIENT_ID', 'your-client-id-here', 'User')
[System.Environment]::SetEnvironmentVariable('AZURE_TENANT_ID', 'your-tenant-id-here', 'User')
[System.Environment]::SetEnvironmentVariable('AZURE_CLIENT_SECRET', 'your-client-secret-here', 'User')
```

Windows (Command Prompt):
```cmd
setx AZURE_CLIENT_ID "your-client-id-here"
setx AZURE_TENANT_ID "your-tenant-id-here"
setx AZURE_CLIENT_SECRET "your-client-secret-here"
```

**Note**: After setting environment variables, restart Claude Code for changes to take effect.

#### Option B: Use .env File (Development Only)

```bash
cp .env.example .env
# Edit .env with your Azure credentials
```

**WARNING**: Never commit `.env` files! They are in `.gitignore` for security.

### Step 4: Build the Server

```bash
npm run build
```

This compiles TypeScript to JavaScript in the `dist/` directory.

### Step 5: Configure Claude Code

The Toto MCP server is already configured for Claude Code in `.claude/mcp.json`.

**Verify the configuration**:
```json
{
  "mcpServers": {
    "toto": {
      "command": "node",
      "args": ["C:\\path\\to\\toto-mcp\\dist\\index.js"],
      "env": {
        "AZURE_CLIENT_ID": "${env:AZURE_CLIENT_ID}",
        "AZURE_TENANT_ID": "${env:AZURE_TENANT_ID}",
        "AZURE_CLIENT_SECRET": "${env:AZURE_CLIENT_SECRET}",
        "TOKEN_STORAGE": "keytar",
        "LOG_LEVEL": "info"
      }
    }
  }
}
```

**Update the path** in `.claude/mcp.json` to match your installation location.

### Step 6: Restart Claude Code

After configuration, restart Claude Code to load the Toto MCP server.

### Step 7: Authenticate with Microsoft

On first use, you'll need to complete OAuth authentication:

1. Ask Claude Code to use a Toto tool (e.g., "list my task lists")
2. The server will return an authentication URL
3. Open the URL in your browser
4. Sign in with your Microsoft account
5. Consent to the requested permissions (`Tasks.Read`, `User.Read`)
6. The server will store your tokens securely in Windows Credential Manager

After initial authentication, tokens are automatically refreshed (valid for 90 days).

## 🔧 Configuration

### Environment Variables

| Variable | Required | Default | Description |
|----------|----------|---------|-------------|
| `AZURE_CLIENT_ID` | ✅ Yes | - | Application (client) ID from Azure |
| `AZURE_TENANT_ID` | ✅ Yes | - | Directory (tenant) ID from Azure |
| `AZURE_CLIENT_SECRET` | ✅ Yes | - | Client secret from Azure |
| `AZURE_REDIRECT_URI` | No | `http://localhost:3000/callback` | OAuth callback URL |
| `TOKEN_STORAGE` | No | `keytar` | Token storage backend (`keytar` or `1password`) |
| `LOG_LEVEL` | No | `info` | Logging level (`error`, `warn`, `info`, `debug`) |
| `RATE_LIMIT_PER_MINUTE` | No | `60` | Microsoft Graph API rate limit |
| `STATE_TIMEOUT_MINUTES` | No | `5` | OAuth state token expiration time |

### Token Storage Options

**Keytar (Default)** - Uses Windows Credential Manager:
- ✅ Free and built-in to Windows
- ✅ No additional setup required
- ✅ Secure OS-level credential storage
- ⚠️ Windows only

**1Password SDK** - Uses 1Password service account:
- ✅ Cross-platform (Windows, macOS, Linux)
- ✅ Enterprise-grade security
- ⚠️ Requires 1Password Teams subscription (~$8/month)
- ⚠️ Requires `OP_SERVICE_ACCOUNT_TOKEN` environment variable

To use 1Password:
1. Set `TOKEN_STORAGE=1password` in environment
2. Set `OP_SERVICE_ACCOUNT_TOKEN` with your service account token
3. See [1Password SDK Documentation](https://developer.1password.com/docs/sdks)

## 🛠️ Development

### Commands

```bash
# Development mode with auto-reload
npm run dev

# Run all tests (161 tests)
npm test

# Run with coverage (91.68%)
npm run test:coverage

# Run security-specific tests only
npm run test:security

# Type checking
npm run type-check

# Linting
npm run lint

# Clean build artifacts
npm run clean

# Production build
npm run build

# Start production server
npm start
```

### Project Structure

```
toto-mcp/
├── src/
│   ├── auth/              # OAuth & token management
│   │   ├── types.ts                      # ITokenManager interface
│   │   ├── keytar-token-manager.ts       # Windows Credential Manager
│   │   ├── one-password-token-manager.ts # 1Password SDK
│   │   ├── token-manager-factory.ts      # Factory pattern
│   │   ├── token-refresher.ts            # Automatic refresh
│   │   └── secure-oauth-client.ts        # CSRF-protected OAuth
│   ├── config/            # Configuration management
│   │   └── environment.ts                # Zod-validated config
│   ├── graph/             # Microsoft Graph API client
│   │   ├── client.ts                     # HTTP client
│   │   ├── rate-limiter.ts               # Token bucket algorithm
│   │   ├── circuit-breaker.ts            # Fail-fast pattern
│   │   └── todo-service.ts               # Business logic
│   ├── mcp/               # MCP server implementation
│   │   ├── server.ts                     # MCP SDK server
│   │   └── types.ts                      # Zod schemas
│   ├── security/          # Security primitives
│   │   ├── logger.ts                     # Winston with redaction
│   │   ├── errors.ts                     # Error classes
│   │   ├── validators.ts                 # Input validation
│   │   └── sanitizers.ts                 # Output sanitization
│   └── index.ts           # Entry point
├── tests/
│   ├── unit/              # Unit tests (144 tests)
│   └── security/          # Security tests (17 tests)
├── dist/                  # Compiled JavaScript (gitignored)
├── .env.example           # Example environment variables
├── .gitignore             # Git ignore rules
├── package.json           # npm configuration
├── tsconfig.json          # TypeScript configuration
├── vitest.config.ts       # Vitest test configuration
├── README.md              # This file
└── SECURITY.md            # Security documentation
```

## 🔒 Security

This project follows **security-first principles**:

### Defense in Depth (7 Layers)

1. **Input Validation** - Whitelist-based validation for OData filters, IDs, search queries
2. **Output Sanitization** - Strips Microsoft Graph internal fields from responses
3. **Logging Security** - Automatic redaction of `access_token`, `refresh_token`, `client_secret`
4. **Rate Limiting** - Token bucket algorithm (60 req/min) prevents API abuse
5. **Circuit Breaker** - Fail-fast pattern prevents cascading failures
6. **OAuth Security** - CSRF protection with 32-byte cryptographic state tokens
7. **Token Management** - Secure storage abstraction (never in files or logs)

### Security Testing

- **161 total tests** with **91.68% code coverage**
- **17 security-specific tests** covering:
  - SQL injection prevention
  - XSS prevention
  - Path traversal prevention
  - Command injection prevention
  - Null byte rejection
  - Error sanitization (no stack traces)
  - Rate limiting bypass prevention
  - Circuit breaker resource exhaustion prevention

### Security Audit

✅ **0 npm vulnerabilities** (verified with `npm audit`)

For detailed security information, see [SECURITY.md](./SECURITY.md).

## 🐛 Troubleshooting

### Authentication Issues

**Problem**: "Authentication required" error

**Solutions**:
1. Check that environment variables are set correctly:
   ```bash
   echo $env:AZURE_CLIENT_ID
   echo $env:AZURE_TENANT_ID
   echo $env:AZURE_CLIENT_SECRET
   ```
2. Verify Azure app registration has correct redirect URI: `http://localhost:3000/callback`
3. Check that OAuth scopes include `Tasks.Read` and `User.Read`
4. Clear stored tokens: Open Windows Credential Manager → Search for "toto-mcp" → Remove

### Rate Limiting

**Problem**: "Rate limit exceeded" errors

**Solutions**:
1. Default limit is 60 requests/minute (Microsoft Graph limit)
2. Increase limit (if needed): Set `RATE_LIMIT_PER_MINUTE` environment variable
3. Check for retry loops in your Claude Code prompts

### Build Errors

**Problem**: TypeScript compilation errors

**Solutions**:
1. Ensure Node.js v18+ and npm v9+ are installed:
   ```bash
   node --version  # Should be v18+
   npm --version   # Should be v9+
   ```
2. Clean and rebuild:
   ```bash
   npm run clean
   npm install
   npm run build
   ```

### Token Storage Issues

**Problem**: "Token storage unavailable" errors

**Solutions**:
1. **Keytar (Windows)**: Ensure Windows Credential Manager is accessible
2. **1Password**: Verify `OP_SERVICE_ACCOUNT_TOKEN` is set correctly
3. Check `TOKEN_STORAGE` environment variable matches your chosen backend

### Claude Code Not Loading MCP Server

**Problem**: Toto tools not available in Claude Code

**Solutions**:
1. Verify `.claude/mcp.json` exists and has correct path to `dist/index.js`
2. Ensure server is built: `npm run build`
3. Restart Claude Code after configuration changes
4. Check Claude Code logs for MCP server errors
5. Verify environment variables are set at system level (not just in terminal)

### Logging Issues

**Problem**: Need to debug server behavior

**Solutions**:
1. Set `LOG_LEVEL=debug` in environment variables
2. Check logs for errors and warnings
3. Winston logs include automatic timestamp and request ID tracking

## 📖 API Documentation

### Tool: `get_auth_status`

**Description**: Check current authentication status

**Parameters**: None

**Returns**:
```json
{
  "authenticated": true,
  "message": "Authentication valid"
}
```

### Tool: `list_task_lists`

**Description**: Get all task lists for the authenticated user

**Parameters**: None

**Returns**:
```json
[
  {
    "id": "list-123",
    "displayName": "My Tasks",
    "isOwner": true,
    "isShared": false,
    "wellknownListName": "defaultList"
  }
]
```

### Tool: `list_tasks`

**Description**: Get tasks from a list with optional filtering

**Parameters**:
- `listId` (optional): Specific list ID, or omit for all lists
- `filter` (optional): `"all"` | `"completed"` | `"incomplete"` | `"high-priority"` | `"today"`
- `limit` (optional): Number of tasks to return (1-100, default: 50)

**Returns**:
```json
[
  {
    "id": "task-123",
    "title": "Buy groceries",
    "status": "notStarted",
    "importance": "high",
    "createdDateTime": "2025-01-01T00:00:00Z",
    "lastModifiedDateTime": "2025-01-02T00:00:00Z",
    "dueDateTime": "2025-01-10T00:00:00Z",
    "isReminderOn": true,
    "body": "Milk, eggs, bread",
    "categories": ["Personal"]
  }
]
```

### Tool: `get_task`

**Description**: Get a specific task by ID

**Parameters**:
- `listId` (required): Task list ID
- `taskId` (required): Task ID

**Returns**: Same format as `list_tasks`

### Tool: `search_tasks`

**Description**: Search tasks across all lists

**Parameters**:
- `query` (required): Search query (1-100 characters)
- `limit` (optional): Number of results (1-50, default: 20)

**Returns**: Same format as `list_tasks`

## 🧪 Testing

### Run All Tests

```bash
npm test
```

**Expected Output**: 161 tests passing

### Run with Coverage

```bash
npm run test:coverage
```

**Expected Coverage**: >80% (currently 91.68%)

### Run Security Tests Only

```bash
npm run test:security
```

**Expected Output**: 17 security tests passing

### Test Structure

- `tests/unit/` - Unit tests for individual modules
- `tests/security/` - Security-specific tests (injection, CSRF, etc.)

## 📄 License

ISC

## 🙏 Acknowledgments

- **Anthropic** - For the Model Context Protocol SDK
- **Microsoft** - For the Graph API and MSAL library
- **Dorothy and Toto** - For the inspiration!

## 📞 Support

For issues or questions:
1. Check this README's troubleshooting section
2. Review [SECURITY.md](./SECURITY.md) for security-related questions
3. Check logs with `LOG_LEVEL=debug` for detailed diagnostics

---

**Built with ❤️ and a security-first mindset**
