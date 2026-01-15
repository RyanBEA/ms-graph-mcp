# MS-Graph MCP Authentication Guide

Complete guide for authenticating the MS-Graph MCP server with Microsoft Graph API.

## Table of Contents

1. [Prerequisites](#prerequisites)
2. [Azure Portal Setup](#azure-portal-setup)
3. [Authentication Methods](#authentication-methods)
4. [Configuration](#configuration)
5. [Token Storage](#token-storage)
6. [Troubleshooting](#troubleshooting)
7. [Re-authentication](#re-authentication)
8. [Security Notes](#security-notes)

---

## Prerequisites

- **Node.js 18+** installed
- **Azure account** with permissions to create app registrations
- **Microsoft 365 account** (personal or work/school)
- **Built server**: Run `npm install && npm run build` first

---

## Azure Portal Setup

### Step 1: Create App Registration

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Microsoft Entra ID** (formerly Azure AD)
3. Click **App registrations** → **New registration**
4. Fill in:
   - **Name**: `ms-graph-mcp` (or your preferred name)
   - **Supported account types**: "Accounts in any organizational directory and personal Microsoft accounts"
   - **Redirect URI**: Select `Web` and enter `http://localhost:3000/callback`
5. Click **Register**

### Step 2: Collect Required Values

From the app's **Overview** page, copy:
- **Application (client) ID** → This is your `AZURE_CLIENT_ID`
- **Directory (tenant) ID** → This is your `AZURE_TENANT_ID`

### Step 3: Create Client Secret

1. Go to **Certificates & secrets** → **Client secrets**
2. Click **New client secret**
3. Add a description and select expiration
4. Click **Add**
5. **IMMEDIATELY copy the Value** → This is your `AZURE_CLIENT_SECRET`

> **WARNING**: The secret VALUE is only visible once! After you leave this page, you can only see the Secret ID (which is NOT what you need). If you lose it, create a new secret.

### Step 4: Configure API Permissions

1. Go to **API permissions** → **Add a permission**
2. Select **Microsoft Graph** → **Delegated permissions**
3. Add these permissions:

| Permission | Description |
|------------|-------------|
| `Tasks.ReadWrite` | Read and write user's To Do and Planner tasks |
| `Calendars.Read` | Read user's calendar events |
| `User.Read` | Sign in and read user profile |
| `offline_access` | Maintain access via refresh tokens |

4. Click **Add permissions**
5. If you see "Grant admin consent" and have admin access, click it

> **Note**: If upgrading from a previous version that used `Tasks.Read`, you must re-authenticate to grant the new `Tasks.ReadWrite` permission.

### Step 5: Configure as Confidential Client (CRITICAL!)

**This server uses a client secret, so it must be configured as a confidential client.**

1. Go to **Authentication**
2. Scroll to **Advanced settings**
3. Set **"Allow public client flows"** to **No**
4. Click **Save**

> **Why No?** This server runs on your machine (like a server) and keeps the client secret secure. Setting to "Yes" makes it a public client which cannot use secrets. Using "No" is the more secure option.

---

## Authentication Methods

### Method 1: MSAL OAuth Flow (Recommended)

This method uses MSAL (Microsoft Authentication Library) for OAuth with automatic token refresh.

```bash
cd C:/ai/mcp-servers/ms-graph
node simple-auth.mjs
```

**What happens:**
1. Script displays an authorization URL
2. Open the URL in an **incognito/private browser window** (important!)
3. Sign in with your Microsoft account
4. Grant the requested permissions
5. Browser redirects to `localhost:3000/callback`
6. Script exchanges the code for tokens using MSAL
7. Tokens saved to `.msal-cache.json`

**After authentication:**
- Create the account file if missing (see Troubleshooting)
- **Restart Claude Code** to load the MCP server with new tokens

**Tips:**
- Always use an incognito/private browser to avoid cached redirect issues
- Ensure port 3000 is not in use by another application
- MSAL automatically refreshes tokens when they expire (~1 hour)

---

## Configuration

### Option A: Environment Variables in `.env`

Create `.env` in the ms-graph directory:

```env
AZURE_CLIENT_ID=your-application-client-id
AZURE_TENANT_ID=your-directory-tenant-id
AZURE_CLIENT_SECRET=your-client-secret-value
AZURE_REDIRECT_URI=http://localhost:3000/callback
TOKEN_STORAGE=msal
LOG_LEVEL=info
```

### Option B: Claude Code Configuration

Add to `~/.claude.json`:

```json
{
  "mcpServers": {
    "ms-graph": {
      "type": "stdio",
      "command": "node",
      "args": ["C:/ai/mcp-servers/ms-graph/dist/index.js"],
      "env": {
        "AZURE_CLIENT_ID": "your-application-client-id",
        "AZURE_TENANT_ID": "your-directory-tenant-id",
        "AZURE_CLIENT_SECRET": "your-client-secret-value",
        "AZURE_REDIRECT_URI": "http://localhost:3000/callback",
        "TOKEN_STORAGE": "msal",
        "LOG_LEVEL": "info"
      },
      "cwd": "C:/ai/mcp-servers/ms-graph"
    }
  }
}
```

> **Note**: The `cwd` setting is important - it determines where token files are stored.

---

## Token Storage

With `TOKEN_STORAGE=msal` (recommended), tokens are stored in:

| File | Purpose |
|------|---------|
| `.msal-cache.json` | MSAL token cache (access, refresh, ID tokens) |
| `.msal-account.json` | Account ID for cache lookup |

These files are located in the ms-graph directory and are gitignored.

**Token Refresh**: MSAL automatically refreshes access tokens when they expire (~1 hour) using `acquireTokenSilent()`. No manual intervention required.

> **Note**: The `file` storage option (`.tokens.json`) is also available but does not support automatic token refresh.

---

## Troubleshooting

### "Not authenticated" after successful auth

**Symptoms**: Auth script says success, but `get_auth_status` returns false.

**Causes & Fixes**:
1. **MSAL files not found**: Verify `.msal-cache.json` and `.msal-account.json` exist in ms-graph directory
2. **Missing account file**: If `.msal-account.json` is missing, extract the account ID from `.msal-cache.json` (look for `home_account_id`) and create it:
   ```json
   {"accountId":"your-home-account-id-here"}
   ```
3. **Wrong working directory**: Ensure `cwd` is set in `~/.claude.json`
4. **Restart required**: Restart Claude Code after authentication
5. **Wrong TOKEN_STORAGE**: Ensure `TOKEN_STORAGE=msal` in config

### "Client is public" error (AADSTS700025)

**Symptoms**: Error says "Client is public so neither 'client_assertion' nor 'client_secret' should be presented"

**Cause**: Azure app is configured as a public client but code is using a client secret.

**Fix**:
1. Azure Portal → App registration → Authentication
2. Set "Allow public client flows" to **No**
3. Clear token cache: `rm .msal-cache.json .msal-account.json`
4. Re-authenticate: `node simple-auth.mjs` (use incognito browser)

### "Invalid client secret" error

**Cause 1**: Using the Secret ID instead of the Secret Value.

**Fix**:
1. Go to Azure Portal → Certificates & secrets
2. Create a new client secret
3. Copy the **Value** (not the ID) immediately
4. Update your `.env` or `~/.claude.json`

**Cause 2**: Windows system environment variable overriding `.env` file.

**Symptoms**: Your `.env` has the correct secret but authentication fails. The wrong secret is being used.

**Diagnosis**:
```bash
# Check if a system env var exists
powershell -Command "[Environment]::GetEnvironmentVariable('AZURE_CLIENT_SECRET', 'User')"
```

**Fix**:
```bash
# Remove the system environment variable
powershell -Command "[Environment]::SetEnvironmentVariable('AZURE_CLIENT_SECRET', '', 'User')"
```

> **Warning**: Never set Azure credentials as Windows system environment variables. They override `.env` files and can cause hard-to-debug issues. Always use `.env` or `~/.claude.json` instead.

### Browser shows scope errors with garbled text

**Cause**: Browser cached redirects from previous failed auth attempts.

**Symptoms**: Error messages show corrupted scopes like `offlne_access` instead of `offline_access`.

**Fix**:
1. Close ALL browser windows completely
2. Use an incognito/private browser window
3. Clear browser cache for Microsoft login sites

### "Port 3000 already in use"

**Cause**: Previous auth server still running.

**Fix (Windows)**:
```bash
netstat -ano | findstr :3000
taskkill /F /PID <pid-from-above>
```

**Fix (Mac/Linux)**:
```bash
lsof -i :3000
kill -9 <pid-from-above>
```

### Token refresh fails silently

**Cause**: MSAL cache not properly initialized.

**Fix**: This was fixed in the codebase. Ensure you have the latest version with `ensureCacheLoaded()` in `msal-token-manager.ts`.

### "AADSTS650053: scope doesn't exist"

**Cause**: API permissions not configured or admin consent not granted.

**Fix**:
1. Azure Portal → API permissions
2. Verify all required permissions are added
3. Click "Grant admin consent" if available

---

## Re-authentication

If tokens get corrupted or you need to change accounts:

1. **Delete existing tokens**:
   ```bash
   cd C:/ai/mcp-servers/ms-graph
   rm .msal-cache.json .msal-account.json
   ```

2. **Run authentication** (use incognito browser):
   ```bash
   node simple-auth.mjs
   ```

3. **Create account file** if missing (extract `home_account_id` from `.msal-cache.json`):
   ```bash
   # Example - use actual ID from your cache file
   echo '{"accountId":"your-home-account-id"}' > .msal-account.json
   ```

4. **Restart Claude Code** to pick up new tokens

> **Note**: With MSAL storage, tokens auto-refresh so re-authentication is rarely needed.

---

## Security Notes

### What NOT to commit
- `.env` file (contains secrets)
- `.msal-cache.json` (contains tokens)
- `.msal-account.json` (contains account info)

These are already in `.gitignore`.

### Client Secret Expiration
- Azure client secrets expire (configurable: 6 months, 1 year, 2 years, or custom)
- Set a calendar reminder before expiration
- To rotate: Create new secret → Update config → Delete old secret

### Token Security
- Access tokens are short-lived (~1 hour)
- Refresh tokens are long-lived but rotated automatically by MSAL
- All tokens are stored locally, never transmitted except to Microsoft

### Task Management Access
This server supports full task management:
- **To Do**: Create, read, update, complete, and delete tasks and task lists
- **Planner**: Create, read, update, complete, and delete Planner tasks
- **Calendar**: Read-only access to calendar events
- Cannot access mail, files, or other sensitive data

---

## Quick Reference

### Authentication Command

```bash
cd C:/ai/mcp-servers/ms-graph
node simple-auth.mjs
# Open URL in incognito browser, then restart Claude Code
```

### Required Environment Variables

```
AZURE_CLIENT_ID      # From app registration Overview
AZURE_TENANT_ID      # From app registration Overview
AZURE_CLIENT_SECRET  # From Certificates & secrets (VALUE, not ID)
TOKEN_STORAGE=msal   # Use MSAL for automatic token refresh
```

### Required Azure Permissions

- `Tasks.ReadWrite` (Delegated) - For To Do and Planner tasks
- `Calendars.Read` (Delegated)
- `User.Read` (Delegated)
- `offline_access` (Delegated)

### Required Azure Settings

- Redirect URI: `http://localhost:3000/callback` (Web platform)
- Allow public client flows: **No** (confidential client with secret)

---

## Need Help?

1. Check this guide's [Troubleshooting](#troubleshooting) section
2. Review the [SECURITY.md](SECURITY.md) for security architecture
3. Check Azure Portal for any consent or permission issues
4. Ensure all required permissions have admin consent
