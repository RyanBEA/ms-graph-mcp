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
| `Tasks.Read` | Read user's tasks and task lists |
| `Calendars.Read` | Read user's calendar events |
| `User.Read` | Sign in and read user profile |
| `offline_access` | Maintain access via refresh tokens |

4. Click **Add permissions**
5. If you see "Grant admin consent" and have admin access, click it

### Step 5: Enable Public Client Flows (CRITICAL!)

**This step is required for Device Code Flow authentication.**

1. Go to **Authentication**
2. Scroll to **Advanced settings**
3. Set **"Allow public client flows"** to **Yes**
4. Click **Save**

> Without this setting, Device Code Flow authentication will fail with an "invalid_client" error.

---

## Authentication Methods

### Method 1: Device Code Flow (Recommended)

Device Code Flow is the most reliable method. It doesn't require a localhost callback, avoiding browser caching issues.

```bash
cd C:/ai/mcp-servers/ms-graph
node device-code-auth.mjs
```

**What happens:**
1. Script displays a URL and a user code
2. Open the URL in any browser (even on a different device)
3. Enter the code shown in the terminal
4. Sign in with your Microsoft account
5. Grant the requested permissions
6. Script automatically saves tokens and exits

**Advantages:**
- No localhost callback required
- Works from any terminal/SSH session
- Avoids browser redirect caching problems
- Can authenticate from a different device

### Method 2: OAuth Callback Flow (Alternative)

Traditional OAuth flow with localhost redirect. Use this if Device Code doesn't work for your scenario.

```bash
cd C:/ai/mcp-servers/ms-graph
node simple-auth.mjs
```

**What happens:**
1. Script displays an authorization URL
2. Open the URL in your browser
3. Sign in and grant permissions
4. Browser redirects to `localhost:3000/callback`
5. Script captures the code and exchanges for tokens

**Known Issues:**
- Browser may cache failed redirects with corrupted data
- Port 3000 must be available
- If issues occur, use an incognito/private browser window

---

## Configuration

### Option A: Environment Variables in `.env`

Create `.env` in the ms-graph directory:

```env
AZURE_CLIENT_ID=your-application-client-id
AZURE_TENANT_ID=your-directory-tenant-id
AZURE_CLIENT_SECRET=your-client-secret-value
AZURE_REDIRECT_URI=http://localhost:3000/callback
TOKEN_STORAGE=file
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
        "TOKEN_STORAGE": "file",
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

After successful authentication, tokens are stored in:

| File | Purpose |
|------|---------|
| `.msal-cache.json` | MSAL token cache (access + refresh tokens) |
| `.msal-account.json` | Account ID reference |

These files are located in the ms-graph directory and are gitignored.

**Token Refresh**: MSAL automatically refreshes access tokens using the stored refresh token. No manual intervention needed.

---

## Troubleshooting

### "Not authenticated" after successful auth

**Symptoms**: Auth script says success, but `get_auth_status` returns false.

**Causes & Fixes**:
1. **Token files not found**: Verify `.msal-cache.json` exists in ms-graph directory
2. **Wrong working directory**: Ensure `cwd` is set in `~/.claude.json`
3. **Restart required**: Restart Claude Code after authentication

### "invalid_client" error during Device Code Flow

**Cause**: Azure app not configured for public client flows.

**Fix**:
1. Azure Portal → App registration → Authentication
2. Set "Allow public client flows" to **Yes**
3. Save and retry

### "Invalid client secret" error

**Cause**: Using the Secret ID instead of the Secret Value.

**Fix**:
1. Go to Azure Portal → Certificates & secrets
2. Create a new client secret
3. Copy the **Value** (not the ID) immediately
4. Update your `.env` or `~/.claude.json`

### Browser shows scope errors with garbled text

**Cause**: Browser cached redirects from previous failed auth attempts.

**Symptoms**: Error messages show corrupted scopes like `offlne_access` instead of `offline_access`.

**Fix**:
1. Close ALL browser windows completely
2. Use an incognito/private browser window
3. Or switch to Device Code Flow (recommended)

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

If tokens expire, get corrupted, or you need to change accounts:

1. **Delete existing tokens**:
   ```bash
   cd C:/ai/mcp-servers/ms-graph
   rm .msal-cache.json .msal-account.json
   ```

2. **Run authentication**:
   ```bash
   node device-code-auth.mjs
   ```

3. **Restart Claude Code** to pick up new tokens

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

### Read-Only Access
This server only uses read permissions:
- Cannot create, modify, or delete tasks
- Cannot create, modify, or delete calendar events
- Cannot access mail, files, or other sensitive data

---

## Quick Reference

### Authentication Commands

```bash
# Recommended: Device Code Flow
node device-code-auth.mjs

# Alternative: OAuth Callback
node simple-auth.mjs
```

### Required Environment Variables

```
AZURE_CLIENT_ID      # From app registration Overview
AZURE_TENANT_ID      # From app registration Overview
AZURE_CLIENT_SECRET  # From Certificates & secrets (VALUE, not ID)
```

### Required Azure Permissions

- `Tasks.Read` (Delegated)
- `Calendars.Read` (Delegated)
- `User.Read` (Delegated)
- `offline_access` (Delegated)

### Required Azure Settings

- Redirect URI: `http://localhost:3000/callback`
- Allow public client flows: **Yes**

---

## Need Help?

1. Check this guide's [Troubleshooting](#troubleshooting) section
2. Review the [SECURITY.md](SECURITY.md) for security architecture
3. Check Azure Portal for any consent or permission issues
4. Ensure all required permissions have admin consent
