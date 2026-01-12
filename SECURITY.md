# Security Architecture - Toto MCP Server

**Last Updated:** 2025-11-17
**Version:** 1.0.0
**Security Review Status:** ✅ Phase 6 Complete

---

## Overview

Toto MCP Server is built with **security-first** principles. Every component follows a zero-trust architecture with defense-in-depth strategies.

### Core Security Principles

1. **Never Trust Input** - All inputs validated before processing
2. **Never Log Secrets** - Automatic redaction of sensitive data
3. **Fail Securely** - Generic error messages, no information disclosure
4. **Read-Only by Default** - No write operations to Microsoft Graph
5. **Rate Limiting** - Prevent abuse and resource exhaustion
6. **Defense in Depth** - Multiple layers of security controls

---

## Threat Model

### Assets Protected

1. **Authentication Tokens** - OAuth access/refresh tokens
2. **User Data** - Microsoft To Do tasks and lists
3. **System Resources** - CPU, memory, network bandwidth
4. **Configuration** - Azure credentials, API keys

### Threat Actors

1. **Malicious Users** - Attempting injection attacks, data exfiltration
2. **Compromised Claude Desktop** - If client is compromised
3. **Network Attackers** - Man-in-the-middle, eavesdropping
4. **Resource Exhaustion** - DoS attempts via rate limiting bypass

### Attack Vectors & Mitigations

| Attack Vector | Mitigation | Implementation |
|--------------|------------|----------------|
| SQL Injection | Input validation with whitelist | `validators.ts` - OData filter whitelist |
| XSS | Input sanitization, output encoding | `validators.ts` - Pattern blocking |
| Command Injection | Shell metacharacter blocking | `validators.ts` - Character whitelist |
| Path Traversal | ID format validation | `validators.ts` - Alphanumeric only |
| Token Theft (Logs) | Automatic log redaction | `logger.ts` - Winston transform |
| Token Theft (Errors) | Error sanitization | `sanitizers.ts` - No stack traces |
| Rate Limit Bypass | Token bucket algorithm | `rate-limiter.ts` - Strict enforcement |
| Resource Exhaustion | Circuit breaker pattern | `circuit-breaker.ts` - Fail-fast |
| Information Disclosure | Generic error messages | All error classes extend `AppError` |
| CSRF | Cryptographic state tokens | `secure-oauth-client.ts` - 32-byte random |

---

## Security Controls

### 1. Input Validation

**Location:** `src/security/validators.ts`

All user inputs pass through Zod schemas and custom validators:

- **OData Filters** - Whitelist of 7 allowed fields, 9 allowed operators
- **List/Task IDs** - Alphanumeric + hyphens/underscores only
- **Search Queries** - Max 1000 chars, blocked shell metacharacters
- **Null Byte Protection** - Rejected in all inputs

**Test Coverage:** 20 tests in `validators.test.ts`

### 2. Output Sanitization

**Location:** `src/security/sanitizers.ts`

All responses sanitized before returning to clients:

- **Task Sanitization** - Removes Microsoft Graph internal fields
- **Error Sanitization** - Never exposes stack traces or file paths
- **Token Redaction** - No authentication tokens in responses

**Test Coverage:** 15 tests in `sanitizers.test.ts`

### 3. Logging Security

**Location:** `src/security/logger.ts`

Winston logger with automatic sensitive data redaction:

- **Redacted Fields:** `access_token`, `refresh_token`, `client_secret`, `Authorization`, `password`
- **Log Levels:** ERROR, WARN, INFO, DEBUG (configurable)
- **Retention:** Daily rotation, max 14 days

**Test Coverage:** 9 tests in `logger.test.ts`

### 4. Rate Limiting

**Location:** `src/graph/rate-limiter.ts`

Token bucket algorithm prevents API abuse:

- **Default:** 60 requests/minute
- **Burst Allowance:** Configurable (default = rate limit)
- **Refill Rate:** Sub-second precision
- **Failure Mode:** Throws `RateLimitError` with retry time

**Test Coverage:** 14 tests in `rate-limiter.test.ts`

### 5. Circuit Breaker

**Location:** `src/graph/circuit-breaker.ts`

Fail-fast pattern prevents cascading failures:

- **States:** CLOSED → OPEN → HALF_OPEN
- **Failure Threshold:** 5 failures (configurable)
- **Recovery Test:** 2 successes needed (configurable)
- **Timeout:** 60 seconds before retry (configurable)

**Test Coverage:** 16 tests in `circuit-breaker.test.ts`

### 6. OAuth Security

**Location:** `src/auth/secure-oauth-client.ts`

CSRF-protected OAuth 2.0 flow:

- **State Tokens:** 32-byte cryptographically secure random
- **One-Time Use:** States deleted after validation
- **Time-Based Expiration:** 5 minutes (configurable)
- **Auto Cleanup:** Expired states removed every 60 seconds

**Test Coverage:** 15 tests in `secure-oauth-client.test.ts`

### 7. Token Management

**Location:** `src/auth/`

Secure token storage with multiple backends:

- **Windows:** Credential Manager via keytar (development)
- **Production:** 1Password SDK (optional, future)
- **In-Memory:** Never stored in code or logs
- **Automatic Refresh:** 5-minute buffer before expiry

**Test Coverage:** 21 tests across token manager tests

---

## Security Testing

### Test Suite Summary

**Total Tests:** 161
**Security-Specific Tests:** 17 (in `tests/security/security.test.ts`)

### Security Test Categories

1. **Input Validation** (6 tests)
   - SQL injection prevention
   - XSS prevention
   - Path traversal prevention
   - Command injection prevention
   - Length limit enforcement
   - Null byte rejection

2. **Output Sanitization** (3 tests)
   - No Microsoft Graph internals exposed
   - No stack traces in errors
   - No authentication tokens in responses

3. **Error Handling** (2 tests)
   - Generic error messages
   - Safe error codes only

4. **Rate Limiting** (2 tests)
   - Bypass prevention
   - No negative token manipulation

5. **Circuit Breaker** (1 test)
   - Resource exhaustion prevention

6. **Data Integrity** (1 test)
   - Required fields validation

7. **CSRF Protection** (1 test)
   - Code execution prevention

### Dependency Security

**npm audit:** ✅ 0 vulnerabilities
**Last Checked:** 2025-11-17

---

## Security Configuration

### Required Environment Variables

```bash
# Azure OAuth (REQUIRED)
AZURE_CLIENT_ID=your-client-id
AZURE_TENANT_ID=your-tenant-id
AZURE_CLIENT_SECRET=your-client-secret  # NEVER commit!

# Token Storage (REQUIRED)
TOKEN_STORAGE=keytar  # or "1password"

# Optional Security Settings
LOG_LEVEL=info  # error|warn|info|debug
RATE_LIMIT_PER_MINUTE=60
STATE_TIMEOUT_MINUTES=5
```

### Security Checklist

Before deploying to production:

- [ ] All environment variables set
- [ ] `AZURE_CLIENT_SECRET` stored securely (not in .env file)
- [ ] Token storage configured (keytar or 1Password)
- [ ] Log level set appropriately (`info` or `warn` for production)
- [ ] Rate limits configured for expected load
- [ ] npm audit shows 0 vulnerabilities
- [ ] All 161 tests passing
- [ ] Security documentation reviewed

---

## Incident Response

### Suspected Token Compromise

1. **Immediate:** Revoke tokens in Azure Portal
2. **Clear Storage:** `tokenManager.clearTokens()`
3. **Review Logs:** Check for suspicious activity
4. **Re-authenticate:** User must complete OAuth flow again

### Suspected Injection Attack

1. **Check Logs:** Review validator warnings
2. **Verify Blocking:** Confirm ValidationError thrown
3. **Update Patterns:** Add new patterns to validators if needed
4. **Test Coverage:** Add test case for new pattern

### Rate Limit Abuse

1. **Check Logs:** Review RateLimitError frequency
2. **Adjust Limits:** Lower `RATE_LIMIT_PER_MINUTE` if needed
3. **Monitor:** Watch for sustained high-frequency requests
4. **Circuit Breaker:** Verify circuit opens appropriately

---

## Security Best Practices for Users

1. **Keep Credentials Secure**
   - Never commit `.env` files
   - Use environment variables or secret managers
   - Rotate `AZURE_CLIENT_SECRET` periodically

2. **Monitor Logs**
   - Review logs for security warnings
   - Set up alerts for ERROR level logs
   - Keep logs for audit trail (14-day retention)

3. **Update Regularly**
   - Run `npm audit` monthly
   - Update dependencies with security patches
   - Review security advisories

4. **Least Privilege**
   - OAuth scopes: `Tasks.Read`, `User.Read` only
   - No write permissions requested
   - Limit Azure app registration permissions

---

## Security Audit History

| Date | Auditor | Findings | Status |
|------|---------|----------|--------|
| 2025-11-17 | Phase 6 Security Hardening | 0 critical, 0 high, 0 medium | ✅ PASSED |

---

## Contact

For security issues, please follow responsible disclosure:

1. **Do NOT** create public GitHub issues for security vulnerabilities
2. Contact the maintainer privately
3. Allow reasonable time for fix before public disclosure

---

## References

- [OWASP Top 10](https://owasp.org/www-project-top-ten/)
- [OAuth 2.0 Security Best Current Practice](https://datatracker.ietf.org/doc/html/draft-ietf-oauth-security-topics)
- [Microsoft Graph Security Best Practices](https://learn.microsoft.com/en-us/graph/security-authorization)
- [Node.js Security Best Practices](https://nodejs.org/en/docs/guides/security/)
