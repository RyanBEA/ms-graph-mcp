#!/usr/bin/env node
/**
 * Authentication helper script for Toto MCP Server.
 * Runs the OAuth flow and stores tokens securely.
 */

import http from 'http';
import { URL } from 'url';
import { SecureOAuthClient } from '../auth/secure-oauth-client.js';
import { getTokenManager } from '../auth/token-manager-factory.js';
import { logger } from '../security/logger.js';
import { loadConfiguration } from '../config/environment.js';

const PORT = 3000;
const TIMEOUT = 5 * 60 * 1000; // 5 minutes

async function authenticate() {
  try {
    console.log('\n🐕 Toto MCP Server - Authentication Flow\n');

    // Load configuration
    loadConfiguration();

    console.log('Initializing token manager...');
    const tokenManager = await getTokenManager();

    // Check if already authenticated
    const hasTokens = await tokenManager.hasValidTokens();
    if (hasTokens) {
      console.log('\n✅ Already authenticated!');
      console.log('Your tokens are stored securely.');
      console.log('\nIf you want to re-authenticate, first clear your tokens:');
      console.log('  - Windows Credential Manager: Search for "Toto MCP"');
      console.log('\n');
      process.exit(0);
    }

    console.log('Creating OAuth client...');
    const oauthClient = new SecureOAuthClient(tokenManager);

    // Generate authorization URL
    console.log('Generating authorization URL...\n');
    const { url } = await oauthClient.generateAuthUrl();

    console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    console.log('📋 Please complete the following steps:');
    console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n');
    console.log('1. Open this URL in your browser:\n');
    console.log(`   ${url}\n`);
    console.log('2. Sign in with your Microsoft account');
    console.log('3. Grant permissions to access your To Do tasks');
    console.log('4. You will be redirected to localhost:3000/callback');
    console.log('   (the authentication will complete automatically)\n');
    console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n');

    // Create a promise that resolves when OAuth completes
    let resolveAuth: (value: boolean) => void;
    const authPromise = new Promise<boolean>((resolve) => {
      resolveAuth = resolve;
    });

    // Start HTTP server to handle callback
    const server = http.createServer(async (req, res) => {
      try {
        if (!req.url) {
          res.writeHead(400, { 'Content-Type': 'text/plain' });
          res.end('Bad Request');
          return;
        }

        const parsedUrl = new URL(req.url, `http://localhost:${PORT}`);

        if (parsedUrl.pathname === '/callback') {
          const code = parsedUrl.searchParams.get('code');
          const returnedState = parsedUrl.searchParams.get('state');

          if (!code || !returnedState) {
            res.writeHead(400, { 'Content-Type': 'text/html' });
            res.end(`
              <!DOCTYPE html>
              <html>
                <head><title>Authentication Error</title></head>
                <body style="font-family: system-ui, sans-serif; padding: 40px; max-width: 600px; margin: 0 auto;">
                  <h1 style="color: #d32f2f;">❌ Authentication Error</h1>
                  <p>Missing authorization code or state parameter.</p>
                  <p>Please try the authentication process again.</p>
                </body>
              </html>
            `);
            resolveAuth(false);
            return;
          }

          try {
            console.log('Received callback, exchanging code for tokens...');
            await oauthClient.handleCallback(code, returnedState);

            res.writeHead(200, { 'Content-Type': 'text/html' });
            res.end(`
              <!DOCTYPE html>
              <html>
                <head><title>Authentication Successful</title></head>
                <body style="font-family: system-ui, sans-serif; padding: 40px; max-width: 600px; margin: 0 auto;">
                  <h1 style="color: #2e7d32;">✅ Authentication Successful!</h1>
                  <p>Your Microsoft To Do account has been connected to Toto.</p>
                  <p>Your tokens are stored securely in Windows Credential Manager.</p>
                  <p><strong>You can now close this window and return to Claude Code.</strong></p>
                </body>
              </html>
            `);

            resolveAuth(true);
          } catch (error) {
            logger.error('OAuth callback failed', { error });

            res.writeHead(500, { 'Content-Type': 'text/html' });
            res.end(`
              <!DOCTYPE html>
              <html>
                <head><title>Authentication Error</title></head>
                <body style="font-family: system-ui, sans-serif; padding: 40px; max-width: 600px; margin: 0 auto;">
                  <h1 style="color: #d32f2f;">❌ Authentication Error</h1>
                  <p>Failed to complete authentication.</p>
                  <p>Error: ${error instanceof Error ? error.message : 'Unknown error'}</p>
                  <p>Please try the authentication process again.</p>
                </body>
              </html>
            `);

            resolveAuth(false);
          }
        } else {
          res.writeHead(404, { 'Content-Type': 'text/plain' });
          res.end('Not Found');
        }
      } catch (error) {
        logger.error('Server error', { error });
        res.writeHead(500, { 'Content-Type': 'text/plain' });
        res.end('Internal Server Error');
      }
    });

    // Start server
    await new Promise<void>((resolve) => {
      server.listen(PORT, () => {
        console.log(`Waiting for OAuth callback on http://localhost:${PORT}/callback`);
        console.log('(This will timeout in 5 minutes)\n');
        resolve();
      });
    });

    // Set timeout
    const timeoutHandle = setTimeout(() => {
      console.log('\n⏰ Timeout: Authentication flow took too long.');
      console.log('Please try again.\n');
      resolveAuth(false);
    }, TIMEOUT);

    // Wait for authentication to complete
    const success = await authPromise;

    // Cleanup
    clearTimeout(timeoutHandle);
    oauthClient.stopStateCleanup();
    server.close();

    if (success) {
      // Manually serialize and save MSAL cache
      console.log('Saving tokens to persistent cache...');
      const msalClient = (tokenManager as any).getMsalClient();
      if (msalClient) {
        const fs = await import('fs/promises');
        const cache = msalClient.getTokenCache();
        const cacheData = cache.serialize();
        await fs.writeFile('./.msal-cache.json', cacheData);
        console.log('Cache file created: .msal-cache.json');
      }

      console.log('\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
      console.log('✅ Authentication Complete!');
      console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n');
      console.log('Your tokens are now stored securely.');
      console.log('You can now use Toto with Claude Code!\n');
      console.log('Try asking Claude: "list my task lists"\n');
      process.exit(0);
    } else {
      console.log('\n❌ Authentication failed. Please try again.\n');
      process.exit(1);
    }
  } catch (error) {
    logger.error('Authentication failed', { error });
    console.error('\n❌ Error:', error instanceof Error ? error.message : 'Unknown error');
    console.error('\nPlease check your configuration and try again.\n');
    process.exit(1);
  }
}

// Run authentication
authenticate();
