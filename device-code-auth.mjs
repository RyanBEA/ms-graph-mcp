/**
 * Device Code Flow Authentication
 *
 * This bypasses localhost redirect issues by using the device code flow.
 * User visits a URL and enters a code manually - no callback server needed.
 */

import { PublicClientApplication } from '@azure/msal-node';
import fs from 'fs/promises';
import dotenv from 'dotenv';

dotenv.config();

const config = {
  auth: {
    clientId: process.env.AZURE_CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}`
  }
};

const pca = new PublicClientApplication(config);
const scopes = ['Tasks.Read', 'Calendars.Read', 'User.Read', 'offline_access'];

console.log('\n=== Device Code Authentication ===\n');

try {
  const response = await pca.acquireTokenByDeviceCode({
    scopes,
    deviceCodeCallback: (response) => {
      console.log('╔════════════════════════════════════════════════════════════╗');
      console.log('║                    AUTHENTICATION REQUIRED                  ║');
      console.log('╠════════════════════════════════════════════════════════════╣');
      console.log('║                                                            ║');
      console.log(`║  1. Open: ${response.verificationUri.padEnd(44)}║`);
      console.log('║                                                            ║');
      console.log(`║  2. Enter code: ${response.userCode.padEnd(38)}║`);
      console.log('║                                                            ║');
      console.log('╚════════════════════════════════════════════════════════════╝');
      console.log('\nWaiting for authentication...\n');
    }
  });

  if (response && response.accessToken) {
    // Save tokens to file
    const tokenData = {
      accessToken: response.accessToken,
      account: response.account,
      expiresOn: response.expiresOn?.toISOString(),
      scopes: response.scopes
    };

    await fs.writeFile('.tokens.json', JSON.stringify(tokenData, null, 2));

    // Save MSAL cache
    const cache = pca.getTokenCache().serialize();
    await fs.writeFile('.msal-cache.json', cache);

    // Save account info for the MCP server
    if (response.account) {
      await fs.writeFile('.msal-account.json', JSON.stringify({
        accountId: response.account.homeAccountId
      }));
    }

    console.log('╔════════════════════════════════════════════════════════════╗');
    console.log('║                    ✅ SUCCESS!                              ║');
    console.log('╠════════════════════════════════════════════════════════════╣');
    console.log('║  Tokens saved to:                                          ║');
    console.log('║    - .tokens.json                                          ║');
    console.log('║    - .msal-cache.json                                      ║');
    console.log('║    - .msal-account.json                                    ║');
    console.log('║                                                            ║');
    console.log('║  You can now restart Claude Code to use the MCP server.   ║');
    console.log('╚════════════════════════════════════════════════════════════╝');
    console.log('\nAuthenticated as:', response.account?.username || 'Unknown');
    console.log('Scopes granted:', response.scopes?.join(', ') || 'Unknown');
  }
} catch (error) {
  console.error('\n❌ Authentication failed:', error.message);
  if (error.errorCode) {
    console.error('Error code:', error.errorCode);
  }
  process.exit(1);
}
