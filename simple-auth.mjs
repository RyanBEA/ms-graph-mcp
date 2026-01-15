import http from 'http';
import { ConfidentialClientApplication } from '@azure/msal-node';
import fs from 'fs/promises';
import dotenv from 'dotenv';

dotenv.config();

const config = {
  auth: {
    clientId: process.env.AZURE_CLIENT_ID,
    clientSecret: process.env.AZURE_CLIENT_SECRET,
    authority: `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}`
  }
};

const pca = new ConfidentialClientApplication(config);
const scopes = ['Tasks.ReadWrite', 'Calendars.Read', 'User.Read', 'offline_access'];

// Generate auth URL
const authUrl = await pca.getAuthCodeUrl({
  scopes,
  redirectUri: 'http://localhost:3000/callback'
});

console.log('\n=== OPEN THIS URL IN YOUR BROWSER ===\n');
console.log(authUrl);
console.log('\n=====================================\n');

// Start server
const server = http.createServer(async (req, res) => {
  if (req.url?.startsWith('/callback')) {
    const url = new URL(req.url, 'http://localhost:3000');
    const code = url.searchParams.get('code');
    
    if (code) {
      try {
        const result = await pca.acquireTokenByCode({
          code,
          scopes,
          redirectUri: 'http://localhost:3000/callback'
        });
        
        // Save MSAL cache (always - this is the secure storage)
        const cache = pca.getTokenCache().serialize();
        await fs.writeFile('.msal-cache.json', cache);

        // Save account ID for token manager lookup
        if (result.account?.homeAccountId) {
          await fs.writeFile('.msal-account.json', JSON.stringify({
            accountId: result.account.homeAccountId
          }));
        }

        // Only write plaintext .tokens.json if explicitly using file storage
        if (process.env.TOKEN_STORAGE === 'file') {
          await fs.writeFile('.tokens.json', JSON.stringify({
            accessToken: result.accessToken,
            refreshToken: result.account?.homeAccountId ? 'stored_in_msal_cache' : null,
            expiresAt: result.expiresOn?.getTime() // numeric timestamp for file storage
          }, null, 2));
          console.log('Note: .tokens.json written (TOKEN_STORAGE=file mode)');
        }
        
        console.log('\n✅ SUCCESS! Tokens saved.\n');
        res.end('<h1>Success! Tokens saved. Close this window.</h1>');
        setTimeout(() => process.exit(0), 1000);
      } catch (err) {
        console.error('Token exchange failed:', err.message);
        res.end('<h1>Error: ' + err.message + '</h1>');
      }
    } else {
      const error = url.searchParams.get('error_description') || url.searchParams.get('error');
      console.error('Auth error:', error);
      res.end('<h1>Error: ' + error + '</h1>');
    }
  }
});

server.listen(3000, () => console.log('Waiting for callback on port 3000...'));
