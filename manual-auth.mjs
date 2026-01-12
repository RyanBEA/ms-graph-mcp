import http from 'http';
import https from 'https';
import fs from 'fs/promises';
import { URL, URLSearchParams } from 'url';
import dotenv from 'dotenv';

dotenv.config();

const CLIENT_ID = process.env.AZURE_CLIENT_ID;
const CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET;
const TENANT_ID = process.env.AZURE_TENANT_ID;
const REDIRECT_URI = 'http://localhost:3000/callback';
const SCOPES = 'Tasks.Read Tasks.ReadWrite User.Read offline_access openid profile';

// Generate auth URL
const authUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/authorize?` +
  `client_id=${CLIENT_ID}&response_type=code&redirect_uri=${encodeURIComponent(REDIRECT_URI)}` +
  `&scope=${encodeURIComponent(SCOPES)}&response_mode=query`;

console.log('\n=== OPEN THIS URL ===\n');
console.log(authUrl);
console.log('\n=====================\n');

// Exchange code for tokens
async function exchangeCode(code) {
  const tokenUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;

  // Manually construct body to avoid URLSearchParams encoding the ~ in client_secret
  // Only encode values that need it (redirect_uri, scope), NOT the client_secret
  const body = [
    `client_id=${CLIENT_ID}`,
    `client_secret=${CLIENT_SECRET}`,  // DO NOT encode - Azure expects it as-is
    `code=${code}`,
    `redirect_uri=${encodeURIComponent(REDIRECT_URI)}`,
    `grant_type=authorization_code`,
    `scope=${encodeURIComponent(SCOPES)}`
  ].join('&');

  console.log('Token request body (secret masked):');
  console.log(body.replace(CLIENT_SECRET, '***MASKED***'));

  return new Promise((resolve, reject) => {
    const req = https.request(tokenUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Content-Length': Buffer.byteLength(body)
      }
    }, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        try {
          resolve(JSON.parse(data));
        } catch (e) {
          reject(new Error(data));
        }
      });
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

// Start server
const server = http.createServer(async (req, res) => {
  if (req.url?.startsWith('/callback')) {
    const url = new URL(req.url, 'http://localhost:3000');
    const code = url.searchParams.get('code');

    if (code) {
      console.log('Got code, exchanging for tokens...');
      try {
        const tokens = await exchangeCode(code);

        if (tokens.error) {
          console.error('Error:', tokens.error_description);
          res.end('<h1>Error: ' + tokens.error_description + '</h1>');
          return;
        }

        await fs.writeFile('.tokens.json', JSON.stringify({
          accessToken: tokens.access_token,
          refreshToken: tokens.refresh_token,
          expiresAt: Date.now() + tokens.expires_in * 1000  // Numeric timestamp for FileTokenManager
        }, null, 2));

        console.log('\n✅ SUCCESS! Tokens saved to .tokens.json\n');
        res.end('<h1>Success! Tokens saved. You can close this window.</h1>');
        setTimeout(() => process.exit(0), 1000);
      } catch (err) {
        console.error('Exchange failed:', err.message);
        res.end('<h1>Error: ' + err.message + '</h1>');
      }
    } else {
      const error = url.searchParams.get('error_description') || url.searchParams.get('error');
      console.error('Auth error:', error);
      res.end('<h1>Error: ' + error + '</h1>');
    }
  }
});

server.listen(3000, () => console.log('Waiting on port 3000...'));
