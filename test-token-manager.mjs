import { KeytarTokenManager } from './dist/auth/keytar-token-manager.js';

async function testTokenManager() {
  console.log('Testing KeytarTokenManager...\n');

  try {
    const manager = new KeytarTokenManager();

    // Test storing tokens
    console.log('1. Storing test tokens...');
    const testTokens = {
      accessToken: 'test-access-token-123',
      refreshToken: 'test-refresh-token-456',
      expiresAt: Date.now() + 3600000,
      scope: 'Tasks.Read User.Read'
    };

    await manager.storeTokens(testTokens);
    console.log('   ✅ Tokens stored');

    // Test retrieving tokens
    console.log('\n2. Retrieving tokens...');
    const retrieved = await manager.getTokens();
    console.log('   ✅ Tokens retrieved:', {
      accessToken: retrieved.accessToken.substring(0, 20) + '...',
      refreshToken: retrieved.refreshToken.substring(0, 20) + '...',
      expiresAt: retrieved.expiresAt,
      scope: retrieved.scope
    });

    // Test clearing tokens
    console.log('\n3. Clearing tokens...');
    await manager.clearTokens();
    console.log('   ✅ Tokens cleared');

    console.log('\n✅ All TokenManager tests passed!\n');

  } catch (error) {
    console.error('❌ TokenManager test failed:', error);
    console.error('Error details:', {
      name: error.name,
      message: error.message,
      stack: error.stack
    });
    process.exit(1);
  }
}

testTokenManager();
