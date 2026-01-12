import { ITokenManager } from './types.js';

/**
 * Handles token access with automatic refresh.
 *
 * When using MsalTokenManager, refresh is handled internally by MSAL's
 * acquireTokenSilent(). This class provides a consistent interface.
 */
export class TokenRefresher {
  private tokenManager: ITokenManager;

  constructor(tokenManager: ITokenManager) {
    this.tokenManager = tokenManager;
  }

  /**
   * Get a valid access token, refreshing if necessary.
   *
   * Note: When using MsalTokenManager, the refresh is handled internally
   * by MSAL's acquireTokenSilent(). We simply delegate to getTokens().
   *
   * @returns Valid access token
   * @throws AuthenticationError if tokens not available
   */
  async getValidAccessToken(): Promise<string> {
    // MsalTokenManager.getTokens() calls acquireTokenSilent() which
    // automatically handles token refresh using MSAL's internal cache.
    const tokens = await this.tokenManager.getTokens();
    return tokens.accessToken;
  }
}
