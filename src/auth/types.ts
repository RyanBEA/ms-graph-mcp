/**
 * Token set containing OAuth tokens and metadata
 */
export interface TokenSet {
  accessToken: string;
  refreshToken: string;
  expiresAt: number; // Unix timestamp in milliseconds
  scope: string;
}

/**
 * Interface for secure token storage implementations
 * Supports both 1Password SDK and keytar (Windows Credential Manager)
 */
export interface ITokenManager {
  /**
   * Store a new token set securely
   * @param tokens - Complete OAuth token set
   */
  storeTokens(tokens: TokenSet): Promise<void>;

  /**
   * Retrieve stored tokens
   * @returns Token set if exists
   * @throws TokenStorageError if tokens don't exist or are corrupted
   */
  getTokens(): Promise<TokenSet>;

  /**
   * Check if valid (non-expired) tokens exist
   * @returns true if tokens exist and are not expired
   */
  hasValidTokens(): Promise<boolean>;

  /**
   * Clear stored tokens (logout)
   */
  clearTokens(): Promise<void>;

  /**
   * Update only the access token (after refresh)
   * @param accessToken - New access token
   * @param expiresAt - New expiration timestamp
   */
  updateAccessToken(accessToken: string, expiresAt: number): Promise<void>;
}
