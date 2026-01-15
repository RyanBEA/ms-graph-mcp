import { z } from 'zod';
import dotenv from 'dotenv';

// Load environment variables
dotenv.config();

// Define configuration schema
const envSchema = z.object({
  // Azure
  AZURE_CLIENT_ID: z.string().min(1, 'Azure Client ID is required'),
  AZURE_TENANT_ID: z.string().min(1, 'Azure Tenant ID is required'),
  AZURE_CLIENT_SECRET: z.string().min(1, 'Azure Client Secret is required'),
  AZURE_REDIRECT_URI: z.string().url('Must be valid URL'),

  // Token Storage (default: msal for security; 'file' requires explicit opt-in)
  TOKEN_STORAGE: z.enum(['1password', 'keytar', 'msal', 'file']).default('msal'),
  OP_SERVICE_ACCOUNT_TOKEN: z.string().optional(),

  // Environment
  NODE_ENV: z.enum(['development', 'production', 'test']).default('development'),
  LOG_LEVEL: z.enum(['error', 'warn', 'info', 'debug']).default('info'),

  // Security
  RATE_LIMIT_PER_MINUTE: z.coerce.number().default(60),
  STATE_TIMEOUT_MINUTES: z.coerce.number().default(5),
});

export type Environment = z.infer<typeof envSchema>;

class ConfigurationError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'ConfigurationError';
  }
}

let config: Environment;

export function loadConfiguration(): Environment {
  try {
    config = envSchema.parse(process.env);

    // Additional validation
    if (config.TOKEN_STORAGE === '1password' && !config.OP_SERVICE_ACCOUNT_TOKEN) {
      throw new ConfigurationError(
        '1Password selected but OP_SERVICE_ACCOUNT_TOKEN not provided'
      );
    }

    return config;
  } catch (error) {
    if (error instanceof z.ZodError) {
      const issues = error.issues.map(i => `${i.path.join('.')}: ${i.message}`);
      throw new ConfigurationError(
        `Configuration validation failed:\n${issues.join('\n')}`
      );
    }
    throw error;
  }
}

export function getConfiguration(): Environment {
  if (!config) {
    throw new ConfigurationError('Configuration not loaded. Call loadConfiguration() first.');
  }
  return config;
}
