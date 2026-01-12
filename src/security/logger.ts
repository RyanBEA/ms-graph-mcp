import winston from 'winston';

// Sensitive fields that should never appear in logs
const sensitiveFields = [
  'access_token',
  'refresh_token',
  'client_secret',
  'password',
  'authorization',
  'cookie',
  'token',
  'secret',
  'key',
  'apikey',
  'api_key',
];

// Export for testing
export function redactSensitiveData(obj: any): any {
  if (typeof obj !== 'object' || obj === null) return obj;

  const redacted = { ...obj };

  for (const key in redacted) {
    if (Object.prototype.hasOwnProperty.call(redacted, key)) {
      // Check if key contains any sensitive field name (case-insensitive)
      if (sensitiveFields.some(field =>
        key.toLowerCase().includes(field.toLowerCase())
      )) {
        redacted[key] = '[REDACTED]';
      } else if (typeof redacted[key] === 'object') {
        redacted[key] = redactSensitiveData(redacted[key]);
      }
    }
  }
  return redacted;
}

// Redact sensitive fields from logs
const redactFormat = winston.format((info) => {
  return redactSensitiveData({ ...info });
});

export const logger = winston.createLogger({
  level: process.env.LOG_LEVEL || 'info',
  format: winston.format.combine(
    winston.format.timestamp(),
    redactFormat(),
    winston.format.errors({ stack: true }),
    winston.format.json()
  ),
  transports: [
    new winston.transports.Console({
      format: winston.format.combine(
        winston.format.colorize(),
        winston.format.simple()
      )
    }),
    new winston.transports.File({
      filename: 'logs/error.log',
      level: 'error'
    }),
    new winston.transports.File({
      filename: 'logs/combined.log'
    })
  ]
});

// Don't log in test environment
if (process.env.NODE_ENV === 'test') {
  logger.silent = true;
}
