/**
 * Input validators for security-sensitive operations.
 * Prevents injection attacks and validates user inputs.
 */

import { ValidationError } from './errors.js';
import { logger } from './logger.js';

/**
 * Allowed OData filter fields for Microsoft To Do API.
 * Whitelist approach prevents injection attacks.
 */
const ALLOWED_FILTER_FIELDS = [
  'status',
  'importance',
  'dueDateTime',
  'createdDateTime',
  'lastModifiedDateTime',
  'isReminderOn',
  'title',
] as const;

/**
 * Allowed OData filter operators.
 * Only comparison and logical operators, no functions.
 */
const ALLOWED_OPERATORS = [
  'eq',  // equals
  'ne',  // not equals
  'gt',  // greater than
  'ge',  // greater than or equal
  'lt',  // less than
  'le',  // less than or equal
  'and', // logical and
  'or',  // logical or
  'not', // logical not
] as const;

/**
 * Validate OData $filter query parameter.
 * Prevents injection by whitelisting allowed fields and operators.
 *
 * @param filter - OData filter string
 * @throws {ValidationError} If filter contains invalid fields or operators
 *
 * @example
 * validateODataFilter("status eq 'completed'") // Valid
 * validateODataFilter("status eq 'completed' and importance eq 'high'") // Valid
 * validateODataFilter("malicious_field eq 'value'") // Throws ValidationError
 */
export function validateODataFilter(filter: string): void {
  if (!filter || filter.trim().length === 0) {
    return; // Empty filters are allowed
  }

  const trimmedFilter = filter.trim();

  // Check for null bytes
  if (trimmedFilter.includes('\x00')) {
    logger.warn('Null byte detected in OData filter');
    throw new ValidationError('Filter contains null bytes');
  }

  // Check for suspicious patterns that might indicate injection attempts
  const suspiciousPatterns = [
    /[;<>{}[\]\\]/,           // Special characters
    /\b(drop|delete|update|insert|exec|script|eval|process|require)\b/i, // SQL/script/code keywords
    /<script/i,               // XSS attempts
    /javascript:/i,           // JavaScript protocol
    /on\w+\s*=/i,            // Event handlers
  ];

  for (const pattern of suspiciousPatterns) {
    if (pattern.test(trimmedFilter)) {
      logger.warn('Suspicious OData filter detected', {
        pattern: pattern.toString(),
        filterLength: trimmedFilter.length,
      });
      throw new ValidationError('Filter contains invalid characters or patterns');
    }
  }

  // Extract field names from filter
  // Support both simple fields (e.g., "status eq 'completed'")
  // and property navigation (e.g., "dueDateTime/dateTime lt '2024-01-01'")
  const fieldPattern = /(\w+)(?:\/\w+)?\s+(?:eq|ne|gt|ge|lt|le)/gi;
  const matches = [...trimmedFilter.matchAll(fieldPattern)];

  for (const match of matches) {
    const field = match[1]; // Capture the parent property (before /)
    if (!ALLOWED_FILTER_FIELDS.includes(field as any)) {
      logger.warn('Invalid OData filter field', {
        field,
        allowed: ALLOWED_FILTER_FIELDS,
      });
      throw new ValidationError(
        `Invalid filter field: ${field}. Allowed fields: ${ALLOWED_FILTER_FIELDS.join(', ')}`
      );
    }
  }

  // Validate operators
  const operatorPattern = /\s+(eq|ne|gt|ge|lt|le|and|or|not)\s+/gi;
  const operatorMatches = [...trimmedFilter.matchAll(operatorPattern)];

  for (const match of operatorMatches) {
    const operator = match[1].toLowerCase();
    if (!ALLOWED_OPERATORS.includes(operator as any)) {
      logger.warn('Invalid OData operator', {
        operator,
        allowed: ALLOWED_OPERATORS,
      });
      throw new ValidationError(
        `Invalid operator: ${operator}. Allowed operators: ${ALLOWED_OPERATORS.join(', ')}`
      );
    }
  }

  logger.debug('OData filter validated successfully', {
    filterLength: trimmedFilter.length,
  });
}

/**
 * Validate list ID format.
 * Microsoft Graph list IDs are alphanumeric with hyphens.
 *
 * @param listId - List identifier
 * @throws {ValidationError} If listId format is invalid
 */
export function validateListId(listId: string): void {
  if (!listId || typeof listId !== 'string') {
    throw new ValidationError('List ID is required');
  }

  const trimmedId = listId.trim();

  // Check for null bytes
  if (trimmedId.includes('\x00')) {
    logger.warn('Null byte detected in list ID');
    throw new ValidationError('List ID contains null bytes');
  }

  // Microsoft Graph IDs are base64-encoded strings (alphanumeric + - _ + =)
  const validIdPattern = /^[a-zA-Z0-9_\-+=]+$/;

  if (!validIdPattern.test(trimmedId)) {
    logger.warn('Invalid list ID format', {
      length: trimmedId.length,
    });
    throw new ValidationError('Invalid list ID format');
  }

  // Reasonable length check (Graph IDs are typically < 100 chars)
  if (trimmedId.length > 200) {
    logger.warn('List ID too long', {
      length: trimmedId.length,
    });
    throw new ValidationError('List ID exceeds maximum length');
  }
}

/**
 * Validate task ID format.
 * Same rules as list ID.
 */
export function validateTaskId(taskId: string): void {
  validateListId(taskId); // Same validation rules
}

// ============================================
// Planner validators
// ============================================

/**
 * Validate Planner plan ID format.
 * Same rules as list ID (Graph IDs are base64-ish).
 */
export function validatePlanId(planId: string): void {
  validateListId(planId); // Same validation rules
}

/**
 * Validate Planner bucket ID format.
 * Same rules as list ID.
 */
export function validateBucketId(bucketId: string): void {
  validateListId(bucketId); // Same validation rules
}

// ============================================
// Write operation validators
// ============================================

/**
 * Validate task title for write operations.
 * Prevents XSS, injection, and other malicious content.
 *
 * @param title - Task title to validate
 * @throws {ValidationError} If title is invalid
 */
export function validateTaskTitle(title: string): void {
  if (!title || typeof title !== 'string') {
    throw new ValidationError('Task title is required');
  }

  const trimmed = title.trim();

  if (trimmed.length === 0) {
    throw new ValidationError('Task title cannot be empty');
  }

  // Microsoft Graph API limit for task titles
  if (trimmed.length > 400) {
    logger.warn('Task title too long', { length: trimmed.length });
    throw new ValidationError('Task title exceeds maximum length (400 characters)');
  }

  // Check for null bytes
  if (trimmed.includes('\x00')) {
    logger.warn('Null byte detected in task title');
    throw new ValidationError('Task title contains null bytes');
  }

  // Check for suspicious patterns (XSS, script injection)
  const suspiciousPatterns = [
    /<script/i,
    /javascript:/i,
    /on\w+\s*=/i,  // Event handlers like onclick=
  ];

  for (const pattern of suspiciousPatterns) {
    if (pattern.test(trimmed)) {
      logger.warn('Suspicious task title detected', {
        pattern: pattern.toString(),
      });
      throw new ValidationError('Task title contains invalid patterns');
    }
  }

  logger.debug('Task title validated', { length: trimmed.length });
}

/**
 * Validate due date for write operations.
 * Accepts ISO 8601 date format.
 *
 * @param dateStr - Date string to validate
 * @returns Validated Date object
 * @throws {ValidationError} If date is invalid
 */
export function validateDueDate(dateStr: string): Date {
  if (!dateStr || typeof dateStr !== 'string') {
    throw new ValidationError('Due date is required');
  }

  const trimmed = dateStr.trim();

  // Check for null bytes
  if (trimmed.includes('\x00')) {
    logger.warn('Null byte detected in due date');
    throw new ValidationError('Due date contains null bytes');
  }

  const date = new Date(trimmed);

  if (isNaN(date.getTime())) {
    logger.warn('Invalid due date format', { providedLength: trimmed.length });
    throw new ValidationError('Invalid due date format. Use ISO format (YYYY-MM-DD)');
  }

  // Allow dates up to 5 years in future
  const maxDate = new Date();
  maxDate.setFullYear(maxDate.getFullYear() + 5);

  if (date > maxDate) {
    logger.warn('Due date too far in future', { date: date.toISOString() });
    throw new ValidationError('Due date cannot be more than 5 years in the future');
  }

  logger.debug('Due date validated', { date: date.toISOString() });

  return date;
}

/**
 * Validate priority value for To Do tasks.
 *
 * @param priority - Priority string
 * @throws {ValidationError} If priority is invalid
 */
export function validatePriority(priority: string): void {
  const allowed = ['low', 'normal', 'high'];
  if (!allowed.includes(priority)) {
    throw new ValidationError(`Invalid priority: ${priority}. Allowed: ${allowed.join(', ')}`);
  }
}

/**
 * Validate Planner percent complete value.
 * Planner only accepts 0, 50, or 100.
 *
 * @param percent - Percent complete value
 * @throws {ValidationError} If value is invalid
 */
export function validatePercentComplete(percent: number): void {
  const allowed = [0, 50, 100];
  if (!allowed.includes(percent)) {
    throw new ValidationError(`Invalid percent complete: ${percent}. Allowed: 0, 50, 100`);
  }
}

/**
 * Validate search query.
 * Prevents excessively long queries and injection attempts.
 *
 * @param query - Search query string
 * @throws {ValidationError} If query is invalid
 */
export function validateSearchQuery(query: string): void {
  if (!query || typeof query !== 'string') {
    throw new ValidationError('Search query is required');
  }

  const trimmedQuery = query.trim();

  if (trimmedQuery.length === 0) {
    throw new ValidationError('Search query cannot be empty');
  }

  if (trimmedQuery.length > 1000) {
    logger.warn('Search query too long', {
      length: trimmedQuery.length,
    });
    throw new ValidationError('Search query exceeds maximum length (1000 characters)');
  }

  // Check for null bytes
  if (trimmedQuery.includes('\x00')) {
    logger.warn('Null byte detected in search query');
    throw new ValidationError('Search query contains null bytes');
  }

  // Check for suspicious patterns including command injection
  const suspiciousPatterns = [
    /<script/i,
    /javascript:/i,
    /on\w+\s*=/i,
    /[;&|`$()]/,  // Command injection characters
  ];

  for (const pattern of suspiciousPatterns) {
    if (pattern.test(trimmedQuery)) {
      logger.warn('Suspicious search query detected', {
        pattern: pattern.toString(),
      });
      throw new ValidationError('Search query contains invalid patterns');
    }
  }

  logger.debug('Search query validated', {
    length: trimmedQuery.length,
  });
}

// ============================================
// Calendar validators
// ============================================

/**
 * Validate date string format.
 * Accepts ISO 8601 date format (YYYY-MM-DD or full ISO datetime).
 *
 * @param dateStr - Date string to validate
 * @param fieldName - Name of the field for error messages
 * @throws {ValidationError} If date format is invalid
 */
export function validateDateString(dateStr: string, fieldName: string): Date {
  if (!dateStr || typeof dateStr !== 'string') {
    throw new ValidationError(`${fieldName} is required`);
  }

  const trimmed = dateStr.trim();

  // Check for null bytes
  if (trimmed.includes('\x00')) {
    logger.warn('Null byte detected in date string');
    throw new ValidationError(`${fieldName} contains null bytes`);
  }

  // Try to parse the date
  const date = new Date(trimmed);

  if (isNaN(date.getTime())) {
    logger.warn('Invalid date format', {
      fieldName,
      providedLength: trimmed.length,
    });
    throw new ValidationError(`${fieldName} is not a valid date format. Use ISO format (e.g., 2025-01-15)`);
  }

  // Reasonable date range check (not too far in past or future)
  const now = new Date();
  const minDate = new Date();
  minDate.setFullYear(now.getFullYear() - 5); // 5 years in the past
  const maxDate = new Date();
  maxDate.setFullYear(now.getFullYear() + 5); // 5 years in the future

  if (date < minDate || date > maxDate) {
    logger.warn('Date out of reasonable range', {
      fieldName,
      date: date.toISOString(),
    });
    throw new ValidationError(`${fieldName} must be within 5 years of today`);
  }

  logger.debug('Date validated', {
    fieldName,
    date: date.toISOString(),
  });

  return date;
}

/**
 * Validate a date range for calendar queries.
 *
 * @param startDate - Start date
 * @param endDate - End date
 * @throws {ValidationError} If date range is invalid
 */
export function validateDateRange(startDate: Date, endDate: Date): void {
  if (endDate <= startDate) {
    throw new ValidationError('End date must be after start date');
  }

  // Max range of 365 days
  const daysDiff = (endDate.getTime() - startDate.getTime()) / (1000 * 60 * 60 * 24);
  if (daysDiff > 365) {
    logger.warn('Date range too large', { daysDiff });
    throw new ValidationError('Date range cannot exceed 365 days');
  }

  logger.debug('Date range validated', {
    startDate: startDate.toISOString(),
    endDate: endDate.toISOString(),
    daysDiff,
  });
}
