#!/usr/bin/env node
/**
 * MsGraph MCP Server - Entry Point
 * Secure Microsoft Graph API integration for Claude Code.
 */

// Load .env FIRST before any other imports so LOG_LEVEL is available
import dotenv from 'dotenv';
dotenv.config();

import { MsGraphMCPServer } from './mcp/server.js';
import { logger } from './security/logger.js';
import { loadConfiguration } from './config/environment.js';

/**
 * Main entry point.
 */
async function main() {
  try {
    // Force logger to debug level if set in environment
    // Need to update both the logger and all transports
    if (process.env.LOG_LEVEL) {
      logger.level = process.env.LOG_LEVEL;
      logger.transports.forEach(transport => {
        transport.level = process.env.LOG_LEVEL;
      });
    }

    logger.info('=== MsGraph MCP Server Starting ===');
    logger.debug('Logger configuration:', {
      LOG_LEVEL: process.env.LOG_LEVEL,
      loggerLevel: logger.level
    });

    // Load and validate configuration first
    loadConfiguration();
    logger.info('Configuration loaded successfully');

    const server = await MsGraphMCPServer.create();
    await server.start();

    // Keep process alive
    process.on('SIGINT', () => {
      logger.info('Received SIGINT, shutting down gracefully');
      process.exit(0);
    });

    process.on('SIGTERM', () => {
      logger.info('Received SIGTERM, shutting down gracefully');
      process.exit(0);
    });
  } catch (error) {
    logger.error('Fatal error starting server', {
      error: error instanceof Error ? error.message : 'Unknown error',
    });
    process.exit(1);
  }
}

// Start the server
main();
