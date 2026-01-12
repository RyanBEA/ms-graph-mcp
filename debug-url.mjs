#!/usr/bin/env node
/**
 * Debug script to test URL construction without running full MCP server
 */

import { SecureGraphClient } from './dist/graph/client.js';
import { TokenRefresher } from './dist/auth/token-refresher.js';
import { FileTokenManager } from './dist/auth/file-token-manager.js';
import { ToDoService } from './dist/graph/todo-service.js';
import { loadConfiguration } from './dist/config/environment.js';

async function test() {
  console.log('=== Testing URL Construction ===\n');

  try {
    // Load configuration
    loadConfiguration();
    console.log('Configuration loaded\n');

    // Create token manager and refresher
    const tokenManager = new FileTokenManager();
    const tokenRefresher = new TokenRefresher(tokenManager);

    // Create graph client
    const graphClient = new SecureGraphClient(tokenRefresher);

    // Create todo service
    const todoService = new ToDoService(graphClient);

    console.log('Services initialized successfully');
    console.log('Attempting to fetch task lists...\n');

    // Try to get task lists - this should trigger the URL construction
    const lists = await todoService.getTaskLists();

    console.log('SUCCESS! Got', lists.length, 'task lists');
    lists.slice(0, 3).forEach(list => {
      console.log(' -', list.displayName);
    });

  } catch (error) {
    console.error('ERROR:', error.message);
    console.error('Stack:', error.stack);
  }
}

test();
