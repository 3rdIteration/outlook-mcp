#!/usr/bin/env node
/**
 * M365 Email MCP Server
 *
 * A focused MCP server that provides only email-related tools:
 * email, folders, and rules.
 *
 * Use this instead of index.js when you only need email functionality,
 * to reduce context window usage.
 */
const { createServer } = require('./server');
const { authTools } = require('./auth');
const { emailTools } = require('./email');
const { folderTools } = require('./folder');
const { rulesTools } = require('./rules');

createServer('m365-email', [
  ...authTools,
  ...emailTools,
  ...folderTools,
  ...rulesTools
]);
