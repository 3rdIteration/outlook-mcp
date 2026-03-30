#!/usr/bin/env node
/**
 * M365 Email Write-Only MCP Server
 *
 * A focused MCP server that provides only write email-related tools:
 * sending, drafting, marking as read, creating folders, moving emails,
 * and managing inbox rules.
 *
 * Use this when you want to give the LLM write access to email
 * without the read/search/list capabilities.
 */
const { createServer } = require('./server');
const { authTools } = require('./auth');
const { emailWriteTools } = require('./email');
const { folderWriteTools } = require('./folder');
const { rulesWriteTools } = require('./rules');

createServer('m365-email-write', [
  ...authTools,
  ...emailWriteTools,
  ...folderWriteTools,
  ...rulesWriteTools
]);
