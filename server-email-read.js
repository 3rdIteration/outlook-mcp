#!/usr/bin/env node
/**
 * M365 Email Read-Only MCP Server
 *
 * A focused MCP server that provides only read-only email-related tools:
 * listing/searching/reading emails, listing folders, and listing rules.
 *
 * Use this when you want to give the LLM read access to email
 * without any ability to send, draft, or modify emails.
 */
const { createServer } = require('./server');
const { authTools } = require('./auth');
const { emailReadTools } = require('./email');
const { folderReadTools } = require('./folder');
const { rulesReadTools } = require('./rules');

createServer('m365-email-read', [
  ...authTools,
  ...emailReadTools,
  ...folderReadTools,
  ...rulesReadTools
]);
