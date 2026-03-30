#!/usr/bin/env node
/**
 * M365 Email Safe-Write MCP Server
 *
 * A focused MCP server that provides only the low-risk write tool:
 * mark-as-read.  This is the "safe write" subset — it can update
 * email read status but cannot send, draft, delete, or move emails.
 *
 * Use this when you want the LLM to manage read/unread status
 * without any higher-risk write capabilities.
 */
const { createServer } = require('./server');
const { authTools } = require('./auth');
const { emailSafeWriteTools } = require('./email');

createServer('m365-email-safewrite', [
  ...authTools,
  ...emailSafeWriteTools
]);
