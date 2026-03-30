#!/usr/bin/env node
/**
 * M365 Calendar Write-Only MCP Server
 *
 * A focused MCP server that provides only write calendar tools:
 * creating, accepting, declining, cancelling, and deleting events.
 *
 * Use this when you want to give the LLM write access to your calendar
 * without the read/list capabilities.
 */
const { createServer } = require('./server');
const { authTools } = require('./auth');
const { calendarWriteTools } = require('./calendar');

createServer('m365-calendar-write', [
  ...authTools,
  ...calendarWriteTools
]);
