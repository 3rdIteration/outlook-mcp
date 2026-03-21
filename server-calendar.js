#!/usr/bin/env node
/**
 * M365 Calendar MCP Server
 *
 * A focused MCP server that provides only calendar-related tools.
 *
 * Use this instead of index.js when you only need calendar functionality,
 * to reduce context window usage.
 */
const { createServer } = require('./server');
const { authTools } = require('./auth');
const { calendarTools } = require('./calendar');

createServer('m365-calendar', [
  ...authTools,
  ...calendarTools
]);
