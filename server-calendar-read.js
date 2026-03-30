#!/usr/bin/env node
/**
 * M365 Calendar Read-Only MCP Server
 *
 * A focused MCP server that provides only read-only calendar tools:
 * listing events.
 *
 * Use this when you want to give the LLM read access to your calendar
 * without any ability to create, accept, decline, or delete events.
 */
const { createServer } = require('./server');
const { authTools } = require('./auth');
const { calendarReadTools } = require('./calendar');

createServer('m365-calendar-read', [
  ...authTools,
  ...calendarReadTools
]);
