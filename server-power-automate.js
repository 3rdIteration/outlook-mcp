#!/usr/bin/env node
/**
 * M365 Power Automate MCP Server
 *
 * A focused MCP server that provides only Power Automate-related tools.
 *
 * Use this instead of index.js when you only need Power Automate functionality,
 * to reduce context window usage.
 */
const { createServer } = require('./server');
const { authTools } = require('./auth');
const { powerAutomateTools } = require('./power-automate');

createServer('m365-power-automate', [
  ...authTools,
  ...powerAutomateTools
]);
