#!/usr/bin/env node
/**
 * M365 OneDrive MCP Server
 *
 * A focused MCP server that provides only OneDrive-related tools.
 *
 * Use this instead of index.js when you only need OneDrive functionality,
 * to reduce context window usage.
 */
const { createServer } = require('./server');
const { authTools } = require('./auth');
const { onedriveTools } = require('./onedrive');

createServer('m365-onedrive', [
  ...authTools,
  ...onedriveTools
]);
