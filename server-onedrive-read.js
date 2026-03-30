#!/usr/bin/env node
/**
 * M365 OneDrive Read-Only MCP Server
 *
 * A focused MCP server that provides only read-only OneDrive tools:
 * listing, searching, and downloading files.
 *
 * Use this when you want to give the LLM read access to OneDrive
 * without any ability to upload, share, create folders, or delete items.
 */
const { createServer } = require('./server');
const { authTools } = require('./auth');
const { onedriveReadTools } = require('./onedrive');

createServer('m365-onedrive-read', [
  ...authTools,
  ...onedriveReadTools
]);
