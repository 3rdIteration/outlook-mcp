#!/usr/bin/env node
/**
 * M365 OneDrive Write-Only MCP Server
 *
 * A focused MCP server that provides only write OneDrive tools:
 * uploading, sharing, creating folders, and deleting items.
 *
 * Use this when you want to give the LLM write access to OneDrive
 * without the read/list/search/download capabilities.
 */
const { createServer } = require('./server');
const { authTools } = require('./auth');
const { onedriveWriteTools } = require('./onedrive');

createServer('m365-onedrive-write', [
  ...authTools,
  ...onedriveWriteTools
]);
