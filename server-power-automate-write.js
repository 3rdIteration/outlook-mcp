#!/usr/bin/env node
/**
 * M365 Power Automate Write-Only MCP Server
 *
 * A focused MCP server that provides only write Power Automate tools:
 * running and toggling flows.
 *
 * Use this when you want to give the LLM write access to Power Automate
 * without the read/list capabilities.
 */
const { createServer } = require('./server');
const { authTools } = require('./auth');
const { powerAutomateWriteTools } = require('./power-automate');

createServer('m365-power-automate-write', [
  ...authTools,
  ...powerAutomateWriteTools
]);
