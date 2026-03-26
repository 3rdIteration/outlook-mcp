#!/usr/bin/env node
/**
 * M365 Power Automate Read-Only MCP Server
 *
 * A focused MCP server that provides only read-only Power Automate tools:
 * listing environments, flows, and run history.
 *
 * Use this when you want to give the LLM read access to Power Automate
 * without any ability to run or toggle flows.
 */
const { createServer } = require('./server');
const { authTools } = require('./auth');
const { powerAutomateReadTools } = require('./power-automate');

createServer('m365-power-automate-read', [
  ...authTools,
  ...powerAutomateReadTools
]);
