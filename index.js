#!/usr/bin/env node
/**
 * M365 Assistant MCP Server - Main entry point (all tools)
 *
 * A Model Context Protocol server that provides access to
 * Microsoft 365 services (Outlook, OneDrive, Power Automate)
 * through the Microsoft Graph API and Flow API.
 *
 * This entry point loads ALL tools. For smaller context windows,
 * use the focused entry points instead:
 *   - server-email.js          (email, folders, rules)
 *   - server-calendar.js       (calendar)
 *   - server-onedrive.js       (OneDrive)
 *   - server-power-automate.js (Power Automate)
 */
const config = require('./config');
const { createServer } = require('./server');

// Import module tools
const { authTools } = require('./auth');
const { calendarTools } = require('./calendar');
const { emailTools } = require('./email');
const { folderTools } = require('./folder');
const { rulesTools } = require('./rules');
const { onedriveTools } = require('./onedrive');
const { powerAutomateTools } = require('./power-automate');

// Combine all tools and start server
createServer(config.SERVER_NAME, [
  ...authTools,
  ...calendarTools,
  ...emailTools,
  ...folderTools,
  ...rulesTools,
  ...onedriveTools,
  ...powerAutomateTools
]);
