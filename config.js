/**
 * Configuration for Outlook MCP Server
 */
const path = require('path');
const os = require('os');

// Ensure we have a home directory path even if process.env.HOME is undefined
const homeDir = process.env.HOME || process.env.USERPROFILE || os.homedir() || '/tmp';

module.exports = {
  // Server information
  SERVER_NAME: "m365-assistant",
  SERVER_VERSION: "2.0.0",
  
  // Test mode setting
  USE_TEST_MODE: process.env.USE_TEST_MODE === 'true',
  
  // Authentication configuration
  AUTH_CONFIG: {
    clientId: process.env.OUTLOOK_CLIENT_ID || '',
    clientSecret: process.env.OUTLOOK_CLIENT_SECRET || '',
    redirectUri: 'http://localhost:3333/auth/callback',
    scopes: ['Mail.Read', 'Mail.ReadWrite', 'Mail.Send', 'User.Read', 'Calendars.Read', 'Calendars.ReadWrite', 'Files.Read', 'Files.ReadWrite'],
    tokenStorePath: path.join(homeDir, '.outlook-mcp-tokens.json'),
    authServerUrl: 'http://localhost:3333'
  },
  
  // Microsoft Graph API
  GRAPH_API_ENDPOINT: 'https://graph.microsoft.com/v1.0/',
  
  // Calendar constants
  CALENDAR_SELECT_FIELDS: 'id,subject,start,end,location,bodyPreview,isAllDay,recurrence,attendees',

  // Email constants
  EMAIL_SELECT_FIELDS: 'id,subject,from,toRecipients,ccRecipients,receivedDateTime,bodyPreview,hasAttachments,importance,isRead',
  EMAIL_DETAIL_FIELDS: 'id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,bodyPreview,body,hasAttachments,importance,isRead,internetMessageHeaders',
  
  // Calendar constants
  CALENDAR_SELECT_FIELDS: 'id,subject,bodyPreview,start,end,location,organizer,attendees,isAllDay,isCancelled',
  
  // Pagination
  DEFAULT_PAGE_SIZE: 25,
  MAX_RESULT_COUNT: 50,

  // Tools listing pagination - number of tools per page for tools/list.
  // Set MCP_TOOLS_PAGE_SIZE env var to control page size.
  // 0 (default) = return all tools in one response (no pagination).
  TOOLS_PAGE_SIZE: parseInt(process.env.MCP_TOOLS_PAGE_SIZE, 10) || 0,

  // Timezone
  DEFAULT_TIMEZONE: "Central European Standard Time",

  // OneDrive constants
  ONEDRIVE_SELECT_FIELDS: 'id,name,size,lastModifiedDateTime,webUrl,folder,file,parentReference',
  ONEDRIVE_UPLOAD_THRESHOLD: 4 * 1024 * 1024, // 4MB - files larger than this need chunked upload

  // Power Automate / Flow constants
  FLOW_API_ENDPOINT: 'https://api.flow.microsoft.com',
  FLOW_SCOPE: 'https://service.flow.microsoft.com/.default',

  // HTTP/SSE server (for OpenWebUI and other HTTP-based MCP clients)
  // Set MCP_HTTP_PORT env var to change port (default: 3001).
  // Set MCP_HTTP_HOST env var to change bind address.
  // Default is 127.0.0.1 (localhost only) — HTTP is safe on the loopback interface.
  // Set to 0.0.0.0 only when using a TLS-terminating reverse proxy in front of this server.
  HTTP_PORT: parseInt(process.env.MCP_HTTP_PORT, 10) || 3001,
  HTTP_HOST: process.env.MCP_HTTP_HOST || '127.0.0.1',
};
