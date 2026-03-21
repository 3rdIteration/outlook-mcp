#!/usr/bin/env node
/**
 * HTTP/SSE MCP Server — for use with OpenWebUI and other HTTP-based MCP clients.
 *
 * Exposes all M365 tools over two transport flavours:
 *   - POST /mcp   Streamable HTTP (MCP spec 2025-03-26, recommended for OpenWebUI ≥ 0.6.6)
 *   - GET  /sse   Legacy SSE stream  (MCP spec 2024-11-05, for older clients)
 *   - POST /messages  Legacy SSE message endpoint (paired with /sse)
 *
 * OAuth authentication routes are also served on this port so a separate
 * auth-server process is not required:
 *   - GET  /auth              Start Microsoft OAuth flow
 *   - GET  /auth/callback     OAuth redirect handler
 *   - GET  /token-status      Check current authentication status
 *
 * Configuration (environment variables):
 *   MCP_HTTP_PORT   Port to listen on          (default: 3001)
 *   MCP_HTTP_HOST   Bind address               (default: 0.0.0.0)
 *   MS_CLIENT_ID    Microsoft app client ID    (required)
 *   MS_CLIENT_SECRET Microsoft app client secret (required)
 *   MS_TENANT_ID    Azure tenant ID            (default: common)
 *
 * Usage:
 *   node sse-server.js
 *   npm run start:http
 */
require('dotenv').config();

const express = require('express');

const { StreamableHTTPServerTransport } = require('@modelcontextprotocol/sdk/server/streamableHttp.js');
const { SSEServerTransport } = require('@modelcontextprotocol/sdk/server/sse.js');

const config = require('./config');
const { createMcpServer } = require('./server');
const TokenStorage = require('./auth/token-storage');
const { setupOAuthRoutes, createAuthConfig } = require('./auth/oauth-server');

// Import all tools (same set as index.js)
const { authTools } = require('./auth');
const { calendarTools } = require('./calendar');
const { emailTools } = require('./email');
const { folderTools } = require('./folder');
const { rulesTools } = require('./rules');
const { onedriveTools } = require('./onedrive');
const { powerAutomateTools } = require('./power-automate');

const ALL_TOOLS = [
  ...authTools,
  ...calendarTools,
  ...emailTools,
  ...folderTools,
  ...rulesTools,
  ...onedriveTools,
  ...powerAutomateTools,
];

// Single shared TokenStorage instance for auth checks
const tokenStorage = new TokenStorage();

/**
 * Middleware: require a valid Microsoft access token before handling MCP requests.
 * Returns 401 with a JSON error body (and a link to /auth) when no token is found.
 */
async function requireAuth(req, res, next) {
  try {
    const token = await tokenStorage.getValidAccessToken();
    if (!token) {
      res.status(401).json({
        error: 'Authentication required.',
        message: 'No valid Microsoft access token found. Visit /auth to authenticate.',
        authUrl: '/auth',
      });
      return;
    }
    next();
  } catch (err) {
    res.status(401).json({
      error: 'Authentication required.',
      message: 'Failed to validate access token. Visit /auth to re-authenticate.',
      authUrl: '/auth',
    });
  }
}

// ---------------------------------------------------------------------------
// Express application
// ---------------------------------------------------------------------------
const app = express();

// Parse JSON bodies (needed for /mcp and /messages)
app.use(express.json());

// CORS — OpenWebUI and other browser clients send cross-origin requests.
// Restrict the allowed origin via the MCP_ALLOWED_ORIGIN env var in production.
app.use((req, res, next) => {
  const allowedOrigin = process.env.MCP_ALLOWED_ORIGIN || req.headers.origin || '*';
  res.setHeader('Access-Control-Allow-Origin', allowedOrigin);
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, DELETE, OPTIONS');
  res.setHeader(
    'Access-Control-Allow-Headers',
    'Content-Type, Authorization, Accept, Mcp-Session-Id, Last-Event-Id'
  );
  res.setHeader('Access-Control-Expose-Headers', 'Mcp-Session-Id');
  if (req.method === 'OPTIONS') {
    res.status(204).end();
    return;
  }
  next();
});

// ---------------------------------------------------------------------------
// Status / info page
// ---------------------------------------------------------------------------
app.get('/', (_req, res) => {
  res.type('text/html').send(`<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>M365 MCP SSE/Streaming Server</title>
  <style>
    body { font-family: Arial, sans-serif; max-width: 700px; margin: 50px auto; padding: 20px; }
    code { background: #f4f4f4; padding: 2px 6px; border-radius: 3px; }
    h2 { margin-top: 1.5em; }
  </style>
</head>
<body>
  <h1>MCP SSE/Streaming Server</h1>
  <p>This server provides MCP access to Microsoft 365 services (Outlook, OneDrive, Power Automate).</p>

  <h2>MCP Endpoints</h2>
  <ul>
    <li><strong>POST <code>/mcp</code></strong> — Streamable HTTP transport (recommended, OpenWebUI ≥ 0.6.6)</li>
    <li><strong>GET <code>/sse</code></strong> — Legacy SSE transport (older clients)</li>
    <li><strong>POST <code>/messages</code></strong> — Legacy SSE message channel</li>
  </ul>

  <h2>Authentication</h2>
  <ul>
    <li><strong>GET <code>/auth</code></strong> — Start Microsoft OAuth flow</li>
    <li><strong>GET <code>/auth/callback</code></strong> — OAuth redirect handler</li>
    <li><strong>GET <code>/token-status</code></strong> — Check authentication status</li>
  </ul>
</body>
</html>`);
});

// ---------------------------------------------------------------------------
// OAuth routes (authentication — no token required)
// ---------------------------------------------------------------------------
const authConfig = createAuthConfig('MS_');
setupOAuthRoutes(app, tokenStorage, authConfig, 'MS_');

// ---------------------------------------------------------------------------
// Streamable HTTP transport  (POST /mcp)
// Preferred by OpenWebUI ≥ 0.6.6.  Each request is handled statelessly:
// a fresh transport + MCP server instance is created per request.  This is
// the SDK-recommended pattern for stateless mode and suits the low-volume
// usage of a personal M365 assistant.
// ---------------------------------------------------------------------------
app.post('/mcp', requireAuth, async (req, res) => {
  const transport = new StreamableHTTPServerTransport({
    sessionIdGenerator: undefined, // stateless — no session tracking
  });
  const mcpServer = createMcpServer(config.SERVER_NAME, ALL_TOOLS);
  await mcpServer.connect(transport);
  await transport.handleRequest(req, res, req.body);
  res.on('finish', () => {
    mcpServer.close().catch(() => {});
  });
});

// Reject GET/DELETE on /mcp with a clear error
app.get('/mcp', (_req, res) => {
  res.status(405).json({ error: 'Use POST for MCP Streamable HTTP requests.' });
});
app.delete('/mcp', requireAuth, (_req, res) => {
  res.status(200).end();
});

// ---------------------------------------------------------------------------
// Legacy SSE transport  (GET /sse  +  POST /messages)
// Kept for backward-compatibility with clients that only support the older
// SSE-based MCP transport (spec 2024-11-05).
// ---------------------------------------------------------------------------

/** Active SSE sessions keyed by session ID. */
const sseSessions = {};

app.get('/sse', requireAuth, async (req, res) => {
  const transport = new SSEServerTransport('/messages', res);
  const sessionId = transport.sessionId;
  sseSessions[sessionId] = transport;

  const mcpServer = createMcpServer(config.SERVER_NAME, ALL_TOOLS);
  await mcpServer.connect(transport);

  res.on('close', () => {
    delete sseSessions[sessionId];
    mcpServer.close().catch(() => {});
  });
});

app.post('/messages', requireAuth, async (req, res) => {
  const sessionId = req.query.sessionId;
  const transport = sseSessions[sessionId];
  if (!transport) {
    res.status(400).json({ error: 'No active SSE session found for the provided sessionId.' });
    return;
  }
  await transport.handlePostMessage(req, res, req.body);
});

// ---------------------------------------------------------------------------
// Server start helper (exported for tests)
// ---------------------------------------------------------------------------
function startServer(port, host) {
  const listenPort = port || config.HTTP_PORT;
  const listenHost = host || config.HTTP_HOST;
  return app.listen(listenPort, listenHost, () => {
    console.error(`MCP HTTP server listening on http://${listenHost}:${listenPort}`);
    console.error(`  Streamable HTTP : http://${listenHost}:${listenPort}/mcp`);
    console.error(`  Legacy SSE      : http://${listenHost}:${listenPort}/sse`);
    console.error(`  Auth            : http://${listenHost}:${listenPort}/auth`);
  });
}

// Only start listening when run as a standalone script
if (require.main === module) {
  startServer();
}

module.exports = { app, startServer };
