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
 * ⚠️  SECURITY — READ BEFORE EXPOSING TO A NETWORK ⚠️
 *
 * By default the server binds to 127.0.0.1 (localhost only). HTTP on the
 * loopback interface is safe because traffic never leaves the machine.
 *
 * If you set MCP_HTTP_HOST=0.0.0.0 (or any non-localhost address) the server
 * becomes reachable over the network. In that case you MUST put a TLS-
 * terminating reverse proxy (nginx, Caddy, Traefik …) in front of it.
 * Running plain HTTP on a network interface exposes:
 *   • All Microsoft 365 data (emails, calendar, files) in plaintext.
 *   • The OAuth authorization code in the /auth/callback URL, which can be
 *     intercepted to steal tokens. (Microsoft also enforces HTTPS for
 *     non-localhost redirect URIs in Azure app registrations.)
 *
 * Configuration (environment variables):
 *   MCP_HTTP_PORT        Port to listen on          (default: 3001)
 *   MCP_HTTP_HOST        Bind address               (default: 127.0.0.1)
 *   MCP_ALLOWED_ORIGIN   Allowed CORS origin        (required for non-localhost)
 *   MS_CLIENT_ID         Microsoft app client ID    (required)
 *   MS_CLIENT_SECRET     Microsoft app client secret (required)
 *   MS_TENANT_ID         Azure tenant ID            (default: common)
 *   MS_REDIRECT_URI      OAuth callback URL         (default: http://localhost:3001/auth/callback)
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

/** Addresses considered local-only — HTTP is safe when bound to these. */
const LOCALHOST_ADDRESSES = new Set(['127.0.0.1', 'localhost', '::1']);

/** Returns true when the server is only reachable over the loopback interface. */
function isLocalhostBinding(host) {
  return LOCALHOST_ADDRESSES.has(host);
}

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
// When the server is bound to localhost, we echo the request origin (safe,
// because only local processes can connect). When bound to a network
// interface, MCP_ALLOWED_ORIGIN MUST be set explicitly; serving a wildcard
// '*' over HTTP on a network interface would let any web page on the LAN
// make authenticated MCP requests on behalf of the user.
app.use((req, res, next) => {
  let allowedOrigin;
  if (process.env.MCP_ALLOWED_ORIGIN) {
    allowedOrigin = process.env.MCP_ALLOWED_ORIGIN;
  } else if (isLocalhostBinding(config.HTTP_HOST)) {
    // Localhost binding: echo the request origin (or wildcard as fallback).
    allowedOrigin = req.headers.origin || '*';
  } else {
    // Network binding without explicit origin: echo origin only.
    // This is NOT safe over plain HTTP — use HTTPS + MCP_ALLOWED_ORIGIN.
    allowedOrigin = req.headers.origin || 'null';
  }
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
    mcpServer.close().catch((err) => {
      console.error('Error closing MCP server after /mcp request:', err.message);
    });
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
    mcpServer.close().catch((err) => {
      console.error('Error closing MCP server after SSE session end:', err.message);
    });
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

  // Warn loudly when the server is reachable over the network without TLS.
  if (!isLocalhostBinding(listenHost)) {
    console.error('');
    console.error('╔══════════════════════════════════════════════════════════════╗');
    console.error('║  ⚠️  SECURITY WARNING — PLAIN HTTP ON NETWORK INTERFACE  ⚠️  ║');
    console.error('╠══════════════════════════════════════════════════════════════╣');
    console.error(`║  Binding to ${String(listenHost).substring(0, 48).padEnd(48)}  ║`);
    console.error('║                                                              ║');
    console.error('║  HTTP traffic is NOT encrypted. This means:                 ║');
    console.error('║  • Email / calendar / file data travels in plaintext.       ║');
    console.error('║  • The OAuth auth code in /auth/callback can be captured.   ║');
    console.error('║                                                              ║');
    console.error('║  You MUST put a TLS-terminating reverse proxy (nginx,       ║');
    console.error('║  Caddy, Traefik …) in front of this server.                 ║');
    console.error('║  See the README for a recommended nginx/Caddy setup.        ║');
    console.error('╚══════════════════════════════════════════════════════════════╝');
    console.error('');
  }

  return app.listen(listenPort, listenHost, () => {
    console.error(`MCP HTTP server listening on http://${listenHost}:${listenPort}`);
    console.error(`  Streamable HTTP : http://${listenHost}:${listenPort}/mcp`);
    console.error(`  Legacy SSE      : http://${listenHost}:${listenPort}/sse`);
    console.error(`  Auth            : http://${listenHost}:${listenPort}/auth`);
    if (isLocalhostBinding(listenHost)) {
      console.error('  (localhost-only — safe for plain HTTP)');
    }
  });
}

// Only start listening when run as a standalone script
if (require.main === module) {
  startServer();
}

module.exports = { app, startServer };
