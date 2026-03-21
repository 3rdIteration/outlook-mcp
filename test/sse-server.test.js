/**
 * Tests for sse-server.js — HTTP/SSE MCP server (for OpenWebUI compatibility)
 */
const request = require('supertest');

// ---------------------------------------------------------------------------
// Mock module-level dependencies before requiring sse-server
// ---------------------------------------------------------------------------

// Prevent SIGTERM handler from accumulating during test runs
jest.mock('../server', () => {
  const actual = jest.requireActual('../server');
  return actual;
});

jest.mock('../auth/token-storage');
jest.mock('../auth/oauth-server');

const TokenStorage = require('../auth/token-storage');
const { setupOAuthRoutes, createAuthConfig } = require('../auth/oauth-server');

// Provide a consistent mock TokenStorage instance
const mockTokenStorageInstance = {
  getValidAccessToken: jest.fn(),
};
TokenStorage.mockImplementation(() => mockTokenStorageInstance);

// setupOAuthRoutes: register a minimal /token-status route so we can test it
setupOAuthRoutes.mockImplementation((app) => {
  app.get('/token-status', (_req, res) => res.json({ status: 'mocked' }));
});
createAuthConfig.mockReturnValue({});

// Now load the server AFTER mocks are in place
const { app } = require('../sse-server');

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

beforeEach(() => {
  jest.clearAllMocks();
});

// --- GET / ---------------------------------------------------------------

describe('GET /', () => {
  it('returns 200 with HTML status page', async () => {
    const res = await request(app).get('/');
    expect(res.status).toBe(200);
    expect(res.type).toMatch(/html/);
    expect(res.text).toContain('MCP SSE/Streaming Server');
  });
});

// --- GET /token-status (OAuth route registered by setupOAuthRoutes) -------

describe('GET /token-status', () => {
  it('returns the mocked token status', async () => {
    const res = await request(app).get('/token-status');
    expect(res.status).toBe(200);
    expect(res.body.status).toBe('mocked');
  });
});

// --- POST /mcp -----------------------------------------------------------

describe('POST /mcp', () => {
  it('returns 401 when no valid token exists', async () => {
    mockTokenStorageInstance.getValidAccessToken.mockResolvedValue(null);
    const res = await request(app).post('/mcp').send({ test: true });
    expect(res.status).toBe(401);
    expect(res.body.error).toBe('Authentication required.');
    expect(res.body.authUrl).toBe('/auth');
  });

  it('returns 401 when token check throws', async () => {
    mockTokenStorageInstance.getValidAccessToken.mockRejectedValue(new Error('disk error'));
    const res = await request(app).post('/mcp').send({});
    expect(res.status).toBe(401);
    expect(res.body.error).toBe('Authentication required.');
  });
});

// --- GET /mcp (wrong method) ---------------------------------------------

describe('GET /mcp', () => {
  it('returns 405 Method Not Allowed', async () => {
    const res = await request(app).get('/mcp');
    expect(res.status).toBe(405);
    expect(res.body.error).toMatch(/POST/);
  });
});

// --- GET /sse ------------------------------------------------------------

describe('GET /sse', () => {
  it('returns 401 when no valid token exists', async () => {
    mockTokenStorageInstance.getValidAccessToken.mockResolvedValue(null);
    const res = await request(app).get('/sse');
    expect(res.status).toBe(401);
    expect(res.body.error).toBe('Authentication required.');
  });
});

// --- POST /messages ------------------------------------------------------

describe('POST /messages', () => {
  it('returns 400 when sessionId is not provided', async () => {
    const res = await request(app).post('/messages').send({});
    expect(res.status).toBe(400);
    expect(res.body.error).toMatch(/session/i);
  });

  it('returns 400 for an unknown sessionId', async () => {
    const res = await request(app)
      .post('/messages')
      .query({ sessionId: 'nonexistent-session-id' })
      .send({});
    expect(res.status).toBe(400);
    expect(res.body.error).toMatch(/session/i);
  });
});

// --- CORS headers --------------------------------------------------------

describe('CORS headers', () => {
  it('includes Access-Control-Allow-Origin on GET /', async () => {
    const res = await request(app).get('/').set('Origin', 'http://localhost:3000');
    expect(res.headers['access-control-allow-origin']).toBeDefined();
  });

  it('responds 204 to OPTIONS preflight', async () => {
    const res = await request(app)
      .options('/mcp')
      .set('Origin', 'http://localhost:3000')
      .set('Access-Control-Request-Method', 'POST');
    expect(res.status).toBe(204);
  });
});
