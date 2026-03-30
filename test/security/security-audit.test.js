/**
 * Security audit tests — validates security hardening across the codebase.
 *
 * These tests verify:
 * 1. Token file permissions (0o600 on non-Windows)
 * 2. OAuth state uses crypto.randomBytes (not Date.now)
 * 3. Debug logging does not leak token contents or file paths
 * 4. Error messages do not leak raw API responses
 * 5. OData escaping preserves legitimate characters
 * 6. HTML escaping in auth error responses
 * 7. Unnecessary API scopes removed
 */

const fs = require('fs');
const path = require('path');
const os = require('os');

describe('Security Audit', () => {

  // ---------- 1. Token file permissions ----------

  describe('Token file permissions (token-manager.js)', () => {
    it('should set 0o600 permissions after saveTokenCache on non-Windows', () => {
      if (process.platform === 'win32') return; // skip on Windows

      const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'sec-audit-'));
      const tmpTokenPath = path.join(tmpDir, '.outlook-mcp-tokens.json');

      try {
        // Require with isolated module cache
        jest.resetModules();
        jest.doMock('../../config', () => ({
          AUTH_CONFIG: {
            tokenStorePath: tmpTokenPath,
            clientId: 'test-id',
            clientSecret: 'test-secret',
          },
          USE_TEST_MODE: false,
        }));
        const tokenManager = require('../../auth/token-manager');

        const tokens = {
          access_token: 'tok_test',
          refresh_token: 'ref_test',
          expires_at: Date.now() + 3600000,
        };

        tokenManager.saveTokenCache(tokens);

        const stats = fs.statSync(tmpTokenPath);
        const perms = stats.mode & 0o777;
        expect(perms).toBe(0o600);
      } finally {
        try { fs.unlinkSync(tmpTokenPath); } catch (_) {}
        try { fs.rmdirSync(tmpDir); } catch (_) {}
        jest.resetModules();
      }
    });

    it('should set 0o600 permissions after saveFlowTokens on non-Windows', () => {
      if (process.platform === 'win32') return;

      const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'sec-audit-'));
      const tmpTokenPath = path.join(tmpDir, '.outlook-mcp-tokens.json');

      try {
        jest.resetModules();
        jest.doMock('../../config', () => ({
          AUTH_CONFIG: {
            tokenStorePath: tmpTokenPath,
            clientId: 'test-id',
            clientSecret: 'test-secret',
          },
          USE_TEST_MODE: false,
        }));
        const tokenManager = require('../../auth/token-manager');

        // First create a base token file
        tokenManager.saveTokenCache({
          access_token: 'tok_test',
          refresh_token: 'ref_test',
          expires_at: Date.now() + 3600000,
        });

        // Now save flow tokens
        tokenManager.saveFlowTokens({
          access_token: 'flow_tok',
          refresh_token: 'flow_ref',
          expires_in: 3600,
        });

        const stats = fs.statSync(tmpTokenPath);
        const perms = stats.mode & 0o777;
        expect(perms).toBe(0o600);
      } finally {
        try { fs.unlinkSync(tmpTokenPath); } catch (_) {}
        try { fs.rmdirSync(tmpDir); } catch (_) {}
        jest.resetModules();
      }
    });
  });

  // ---------- 2. Debug logging removed ----------

  describe('Debug logging (token-manager.js)', () => {
    it('should not contain sensitive DEBUG log statements', () => {
      const src = fs.readFileSync(
        path.join(__dirname, '../../auth/token-manager.js'),
        'utf8'
      );
      // Must not log token file contents
      expect(src).not.toContain('Token file first 200 characters');
      expect(src).not.toContain('Token file contents');
      // Must not log HOME directory
      expect(src).not.toContain('HOME directory');
      // Must not log parsed token keys
      expect(src).not.toContain('Parsed tokens keys');
      // Must not contain the generic [DEBUG] prefix with sensitive data
      expect(src).not.toMatch(/console\.error\(.*\[DEBUG\].*token/i);
    });
  });

  describe('Debug logging (auth/tools.js)', () => {
    it('should not contain sensitive CHECK-AUTH-STATUS log statements', () => {
      const src = fs.readFileSync(
        path.join(__dirname, '../../auth/tools.js'),
        'utf8'
      );
      expect(src).not.toContain('Token expires at');
      expect(src).not.toContain('[CHECK-AUTH-STATUS]');
    });
  });

  // ---------- 3. Error message sanitization ----------

  describe('Error messages (graph-api.js)', () => {
    it('should not include raw API response data in error messages', () => {
      const src = fs.readFileSync(
        path.join(__dirname, '../../utils/graph-api.js'),
        'utf8'
      );
      // The error messages should NOT append ${responseData}
      expect(src).not.toMatch(/reject\(new Error\(`.*\$\{responseData\}`\)\)/);
    });
  });

  describe('Error messages (token-storage.js)', () => {
    it('should not include raw response data in token exchange errors', () => {
      const src = fs.readFileSync(
        path.join(__dirname, '../../auth/token-storage.js'),
        'utf8'
      );
      // Must not log full responseBody objects
      expect(src).not.toMatch(/console\.error\(.*responseBody\)/);
      // Must not include "Raw data:" in error messages
      expect(src).not.toContain('Raw data:');
    });
  });

  // ---------- 4. OAuth state parameter ----------

  describe('OAuth state parameter (outlook-auth-server.js)', () => {
    it('should use crypto.randomBytes for state, not Date.now()', () => {
      const src = fs.readFileSync(
        path.join(__dirname, '../../outlook-auth-server.js'),
        'utf8'
      );
      // Must use crypto.randomBytes for state
      expect(src).toContain("crypto.randomBytes");
      // Must NOT use Date.now() for state
      expect(src).not.toMatch(/state:\s*Date\.now\(\)/);
    });

    it('should validate state parameter in callback', () => {
      const src = fs.readFileSync(
        path.join(__dirname, '../../outlook-auth-server.js'),
        'utf8'
      );
      // Must check state in callback
      expect(src).toContain('pendingStates');
    });
  });

  // ---------- 5. HTML escaping in error responses ----------

  describe('HTML escaping (outlook-auth-server.js)', () => {
    it('should HTML-escape error parameters in responses', () => {
      const src = fs.readFileSync(
        path.join(__dirname, '../../outlook-auth-server.js'),
        'utf8'
      );
      // Must use escapeHtml for query.error and query.error_description
      expect(src).toContain('escapeHtml(query.error)');
      expect(src).toContain('escapeHtml(query.error_description');
    });

    it('should not embed raw error.message in HTML response', () => {
      const src = fs.readFileSync(
        path.join(__dirname, '../../outlook-auth-server.js'),
        'utf8'
      );
      // The token exchange error HTML should not contain ${error.message}
      const errorInHtml = src.match(/<p>\$\{error\.message\}<\/p>/g);
      expect(errorInHtml).toBeNull();
    });
  });

  // ---------- 6. OData escaping ----------

  describe('OData escaping (odata-helpers.js)', () => {
    const { escapeODataString } = require('../../utils/odata-helpers');

    it('should double single quotes', () => {
      expect(escapeODataString("it's")).toBe("it''s");
    });

    it('should preserve hyphens, colons, and slashes', () => {
      expect(escapeODataString('2024-01-15')).toBe('2024-01-15');
      expect(escapeODataString('10:30:00')).toBe('10:30:00');
      expect(escapeODataString('path/to/file')).toBe('path/to/file');
    });

    it('should preserve @, #, $ symbols', () => {
      expect(escapeODataString('user@example.com')).toBe('user@example.com');
    });

    it('should remove control characters', () => {
      expect(escapeODataString('hello\x00world')).toBe('helloworld');
      expect(escapeODataString('test\x1Fvalue')).toBe('testvalue');
    });

    it('should return empty/falsy values as-is', () => {
      expect(escapeODataString('')).toBe('');
      expect(escapeODataString(null)).toBe(null);
      expect(escapeODataString(undefined)).toBe(undefined);
    });
  });

  // ---------- 7. Unnecessary scopes removed ----------

  describe('API scopes (outlook-auth-server.js)', () => {
    it('should not request Contacts.Read scope', () => {
      const src = fs.readFileSync(
        path.join(__dirname, '../../outlook-auth-server.js'),
        'utf8'
      );
      expect(src).not.toContain("'Contacts.Read'");
    });
  });

  // ---------- 8. Token file permissions in outlook-auth-server.js ----------

  describe('Token file permissions (outlook-auth-server.js)', () => {
    it('should call chmodSync after writing token file', () => {
      const src = fs.readFileSync(
        path.join(__dirname, '../../outlook-auth-server.js'),
        'utf8'
      );
      expect(src).toContain('chmodSync');
      expect(src).toContain('0o600');
    });
  });

  // ---------- 9. npm audit vulnerabilities ----------

  describe('Dependency security', () => {
    it('should have @modelcontextprotocol/inspector at a non-vulnerable version', () => {
      const pkg = JSON.parse(fs.readFileSync(
        path.join(__dirname, '../../package.json'),
        'utf8'
      ));
      const inspectorVersion = pkg.devDependencies['@modelcontextprotocol/inspector'];
      expect(inspectorVersion).toBeDefined();
      // The vulnerable range is <=0.16.5; we need > 0.16.5
      expect(inspectorVersion).not.toMatch(/^\^?0\.(1[0-6]|[0-9])\./);
    });
  });
});
