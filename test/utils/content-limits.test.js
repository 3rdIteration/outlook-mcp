/**
 * Tests for configurable content length limits.
 *
 * Verifies that subject, sender/recipient, body preview, and full email body
 * are truncated to the configured limits — protecting the LLM context window
 * from overflow via malicious or abnormally large content.
 */

const config = require('../../config');
const { sanitizeMetadata } = require('../../utils/metadata-sanitizer');
const { processHtmlEmail } = require('../../utils/html-sanitizer');

// --- config defaults ---

describe('content length limit config defaults', () => {
  test('MAX_SUBJECT_LENGTH defaults to 500', () => {
    expect(config.MAX_SUBJECT_LENGTH).toBe(500);
  });

  test('MAX_SENDER_LENGTH defaults to 200', () => {
    expect(config.MAX_SENDER_LENGTH).toBe(200);
  });

  test('MAX_BODY_PREVIEW_LENGTH defaults to 500', () => {
    expect(config.MAX_BODY_PREVIEW_LENGTH).toBe(500);
  });

  test('MAX_BODY_LENGTH defaults to 50000', () => {
    expect(config.MAX_BODY_LENGTH).toBe(50000);
  });
});

// --- sanitizeMetadata with per-field limits ---

describe('sanitizeMetadata respects per-field limits', () => {
  test('subject is truncated to MAX_SUBJECT_LENGTH', () => {
    const longSubject = 'A'.repeat(1000);
    const result = sanitizeMetadata(longSubject, config.MAX_SUBJECT_LENGTH);
    // +1 for the ellipsis character
    expect(result.length).toBeLessThanOrEqual(config.MAX_SUBJECT_LENGTH + 1);
    expect(result).toContain('…');
  });

  test('sender name is truncated to MAX_SENDER_LENGTH', () => {
    const longName = 'B'.repeat(500);
    const result = sanitizeMetadata(longName, config.MAX_SENDER_LENGTH);
    expect(result.length).toBeLessThanOrEqual(config.MAX_SENDER_LENGTH + 1);
    expect(result).toContain('…');
  });

  test('body preview is truncated to MAX_BODY_PREVIEW_LENGTH', () => {
    const longPreview = 'C'.repeat(1000);
    const result = sanitizeMetadata(longPreview, config.MAX_BODY_PREVIEW_LENGTH);
    expect(result.length).toBeLessThanOrEqual(config.MAX_BODY_PREVIEW_LENGTH + 1);
    expect(result).toContain('…');
  });

  test('short strings are not truncated', () => {
    const shortSubject = 'Hello World';
    const result = sanitizeMetadata(shortSubject, config.MAX_SUBJECT_LENGTH);
    expect(result).toBe('Hello World');
    expect(result).not.toContain('…');
  });

  test('string at exactly the limit is not truncated', () => {
    const exactSubject = 'X'.repeat(config.MAX_SUBJECT_LENGTH);
    const result = sanitizeMetadata(exactSubject, config.MAX_SUBJECT_LENGTH);
    expect(result).toBe(exactSubject);
    expect(result).not.toContain('…');
  });

  test('custom limit can be passed directly', () => {
    const input = 'A'.repeat(100);
    const result = sanitizeMetadata(input, 50);
    expect(result.length).toBeLessThanOrEqual(51); // 50 + ellipsis
    expect(result).toContain('…');
  });
});

// --- processHtmlEmail body truncation ---

describe('processHtmlEmail maxLength option', () => {
  test('long body is truncated when maxLength is set', () => {
    const longBody = '<p>' + 'D'.repeat(100000) + '</p>';
    const result = processHtmlEmail(longBody, {
      addBoundary: false,
      maxLength: 1000
    });
    // Truncated text + "…[truncated]" marker
    expect(result.length).toBeLessThanOrEqual(1000 + 20);
    expect(result).toContain('…[truncated]');
  });

  test('short body is not truncated', () => {
    const shortBody = '<p>Hello World</p>';
    const result = processHtmlEmail(shortBody, {
      addBoundary: false,
      maxLength: 1000
    });
    expect(result).not.toContain('…[truncated]');
    expect(result).toContain('Hello World');
  });

  test('body at exactly the limit is not truncated', () => {
    // Create content that after HTML stripping is exactly at the limit
    const text = 'E'.repeat(100);
    const html = `<p>${text}</p>`;
    const result = processHtmlEmail(html, {
      addBoundary: false,
      maxLength: 100
    });
    expect(result).not.toContain('…[truncated]');
  });

  test('maxLength 0 means no truncation', () => {
    const longBody = '<p>' + 'F'.repeat(100000) + '</p>';
    const result = processHtmlEmail(longBody, {
      addBoundary: false,
      maxLength: 0
    });
    expect(result).not.toContain('…[truncated]');
    expect(result.length).toBeGreaterThanOrEqual(100000);
  });

  test('default (no maxLength) means no truncation', () => {
    const longBody = '<p>' + 'G'.repeat(100000) + '</p>';
    const result = processHtmlEmail(longBody, {
      addBoundary: false
    });
    expect(result).not.toContain('…[truncated]');
  });

  test('truncation happens before boundary wrapping', () => {
    const longBody = '<p>' + 'H'.repeat(10000) + '</p>';
    const result = processHtmlEmail(longBody, {
      addBoundary: true,
      maxLength: 500,
      metadata: { from: 'test@example.com', subject: 'Test' }
    });
    // Should contain the truncation marker inside the boundary
    expect(result).toContain('…[truncated]');
    // Should still have boundary markers
    expect(result).toContain('EMAIL CONTENT START');
    expect(result).toContain('EMAIL CONTENT END');
  });
});

// --- env var override simulation ---

describe('content limits are overridable via env vars', () => {
  const originalEnv = { ...process.env };

  afterEach(() => {
    // Restore original env
    process.env = { ...originalEnv };
  });

  test('MCP_MAX_SUBJECT_LENGTH env var overrides default', () => {
    process.env.MCP_MAX_SUBJECT_LENGTH = '100';
    // Re-require config to pick up new env
    jest.resetModules();
    const freshConfig = require('../../config');
    expect(freshConfig.MAX_SUBJECT_LENGTH).toBe(100);
  });

  test('MCP_MAX_SENDER_LENGTH env var overrides default', () => {
    process.env.MCP_MAX_SENDER_LENGTH = '50';
    jest.resetModules();
    const freshConfig = require('../../config');
    expect(freshConfig.MAX_SENDER_LENGTH).toBe(50);
  });

  test('MCP_MAX_BODY_PREVIEW_LENGTH env var overrides default', () => {
    process.env.MCP_MAX_BODY_PREVIEW_LENGTH = '250';
    jest.resetModules();
    const freshConfig = require('../../config');
    expect(freshConfig.MAX_BODY_PREVIEW_LENGTH).toBe(250);
  });

  test('MCP_MAX_BODY_LENGTH env var overrides default', () => {
    process.env.MCP_MAX_BODY_LENGTH = '10000';
    jest.resetModules();
    const freshConfig = require('../../config');
    expect(freshConfig.MAX_BODY_LENGTH).toBe(10000);
  });
});
