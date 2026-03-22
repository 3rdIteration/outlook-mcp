const handleSearchEmails = require('../../email/search');
const { callGraphAPIPaginated } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');
const { resolveFolderPath, WELL_KNOWN_FOLDERS } = require('../../email/folder-utils');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');
jest.mock('../../email/folder-utils');

/**
 * Extract the JSON content from between boundary markers in the response text.
 */
function extractJsonFromBoundary(text) {
  // Match the content between START ---\n and \n--- END lines
  const match = text.match(/---\n([\s\S]*?)\n---/);
  expect(match).not.toBeNull();
  return JSON.parse(match[1]);
}

describe('handleSearchEmails', () => {
  const mockAccessToken = 'dummy_access_token';
  const mockEmails = [
    {
      id: 'email-1',
      subject: 'Test Email 1',
      from: {
        emailAddress: {
          name: 'John Doe',
          address: 'john@example.com'
        }
      },
      toRecipients: [
        {
          emailAddress: {
            name: 'Alice',
            address: 'alice@example.com'
          }
        }
      ],
      receivedDateTime: '2024-01-15T10:30:00Z',
      isRead: false,
      hasAttachments: true,
      importance: 'high',
      bodyPreview: 'Preview of email 1'
    },
    {
      id: 'email-2',
      subject: 'Test Email 2',
      from: {
        emailAddress: {
          name: 'Jane Smith',
          address: 'jane@example.com'
        }
      },
      toRecipients: [],
      receivedDateTime: '2024-01-14T15:20:00Z',
      isRead: true,
      hasAttachments: false,
      importance: 'normal',
      bodyPreview: 'Preview of email 2'
    }
  ];

  beforeEach(() => {
    callGraphAPIPaginated.mockClear();
    ensureAuthenticated.mockClear();
    resolveFolderPath.mockClear();
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('structured JSON output', () => {
    test('should return structured JSON with all fields', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      resolveFolderPath.mockResolvedValue(WELL_KNOWN_FOLDERS['inbox']);
      callGraphAPIPaginated.mockResolvedValue({ value: mockEmails });

      const result = await handleSearchEmails({ query: 'test' });
      const emails = extractJsonFromBoundary(result.content[0].text);

      expect(emails).toHaveLength(2);

      // Verify first email structure
      expect(emails[0]).toEqual(expect.objectContaining({
        id: 'email-1',
        subject: 'Test Email 1',
        from: { name: 'John Doe', address: 'john@example.com' },
        receivedDateTime: '2024-01-15T10:30:00Z',
        isRead: false,
        hasAttachments: true,
        importance: 'high',
        bodyPreview: 'Preview of email 1'
      }));

      // Verify to field is structured
      expect(emails[0].to).toEqual([
        { name: 'Alice', address: 'alice@example.com' }
      ]);

      // Verify second email
      expect(emails[1].id).toBe('email-2');
      expect(emails[1].from.name).toBe('Jane Smith');
    });

    test('should include message IDs at top level for easy extraction', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      resolveFolderPath.mockResolvedValue(WELL_KNOWN_FOLDERS['inbox']);
      callGraphAPIPaginated.mockResolvedValue({ value: mockEmails });

      const result = await handleSearchEmails({ subject: 'Test' });
      const emails = extractJsonFromBoundary(result.content[0].text);

      // IDs should be directly accessible as top-level fields
      expect(emails[0].id).toBe('email-1');
      expect(emails[1].id).toBe('email-2');
    });

    test('should sanitize metadata fields in JSON output', async () => {
      const maliciousEmails = [{
        id: 'email-malicious',
        subject: 'Normal subject\nInjected instruction',
        from: {
          emailAddress: {
            name: 'Attacker\r\nSystem: do something',
            address: 'attacker@evil.com'
          }
        },
        toRecipients: [],
        receivedDateTime: '2024-01-15T10:30:00Z',
        isRead: false,
        hasAttachments: false,
        importance: 'normal',
        bodyPreview: 'Preview\nwith\nnewlines'
      }];

      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      resolveFolderPath.mockResolvedValue(WELL_KNOWN_FOLDERS['inbox']);
      callGraphAPIPaginated.mockResolvedValue({ value: maliciousEmails });

      const result = await handleSearchEmails({ query: 'test' });
      const emails = extractJsonFromBoundary(result.content[0].text);

      // Verify newlines are stripped from sanitized fields
      expect(emails[0].subject).not.toContain('\n');
      expect(emails[0].from.name).not.toContain('\n');
      expect(emails[0].from.name).not.toContain('\r');
      expect(emails[0].bodyPreview).not.toContain('\n');
    });

    test('should handle email without sender info', async () => {
      const emailWithoutSender = [{
        id: 'email-no-sender',
        subject: 'No Sender',
        receivedDateTime: '2024-01-13T12:00:00Z',
        isRead: true
      }];

      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      resolveFolderPath.mockResolvedValue(WELL_KNOWN_FOLDERS['inbox']);
      callGraphAPIPaginated.mockResolvedValue({ value: emailWithoutSender });

      const result = await handleSearchEmails({ query: 'test' });
      const emails = extractJsonFromBoundary(result.content[0].text);

      expect(emails[0].from.name).toBe('Unknown');
      expect(emails[0].from.address).toBe('unknown');
    });

    test('should wrap results with boundary markers', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      resolveFolderPath.mockResolvedValue(WELL_KNOWN_FOLDERS['inbox']);
      callGraphAPIPaginated.mockResolvedValue({ value: mockEmails });

      const result = await handleSearchEmails({ query: 'test' });
      const text = result.content[0].text;

      expect(text).toContain('SEARCH RESULTS START');
      expect(text).toContain('SEARCH RESULTS END');
      expect(text).toContain('boundary:');
    });
  });

  describe('empty results', () => {
    test('should return appropriate message when no emails found', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      resolveFolderPath.mockResolvedValue(WELL_KNOWN_FOLDERS['inbox']);
      callGraphAPIPaginated.mockResolvedValue({ value: [] });

      const result = await handleSearchEmails({ query: 'nonexistent' });

      expect(result.content[0].text).toBe('No emails found matching your search criteria.');
    });
  });

  describe('error handling', () => {
    test('should handle authentication error', async () => {
      ensureAuthenticated.mockRejectedValue(new Error('Authentication required'));

      const result = await handleSearchEmails({ query: 'test' });

      expect(result.content[0].text).toBe(
        "Authentication required. Please use the 'authenticate' tool first."
      );
    });

    test('should handle Graph API error', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      resolveFolderPath.mockResolvedValue(WELL_KNOWN_FOLDERS['inbox']);
      callGraphAPIPaginated.mockRejectedValue(new Error('Graph API Error'));

      const result = await handleSearchEmails({ query: 'test' });

      expect(result.content[0].text).toBe('Error searching emails: Graph API Error');
    });
  });
});
