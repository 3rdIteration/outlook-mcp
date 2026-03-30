describe('handleReadEmail', () => {
  const mockEmail = {
    id: 'email-123',
    subject: 'Quarterly Report',
    from: {
      emailAddress: {
        name: 'Alice Example',
        address: 'alice@example.com'
      }
    },
    toRecipients: [],
    ccRecipients: [],
    bccRecipients: [],
    receivedDateTime: '2024-01-15T10:30:00Z',
    body: {
      contentType: 'html',
      content: '<div>Hello team</div><script>alert("x")</script>'
    },
    bodyPreview: 'Hello team',
    importance: 'normal',
    hasAttachments: false
  };

  async function loadHandler({ allowUnsafeRawHtmlDebug = false } = {}) {
    jest.resetModules();
    process.env.MCP_ALLOW_UNSAFE_RAW_HTML_DEBUG = allowUnsafeRawHtmlDebug ? 'true' : 'false';

    const callGraphAPI = jest.fn().mockResolvedValue(mockEmail);
    const ensureAuthenticated = jest.fn().mockResolvedValue('access-token');

    jest.doMock('../../utils/graph-api', () => ({
      callGraphAPI
    }));
    jest.doMock('../../auth', () => ({
      ensureAuthenticated
    }));

    const handleReadEmail = require('../../email/read');

    return {
      handleReadEmail,
      callGraphAPI,
      ensureAuthenticated
    };
  }

  afterEach(() => {
    delete process.env.MCP_ALLOW_UNSAFE_RAW_HTML_DEBUG;
    jest.resetModules();
    jest.clearAllMocks();
  });

  it('omits raw HTML by default even when requested', async () => {
    const { handleReadEmail } = await loadHandler();

    const result = await handleReadEmail({
      emailId: 'email-123',
      includeRawHtml: true
    });

    expect(result.content[0].text).toContain('Raw HTML omitted for security');
    expect(result.content[0].text).not.toContain('<script>alert("x")</script>');
    expect(result.content[0].text).not.toContain('--- RAW HTML (UNSAFE - FOR DEBUGGING ONLY) ---');
  });

  it('includes raw HTML only when explicitly enabled', async () => {
    const { handleReadEmail } = await loadHandler({ allowUnsafeRawHtmlDebug: true });

    const result = await handleReadEmail({
      emailId: 'email-123',
      includeRawHtml: true
    });

    expect(result.content[0].text).toContain('--- RAW HTML (UNSAFE - FOR DEBUGGING ONLY) ---');
    expect(result.content[0].text).toContain('<script>alert("x")</script>');
  });
});
