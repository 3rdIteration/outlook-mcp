const { handleListAttachments, handleDownloadAttachment, handleDownloadAttachments } = require('../../email/attachments');
const { callGraphAPI } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');
const fs = require('fs');
const path = require('path');
const os = require('os');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');

describe('handleListAttachments', () => {
  const mockAccessToken = 'dummy_access_token';

  beforeEach(() => {
    callGraphAPI.mockClear();
    ensureAuthenticated.mockClear();
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  test('should require email ID', async () => {
    const result = await handleListAttachments({});
    expect(result.content[0].text).toBe('Email ID is required.');
  });

  test('should list attachments for an email', async () => {
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValue({
      value: [
        {
          id: 'att-1',
          name: 'report.pdf',
          contentType: 'application/pdf',
          size: 2048,
          isInline: false
        },
        {
          id: 'att-2',
          name: 'image.png',
          contentType: 'image/png',
          size: 51200,
          isInline: true
        }
      ]
    });

    const result = await handleListAttachments({ id: 'email-123' });

    expect(ensureAuthenticated).toHaveBeenCalledTimes(1);
    expect(callGraphAPI).toHaveBeenCalledWith(
      mockAccessToken,
      'GET',
      'me/messages/email-123/attachments',
      null,
      { $select: 'id,name,contentType,size,isInline' }
    );
    expect(result.content[0].text).toContain('Found 2 attachment(s)');
    expect(result.content[0].text).toContain('report.pdf');
    expect(result.content[0].text).toContain('image.png');
    expect(result.content[0].text).toContain('(inline)');
    expect(result.content[0].text).toContain('ID: att-1');
  });

  test('should handle email with no attachments', async () => {
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValue({ value: [] });

    const result = await handleListAttachments({ id: 'email-123' });

    expect(result.content[0].text).toBe('This email has no attachments.');
  });

  test('should handle authentication error', async () => {
    ensureAuthenticated.mockRejectedValue(new Error('Authentication required'));

    const result = await handleListAttachments({ id: 'email-123' });

    expect(result.content[0].text).toBe(
      "Authentication required. Please use the 'authenticate' tool first."
    );
  });

  test('should handle API error', async () => {
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockRejectedValue(new Error('API Error'));

    const result = await handleListAttachments({ id: 'email-123' });

    expect(result.content[0].text).toBe('Failed to list attachments: API Error');
  });

  test('should handle invalid mailbox error', async () => {
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockRejectedValue(new Error("The email ID doesn't belong to the targeted mailbox"));

    const result = await handleListAttachments({ id: 'bad-id' });

    expect(result.content[0].text).toContain('email ID seems invalid');
  });
});

describe('handleDownloadAttachment', () => {
  const mockAccessToken = 'dummy_access_token';

  beforeEach(() => {
    callGraphAPI.mockClear();
    ensureAuthenticated.mockClear();
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  test('should require both emailId and attachmentId', async () => {
    let result = await handleDownloadAttachment({});
    expect(result.content[0].text).toBe('Both emailId and attachmentId are required.');

    result = await handleDownloadAttachment({ emailId: 'email-123' });
    expect(result.content[0].text).toBe('Both emailId and attachmentId are required.');

    result = await handleDownloadAttachment({ attachmentId: 'att-1' });
    expect(result.content[0].text).toBe('Both emailId and attachmentId are required.');
  });

  test('should download text file attachment and decode content', async () => {
    const textContent = 'Hello, this is a text file.';
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValue({
      '@odata.type': '#microsoft.graph.fileAttachment',
      id: 'att-1',
      name: 'notes.txt',
      contentType: 'text/plain',
      size: textContent.length,
      contentBytes: Buffer.from(textContent).toString('base64')
    });

    const result = await handleDownloadAttachment({ emailId: 'email-123', attachmentId: 'att-1' });

    expect(callGraphAPI).toHaveBeenCalledWith(
      mockAccessToken,
      'GET',
      'me/messages/email-123/attachments/att-1'
    );
    expect(result.content[0].text).toContain('notes.txt');
    expect(result.content[0].text).toContain('--- Content ---');
    expect(result.content[0].text).toContain(textContent);
  });

  test('should return base64 for binary file attachments', async () => {
    const base64Content = Buffer.from([0x89, 0x50, 0x4e, 0x47]).toString('base64');
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValue({
      '@odata.type': '#microsoft.graph.fileAttachment',
      id: 'att-2',
      name: 'image.png',
      contentType: 'image/png',
      size: 4,
      contentBytes: base64Content
    });

    const result = await handleDownloadAttachment({ emailId: 'email-123', attachmentId: 'att-2' });

    expect(result.content[0].text).toContain('image.png');
    expect(result.content[0].text).toContain('Base64-encoded content');
    expect(result.content[0].text).toContain(base64Content);
  });

  test('should handle item attachment type', async () => {
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValue({
      '@odata.type': '#microsoft.graph.itemAttachment',
      id: 'att-3',
      name: 'Forwarded Meeting',
      contentType: 'application/octet-stream',
      size: 512
    });

    const result = await handleDownloadAttachment({ emailId: 'email-123', attachmentId: 'att-3' });

    expect(result.content[0].text).toContain('Outlook item attachment');
  });

  test('should handle reference attachment type', async () => {
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValue({
      '@odata.type': '#microsoft.graph.referenceAttachment',
      id: 'att-4',
      name: 'SharedDoc.docx',
      contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      size: 0
    });

    const result = await handleDownloadAttachment({ emailId: 'email-123', attachmentId: 'att-4' });

    expect(result.content[0].text).toContain('Reference attachment');
    expect(result.content[0].text).toContain('cloud file link');
  });

  test('should handle authentication error', async () => {
    ensureAuthenticated.mockRejectedValue(new Error('Authentication required'));

    const result = await handleDownloadAttachment({ emailId: 'email-123', attachmentId: 'att-1' });

    expect(result.content[0].text).toBe(
      "Authentication required. Please use the 'authenticate' tool first."
    );
  });

  test('should handle API error', async () => {
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockRejectedValue(new Error('Not Found'));

    const result = await handleDownloadAttachment({ emailId: 'email-123', attachmentId: 'att-1' });

    expect(result.content[0].text).toBe('Failed to download attachment: Not Found');
  });

  test('should handle JSON content type as text', async () => {
    const jsonContent = '{"key": "value"}';
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValue({
      '@odata.type': '#microsoft.graph.fileAttachment',
      id: 'att-5',
      name: 'data.json',
      contentType: 'application/json',
      size: jsonContent.length,
      contentBytes: Buffer.from(jsonContent).toString('base64')
    });

    const result = await handleDownloadAttachment({ emailId: 'email-123', attachmentId: 'att-5' });

    expect(result.content[0].text).toContain('--- Content ---');
    expect(result.content[0].text).toContain(jsonContent);
  });
});

describe('handleDownloadAttachments', () => {
  const mockAccessToken = 'dummy_access_token';
  let tmpDir;

  beforeEach(() => {
    callGraphAPI.mockClear();
    ensureAuthenticated.mockClear();
    jest.spyOn(console, 'error').mockImplementation(() => {});
    // Create a fresh temp directory for each test
    tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'att-test-'));
  });

  afterEach(() => {
    console.error.mockRestore();
    // Clean up temp directory
    if (tmpDir && fs.existsSync(tmpDir)) {
      fs.rmSync(tmpDir, { recursive: true, force: true });
    }
  });

  test('should require emailId', async () => {
    const result = await handleDownloadAttachments({});
    expect(result.content[0].text).toBe('Email ID is required.');
  });

  test('should download all file attachments from an email', async () => {
    const textContent = 'Hello from attachment';
    const binaryContent = Buffer.from([0x89, 0x50]).toString('base64');
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValue({
      value: [
        {
          '@odata.type': '#microsoft.graph.fileAttachment',
          id: 'att-1',
          name: 'notes.txt',
          contentType: 'text/plain',
          size: textContent.length,
          contentBytes: Buffer.from(textContent).toString('base64')
        },
        {
          '@odata.type': '#microsoft.graph.fileAttachment',
          id: 'att-2',
          name: 'image.png',
          contentType: 'image/png',
          size: 2,
          contentBytes: binaryContent
        }
      ]
    });

    const result = await handleDownloadAttachments({ emailId: 'email-123' });

    expect(callGraphAPI).toHaveBeenCalledWith(
      mockAccessToken,
      'GET',
      'me/messages/email-123/attachments'
    );
    expect(result.content[0].text).toContain('Downloaded 2 attachment(s)');
    expect(result.content[0].text).toContain('notes.txt');
    expect(result.content[0].text).toContain('image.png');
    // Attachment content blocks
    expect(result.content.length).toBe(3); // summary + 2 attachments
    expect(result.content[1].text).toContain('--- Attachment: notes.txt');
    expect(result.content[2].text).toContain('--- Attachment: image.png');
  });

  test('should handle email with no attachments', async () => {
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValue({ value: [] });

    const result = await handleDownloadAttachments({ emailId: 'email-123' });

    expect(result.content[0].text).toBe('This email has no attachments.');
  });

  test('should skip non-file attachments and report them', async () => {
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValue({
      value: [
        {
          '@odata.type': '#microsoft.graph.fileAttachment',
          id: 'att-1',
          name: 'doc.txt',
          contentType: 'text/plain',
          size: 5,
          contentBytes: Buffer.from('hello').toString('base64')
        },
        {
          '@odata.type': '#microsoft.graph.itemAttachment',
          id: 'att-2',
          name: 'Meeting Invite',
          contentType: 'application/octet-stream',
          size: 512
        }
      ]
    });

    const result = await handleDownloadAttachments({ emailId: 'email-123' });

    expect(result.content[0].text).toContain('Downloaded 1 attachment(s)');
    expect(result.content[0].text).toContain('1 non-downloadable attachment(s) were skipped');
  });

  test('should report when only non-file attachments exist', async () => {
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValue({
      value: [
        {
          '@odata.type': '#microsoft.graph.itemAttachment',
          id: 'att-1',
          name: 'Forwarded Message',
          contentType: 'application/octet-stream',
          size: 256
        }
      ]
    });

    const result = await handleDownloadAttachments({ emailId: 'email-123' });

    expect(result.content[0].text).toContain('none are downloadable file attachments');
    expect(result.content[0].text).toContain('Outlook item');
  });

  test('should save attachments to disk when saveToPath is provided', async () => {
    const textContent = 'File content here';
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValue({
      value: [
        {
          '@odata.type': '#microsoft.graph.fileAttachment',
          id: 'att-1',
          name: 'saved-file.txt',
          contentType: 'text/plain',
          size: textContent.length,
          contentBytes: Buffer.from(textContent).toString('base64')
        }
      ]
    });

    const result = await handleDownloadAttachments({ emailId: 'email-123', saveToPath: tmpDir });

    expect(result.content[0].text).toContain('Saved 1 attachment(s)');
    expect(result.content[0].text).toContain('saved-file.txt');
    // Only the summary text, no base64 content blocks
    expect(result.content.length).toBe(1);

    // Verify file was written
    const savedPath = path.join(tmpDir, 'saved-file.txt');
    expect(fs.existsSync(savedPath)).toBe(true);
    expect(fs.readFileSync(savedPath, 'utf-8')).toBe(textContent);
  });

  test('should reject non-existent saveToPath', async () => {
    const result = await handleDownloadAttachments({
      emailId: 'email-123',
      saveToPath: '/nonexistent/directory/path'
    });

    expect(result.content[0].text).toContain('does not exist');
  });

  test('should reject saveToPath that is a file, not a directory', async () => {
    const filePath = path.join(tmpDir, 'afile.txt');
    fs.writeFileSync(filePath, 'data');

    const result = await handleDownloadAttachments({
      emailId: 'email-123',
      saveToPath: filePath
    });

    expect(result.content[0].text).toContain('is not a directory');
  });

  test('should sanitize filenames to prevent path traversal', async () => {
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValue({
      value: [
        {
          '@odata.type': '#microsoft.graph.fileAttachment',
          id: 'att-1',
          name: '../../../etc/evil.txt',
          contentType: 'text/plain',
          size: 4,
          contentBytes: Buffer.from('test').toString('base64')
        }
      ]
    });

    const result = await handleDownloadAttachments({ emailId: 'email-123', saveToPath: tmpDir });

    expect(result.content[0].text).toContain('Saved 1 attachment(s)');
    // File should be saved with just the basename, not the traversal path
    const safePath = path.join(tmpDir, 'evil.txt');
    expect(fs.existsSync(safePath)).toBe(true);
    // The traversal path should NOT exist
    expect(fs.existsSync(path.join(tmpDir, '..', '..', '..', 'etc', 'evil.txt'))).toBe(false);
  });

  test('should handle authentication error', async () => {
    ensureAuthenticated.mockRejectedValue(new Error('Authentication required'));

    const result = await handleDownloadAttachments({ emailId: 'email-123' });

    expect(result.content[0].text).toBe(
      "Authentication required. Please use the 'authenticate' tool first."
    );
  });

  test('should handle API error', async () => {
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockRejectedValue(new Error('Server Error'));

    const result = await handleDownloadAttachments({ emailId: 'email-123' });

    expect(result.content[0].text).toBe('Failed to download attachments: Server Error');
  });
});
